/**
	Opens a document's bytes into the model a view shows, and writes edited
	spreadsheets back. Runs in a Web Worker (source/core/Worker.js) so that
	unzipping and parsing a large file never blocks the page; when a worker
	cannot be started the page runs the same engine itself.

		var engine = new DocEngine({JSZip: JSZip, XLSX: XLSX});
		engine.open({name: "a.odt", buffer: arrayBuffer}).then(function(result) {
			result.kind      // "odt", "xlsx", ...
			result.family    // "text", "sheet", "slides"
			result.model     // DocOdf / DocPptx / DocSheets / text model
			result.images    // {zipPath: {bytes: Uint8Array, mime}}
			result.meta      // title, author, ...
		});
		engine.save({edits: [{sheet, r, c, text}], format: "xlsx"}).then(function(bytes) {...});

	Libraries are passed in (or found as globals), so node tests can use the
	bundled files directly.
*/
(function(root) {
	function need(inName, inPath) {
		if (root[inName]) {
			return root[inName];
		}
		return typeof require === "function" ? require(inPath) : null;
	}

	var FT = need("FileTypes", "./FileTypes.js");
	var Odf = need("DocOdf", "./Odf.js");
	var Pptx = need("DocPptx", "./Pptx.js");
	var Text = need("DocText", "./Text.js");
	var Sheets = need("DocSheets", "./Sheets.js");
	var Docx = need("DocDocx", "./Docx.js");
	var Stream = need("DocSheetStream", "./SheetStream.js");

	/**
		Workbooks whose sheet XML is larger than this (uncompressed) are
		streamed (SheetStream.js, read-only) instead of read whole by SheetJS:
		SheetJS needs about 570 bytes of memory per cell, and the inflated
		XML as one string on top. Smaller ones get SheetJS, with editing.
	*/
	var FULL_SHEET_XML_LIMIT = 4 * 1024 * 1024;
	//* At most this many cells of a streamed workbook are kept (all sheets).
	var STREAM_MAX_CELLS = 300000;
	//* .xls files larger than this are read only up to XLS_MAX_ROWS rows per sheet.
	var XLS_FULL_LIMIT = 3 * 1024 * 1024;
	var XLS_MAX_ROWS = 20000;
	//* Rows per progress message while streaming.
	var STREAM_BATCH_ROWS = 1500;

	/**
		Word documents up to this much text XML are laid out by docx-preview
		(closest to Word); larger ones by the faster reader in Docx.js, which
		the page renders a slice at a time. docx-preview takes about 14 s per
		MB of document.xml on an HP TouchPad.
	*/
	var DOCX_PREVIEW_LIMIT = 250 * 1024;

	var imageTypes = {png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp",
		svg: "image/svg+xml", webp: "image/webp", tif: "image/tiff", tiff: "image/tiff"};

	//* Pictures bigger than this are not handed to the page (a broken picture is shown).
	var MAX_IMAGE_BYTES = 12 * 1024 * 1024;

	function DocEngine(inLibs) {
		this.libs = inLibs || {};
		this.workbook = null;
		this.csvRows = null;
	}

	//* A library by its global name; the spreadsheet reader may be loaded on first use (libs.loadXLSX).
	DocEngine.prototype.lib = function(inName) {
		var lib = this.libs[inName] || root[inName] || null;
		if (!lib && inName === "XLSX" && this.libs.loadXLSX) {
			lib = this.libs.XLSX = this.libs.loadXLSX();
		}
		return lib;
	};

	/**
		inRequest: {name, buffer (ArrayBuffer)}. Resolves with the result
		described above, or rejects with an Error whose message is shown.
	*/
	DocEngine.prototype.open = function(inRequest) {
		var self = this;
		var bytes = new Uint8Array(inRequest.buffer);
		var verdict = FT.resolve(inRequest.name, bytes.subarray(0, 4096));
		this.workbook = null;
		this.csvRows = null;
		this.kind = null;
		if (verdict.kind === "zip" || ["docx", "odt", "xlsx", "ods", "pptx", "odp"].indexOf(verdict.kind) >= 0 &&
				FT.sniff(bytes.subarray(0, 8)) === "zip") {
			return this.openZip(inRequest.name, inRequest.buffer, bytes, inRequest);
		}
		return new Promise(function(resolve) {
			resolve(self.openFlat(verdict, inRequest.name, bytes, inRequest));
		});
	};

	//* The rows of a long CSV after the first part, a batch per turn of the event loop.
	DocEngine.prototype.sendRows = function(inRows, inFrom, inTruncated, inProgress) {
		var progress = inProgress || function() {};
		var at = inFrom;
		var cols = 0;
		for (var c = 0; c < inRows.length; c++) {
			cols = Math.max(cols, inRows[c].length);
		}
		var step = function() {
			var batch = [];
			var end = Math.min(inRows.length, at + STREAM_BATCH_ROWS);
			for (; at < end; at++) {
				var texts = inRows[at], kinds = [];
				for (var k = 0; k < texts.length; k++) {
					kinds[k] = Sheets.isNumeric(texts[k]) ? "n" : "s";
				}
				batch.push([at, texts, kinds]);
			}
			progress({sheet: 0, rows: batch});
			if (at < inRows.length) {
				setTimeout(step, 0);
			} else {
				var widths = [];
				for (var w = 0; w < Math.min(cols, 1024); w++) {
					widths.push(100);
				}
				progress({sheet: 0, done: true, info: {merges: [], colWidths: widths, rowHeights: {}, truncated: inTruncated}});
				progress({finished: true});
			}
		};
		setTimeout(step, 0);
	};

	DocEngine.prototype.openFlat = function(inVerdict, inName, inBytes, inRequest) {
		var kind = inVerdict.kind;
		this.kind = kind;
		if (kind === "txt" || kind === "md") {
			var decoded = Text.decode(inBytes);
			return {kind: kind, family: "text", meta: {encoding: decoded.encoding},
				model: kind === "md" ? {html: Text.markdown(decoded.text)} : {text: decoded.text}, images: {}};
		}
		if (kind === "csv") {
			var d = Text.decode(inBytes);
			var parsed = Text.csv(d.text, /\.tsv$/i.test(inName) ? "\t" : null, STREAM_MAX_CELLS);
			var rows = parsed.rows;
			this.csvRows = parsed.truncated ? null : rows;
			this.csvDelimiter = parsed.delimiter;
			var name = FT.extensionOf(inName) === "tsv" ? "TSV" : "CSV";
			var meta = {encoding: d.encoding, delimiter: parsed.delimiter === "\t" ? "Tab" : parsed.delimiter};
			// A long file goes to the page in parts, like a streamed workbook, so the page never
			// has to take a few hundred thousand cells in one message.
			var firstPart = rows.length > 2 * STREAM_BATCH_ROWS ? 500 : rows.length;
			var model = Sheets.fromRows(rows.slice(0, firstPart), name);
			if (parsed.truncated) {
				// Cut short: a saved copy would lose the rest, so this one is read-only.
				model.streaming = true;
			}
			if (firstPart < rows.length) {
				this.sendRows(rows, firstPart, parsed.truncated, inRequest && inRequest.onProgress);
			} else {
				model.sheets[0].truncated = parsed.truncated;
			}
			return {kind: "csv", family: "sheet", model: model, images: {}, meta: meta, streamed: parsed.truncated};
		}
		if (kind === "xls") {
			// The binary format cannot be streamed: a large one is read only so far.
			return this.openWorkbook("xls", inBytes, inBytes.length > XLS_FULL_LIMIT ? XLS_MAX_ROWS : 0);
		}
		if (/^fod[tsp]$/.test(FT.extensionOf(inName))) {
			return this.openFlatOdf(kind, inBytes);
		}
		throw this.unsupported(kind, inVerdict.label);
	};

	//* A flat OpenDocument file (.fodt, .fods, .fodp): styles and content in one XML document.
	DocEngine.prototype.openFlatOdf = function(inKind, inBytes) {
		if (inKind === "ods") {
			return this.openWorkbook("ods", inBytes);
		}
		var xml = Text.decode(inBytes).text;
		var files = {"content.xml": xml, "styles.xml": xml, "meta.xml": xml};
		var model = inKind === "odp" ? Odf.presentation(files) : Odf.text(files);
		// Pictures in a flat file are inline (office:binary-data); they are not shown.
		return {kind: inKind, family: inKind === "odp" ? "slides" : "text", model: model, images: {}, meta: model.meta || {}};
	};

	DocEngine.prototype.unsupported = function(inKind, inLabel) {
		var e = new Error(inKind === "doc" || inKind === "ppt" ?
			inLabel + " files (the binary format from before 2007) cannot be opened. Save it as ." + inKind + "x or as OpenDocument." :
			inKind === "pdf" ? "This is a PDF document, and no app to show PDF documents answered." :
			inKind === "rtf" ? "Rich text (.rtf) files are not supported yet." :
			"This file is not a document Documents can open.");
		e.code = "unsupported";
		e.kind = inKind;
		return e;
	};

	DocEngine.prototype.openWorkbook = function(inKind, inBytes, inMaxRows) {
		var XLSX = this.lib("XLSX");
		if (!XLSX) {
			throw new Error("The spreadsheet reader did not load.");
		}
		var opts = {type: "array", dense: true, cellFormula: true, cellStyles: false, cellDates: false,
			cellNF: false, sheetStubs: false, WTF: false};
		if (inMaxRows) {
			opts.sheetRows = inMaxRows;
		}
		var wb = XLSX.read(inBytes, opts);
		this.workbook = wb;
		this.kind = inKind;
		var meta = {};
		var props = wb.Props || {};
		if (props.Title) {
			meta.title = String(props.Title);
		}
		if (props.Author) {
			meta.author = String(props.Author);
		}
		if (props.ModifiedDate) {
			meta.modified = String(props.ModifiedDate);
		}
		var model = Sheets.fromWorkbook(XLSX, wb);
		if (inMaxRows) {
			// Cut at inMaxRows: read-only, as a saved copy would lose the rest.
			var cut = false;
			wb.SheetNames.forEach(function(n, i) {
				var ws = wb.Sheets[n];
				if (ws && ws["!fullref"] && ws["!fullref"] !== ws["!ref"] && model.sheets[i]) {
					model.sheets[i].truncated = true;
					cut = true;
				}
			});
			if (cut) {
				model.streaming = true;
				this.workbook = null;
				return {kind: inKind, family: "sheet", model: model, images: {}, meta: meta, streamed: true};
			}
		}
		return {kind: inKind, family: "sheet", model: model, images: {}, meta: meta};
	};

	DocEngine.prototype.openZip = function(inName, inBuffer, inBytes, inRequest) {
		var self = this;
		var JSZip = this.lib("JSZip");
		if (!JSZip) {
			return Promise.reject(new Error("The zip reader did not load."));
		}
		return JSZip.loadAsync(inBuffer).then(function(zip) {
			var names = Object.keys(zip.files);
			var mimetype = zip.file("mimetype");
			return (mimetype ? mimetype.async("string") : Promise.resolve("")).then(function(mt) {
				var verdict = FT.resolve(inName, inBytes.subarray(0, 8), {names: names, mimetype: mt});
				self.kind = verdict.kind;
				switch (verdict.kind) {
				case "odt":
					return self.readParts(zip, ["content.xml", "styles.xml", "meta.xml"]).then(function(files) {
						return self.withImages(zip, {kind: "odt", family: "text", model: Odf.text(files)}, files);
					});
				case "odp":
					return self.readParts(zip, ["content.xml", "styles.xml", "meta.xml"]).then(function(files) {
						return self.withImages(zip, {kind: "odp", family: "slides", model: Odf.presentation(files)}, files);
					});
				case "pptx":
					var parts = names.filter(function(n) {
						return /\.(xml|rels)$/i.test(n) && !/^ppt\/(slides\/)?(notesSlides|comments|embeddings)\//.test(n);
					});
					return self.readParts(zip, parts).then(function(files) {
						var result = {kind: "pptx", family: "slides", model: Pptx.parse(files)};
						result.model.meta = Pptx.coreProps(files["docProps/core.xml"]);
						return self.withImages(zip, result, files);
					});
				case "docx":
					var size = self.partSize(zip, "word/document.xml");
					if (size <= DOCX_PREVIEW_LIMIT && !self.libs.noDocxPreview) {
						// Laid out in the page by docx-preview, which reads the bytes itself.
						return self.readParts(zip, ["docProps/core.xml"]).then(function(files) {
							return {kind: "docx", family: "text", model: null, images: {}, meta: Pptx.coreProps(files["docProps/core.xml"])};
						});
					}
					var wordParts = names.filter(function(n) {
						return /^word\/(document|styles|numbering|footnotes|endnotes)\.xml$|^word\/_rels\/document\.xml\.rels$|^docProps\/core\.xml$/.test(n);
					});
					return self.readParts(zip, wordParts).then(function(files) {
						var result = {kind: "docx", family: "text", model: Docx.parse(files), simplified: true};
						result.model.meta = Pptx.coreProps(files["docProps/core.xml"]);
						return self.withImages(zip, result, files);
					});
				case "xlsx":
				case "ods":
					var xmlSize = 0;
					names.forEach(function(n) {
						if (/^xl\/(worksheets\/[^\/]+\.xml|sharedStrings\.xml)$/.test(n) || n === "content.xml") {
							xmlSize += self.partSize(zip, n);
						}
					});
					if (xmlSize > FULL_SHEET_XML_LIMIT || self.libs.forceStream) {
						return verdict.kind === "xlsx" ? self.streamXlsx(zip, inRequest) : self.streamOds(zip, inRequest);
					}
					return self.openWorkbook(verdict.kind, inBytes);
				default:
					throw self.unsupported(verdict.kind, verdict.label);
				}
			});
		});
	};

	// --- Streamed workbooks ------------------------------------------------

	//* Feeds a zip entry to inSink(chunk) as it inflates; resolves at its end (or when inSink returns false).
	DocEngine.prototype.streamPart = function(inZip, inName, inSink) {
		var f = inZip.file(inName);
		if (!f) {
			return Promise.resolve();
		}
		return new Promise(function(resolve, reject) {
			var helper = f.internalStream("string");
			var stopped = false;
			helper.on("data", function(inChunk) {
				if (stopped) {
					return;
				}
				if (inSink(inChunk) === false) {
					stopped = true;
					helper.pause();
					resolve();
				}
			}).on("error", function(e) {
				if (!stopped) {
					reject(e);
				}
			}).on("end", function() {
				if (!stopped) {
					resolve();
				}
			}).resume();
		});
	};

	/**
		A streamed workbook: open() resolves with the sheets as soon as the
		first rows are in (sheets: [{name, rows, cols, text, kinds, ... ,
		streaming: true}]); inRequest.onProgress then gets
		{sheet, rows: [[row, texts, kinds], ...]} batches and finally
		{sheet, done: true, info: {merges, colWidths, rowHeights, truncated}}
		per sheet and {finished: true}.
	*/
	DocEngine.prototype.streamXlsx = function(inZip, inRequest) {
		var self = this;
		var XLSX = null;
		var progress = (inRequest && inRequest.onProgress) || function() {};
		var wbPath = "xl/workbook.xml";
		return this.readParts(inZip, [wbPath, "xl/_rels/workbook.xml.rels", "xl/styles.xml", "docProps/core.xml"]).then(function(files) {
			var X = need("DocXml", "./Xml.js");
			var wb = X.parse(files[wbPath] || "<x/>");
			var rels = {};
			if (files["xl/_rels/workbook.xml.rels"]) {
				X.children(X.parse(files["xl/_rels/workbook.xml.rels"])).forEach(function(r) {
					rels[r.a.Id] = Pptx.resolvePath(wbPath, r.a.Target);
				});
			}
			var sheets = X.children(X.child(wb, "x:sheets"), "x:sheet").map(function(sh, i) {
				return {name: sh.a.name || "Sheet" + (i + 1), path: rels[sh.a["r:id"]], hidden: !!sh.a.state && sh.a.state !== "visible"};
			}).filter(function(sh) {
				return sh.path && inZip.file(sh.path);
			});
			var styles = Stream.readStyles(files["xl/styles.xml"]);
			// SheetJS (its number formatter) is loaded only when a cell style has a format other than
			// General: it is the largest library, and takes seconds to load on a slow device.
			var formatted = styles.xfFmt.some(function(f) {
				return f !== 0 && f !== "General";
			});
			if (formatted) {
				XLSX = self.lib("XLSX");
			}
			var sst = new Stream.SharedStrings();
			return self.streamPart(inZip, "xl/sharedStrings.xml", function(c) {
				sst.feed(c);
			}).then(function() {
				var strings = sst.end();
				var model = {sheets: sheets.map(function(sh, i) {
					return emptySheet(sh.name, i, sh.hidden);
				}), streaming: true};
				var cellsLeft = STREAM_MAX_CELLS;
				var first = true;
				return new Promise(function(resolveFirst, rejectFirst) {
					var resolved = false;
					var giveFirst = function() {
						if (!resolved) {
							resolved = true;
							resolveFirst({kind: "xlsx", family: "sheet", model: model, images: {}, streamed: true,
								meta: Pptx.coreProps(files["docProps/core.xml"])});
						}
					};
					var chain = Promise.resolve();
					sheets.forEach(function(sh, index) {
						chain = chain.then(function() {
							if (cellsLeft <= 0) {
								progress({sheet: index, done: true, info: {merges: [], colWidths: [], rowHeights: {}, truncated: true}});
								return null;
							}
							var batch = [];
							var flush = function() {
								if (batch.length) {
									if (first) {
										// The first rows go with the result itself.
										first = false;
										var s0 = model.sheets[index];
										batch.forEach(function(b) {
											s0.text[b[0]] = b[1];
											s0.kinds[b[0]] = b[2];
											s0.rows = Math.max(s0.rows, b[0] + 1);
											s0.cols = Math.max(s0.cols, b[1].length);
										});
										batch = [];
										giveFirst();
										return;
									}
									progress({sheet: index, rows: batch});
									batch = [];
								}
							};
							var parser = new Stream.XlsxSheetParser({sst: strings, styles: styles, maxCells: cellsLeft,
								format: XLSX && XLSX.SSF ? function(f, v) {
									return XLSX.SSF.format(f, v);
								} : null,
								onRow: function(r, texts, kinds) {
									batch.push([r, texts, kinds]);
									if (batch.length >= (first ? 120 : STREAM_BATCH_ROWS)) {
										flush();
									}
								}});
							return self.streamPart(inZip, sh.path, function(c) {
								parser.feed(c);
								return !parser.done;
							}).then(function() {
								parser.end();
								flush();
								cellsLeft -= parser.cells;
								var info = {merges: parser.merges, colWidths: parser.colWidths, rowHeights: parser.rowHeights,
									truncated: parser.truncated || cellsLeft <= 0 && index < sheets.length - 1};
								if (index === 0 && !resolved) {
									// A first sheet with very few rows: its info goes with the result too.
									applyInfo(model.sheets[0], info);
									giveFirst();
								} else {
									progress({sheet: index, done: true, info: info});
								}
							});
						});
					});
					chain.then(function() {
						giveFirst();
						progress({finished: true});
					}, function(e) {
						if (!resolved) {
							rejectFirst(e);
						} else {
							progress({finished: true, error: String(e && e.message || e)});
						}
					});
				});
			});
		});
	};

	DocEngine.prototype.streamOds = function(inZip, inRequest) {
		var self = this;
		var progress = (inRequest && inRequest.onProgress) || function() {};
		return this.readParts(inZip, ["meta.xml"]).then(function(files) {
			var model = {sheets: [], streaming: true};
			var meta = Odf.meta(files["meta.xml"]);
			return new Promise(function(resolveFirst, rejectFirst) {
				var resolved = false;
				var batch = [];
				var current = -1;
				var giveFirst = function() {
					if (!resolved) {
						resolved = true;
						resolveFirst({kind: "ods", family: "sheet", model: model, images: {}, streamed: true, meta: meta});
					}
				};
				var flush = function() {
					if (!batch.length) {
						return;
					}
					if (!resolved) {
						var s0 = model.sheets[current];
						batch.forEach(function(b) {
							s0.text[b[0]] = b[1];
							s0.kinds[b[0]] = b[2];
							s0.rows = Math.max(s0.rows, b[0] + 1);
							s0.cols = Math.max(s0.cols, b[1].length);
						});
						batch = [];
						giveFirst();
						return;
					}
					progress({sheet: current, rows: batch});
					batch = [];
				};
				var parser = new Stream.OdsParser({maxCells: STREAM_MAX_CELLS,
					onSheet: function(inName) {
						flush();
						current = model.sheets.length;
						var sheet = emptySheet(inName, current, false);
						if (resolved) {
							progress({newSheet: sheet});
						}
						model.sheets.push(sheet);
					},
					onRow: function(r, texts, kinds) {
						// Rows repeated by the format share their arrays: give each its own.
						batch.push([r, texts.slice(0), kinds.slice(0)]);
						if (batch.length >= (resolved ? STREAM_BATCH_ROWS : 120)) {
							flush();
						}
					},
					onSheetEnd: function(inInfo) {
						flush();
						if (!resolved) {
							applyInfo(model.sheets[current], inInfo);
						} else {
							progress({sheet: current, done: true, info: inInfo});
						}
					}});
				self.streamPart(inZip, "content.xml", function(c) {
					parser.feed(c);
					return !parser.done;
				}).then(function() {
					parser.end();
					flush();
					giveFirst();
					progress({finished: true});
				}, function(e) {
					if (!resolved) {
						rejectFirst(e);
					} else {
						progress({finished: true, error: String(e && e.message || e)});
					}
				});
			});
		});
	};

	function emptySheet(inName, inIndex, inHidden) {
		return {name: inName, rows: 0, cols: 0, text: [], kinds: [], formulas: {}, merges: [], colWidths: [], rowHeights: {},
			truncated: false, hidden: inHidden, index: inIndex, streaming: true};
	}

	//* Applies a sheet's final info (merges, widths, heights) to the page-side sheet model.
	function applyInfo(inSheet, inInfo) {
		inSheet.merges = inInfo.merges || [];
		inSheet.rowHeights = inInfo.rowHeights || {};
		inSheet.truncated = !!inInfo.truncated;
		var widths = [];
		for (var c = 0; c < inSheet.cols; c++) {
			widths.push(inInfo.colWidths && inInfo.colWidths[c] !== undefined ? inInfo.colWidths[c] : 80);
		}
		inSheet.colWidths = widths;
		inSheet.streaming = false;
		return inSheet;
	}

	DocEngine.applyInfo = applyInfo;

	//* The uncompressed size of a zip entry (from the zip's directory, without inflating it).
	DocEngine.prototype.partSize = function(inZip, inName) {
		var f = inZip.file(inName);
		if (!f) {
			return 0;
		}
		var d = f._data;
		return d && typeof d.uncompressedSize === "number" ? d.uncompressedSize : 0;
	};

	DocEngine.prototype.readParts = function(inZip, inNames) {
		var files = {};
		return Promise.all(inNames.map(function(n) {
			var f = inZip.file(n);
			if (!f) {
				return null;
			}
			return f.async("string").then(function(s) {
				files[n] = s;
			});
		})).then(function() {
			return files;
		});
	};

	//* Pulls out the pictures the model uses.
	DocEngine.prototype.withImages = function(inZip, inResult, inFiles) {
		var model = inResult.model;
		inResult.meta = model.meta || {};
		inResult.images = {};
		var list = model.images || [];
		return Promise.all(list.map(function(p) {
			var f = inZip.file(p) || inZip.file(decodeURIComponent(p));
			var ext = p.slice(p.lastIndexOf(".") + 1).toLowerCase();
			if (!f || !imageTypes[ext]) {
				inResult.images[p] = {missing: true, ext: ext};
				return null;
			}
			return f.async("uint8array").then(function(bytes) {
				inResult.images[p] = bytes.length > MAX_IMAGE_BYTES ? {missing: true, ext: ext} : {bytes: bytes, mime: imageTypes[ext]};
			});
		})).then(function() {
			return inResult;
		});
	};

	/**
		Writes a copy with inRequest.edits applied: [{sheet (index in the
		workbook), r, c, text}]. Format: "xlsx", "ods" or "csv". Resolves
		with a Uint8Array.
	*/
	DocEngine.prototype.save = function(inRequest) {
		var self = this;
		return new Promise(function(resolve) {
			var edits = inRequest.edits || [];
			var format = inRequest.format;
			if (self.csvRows) {
				for (var i = 0; i < edits.length; i++) {
					var e = edits[i];
					while (self.csvRows.length <= e.r) {
						self.csvRows.push([]);
					}
					var row = self.csvRows[e.r];
					while (row.length <= e.c) {
						row.push("");
					}
					row[e.c] = e.text;
				}
				if (format === "csv") {
					var text = Text.toCsv(self.csvRows, self.csvDelimiter);
					resolve(new TextEncoder().encode(text));
					return;
				}
				var XLSXc = self.lib("XLSX");
				var wbc = XLSXc.utils.book_new();
				XLSXc.utils.book_append_sheet(wbc, XLSXc.utils.aoa_to_sheet(self.csvRows), "Sheet1");
				resolve(new Uint8Array(Sheets.write(XLSXc, wbc, format)));
				return;
			}
			var XLSX = self.lib("XLSX");
			if (!self.workbook || !XLSX) {
				throw new Error("Nothing to save.");
			}
			for (var j = 0; j < edits.length; j++) {
				Sheets.setCell(XLSX, self.workbook, edits[j].sheet, edits[j].r, edits[j].c, edits[j].text);
			}
			resolve(new Uint8Array(Sheets.write(XLSX, self.workbook, format, inRequest.sheet)));
		});
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocEngine;
	} else {
		root.DocEngine = DocEngine;
	}
})(this);
