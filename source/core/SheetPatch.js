/**
	Saves edited cells into a copy of the workbook's own file: only the
	edited cells change in the sheet's XML, so everything Documents does
	not read (cell styles, conditional formatting, charts, pictures,
	comments, defined names) is kept as it was.

		var files = DocSheetPatch.xlsx(parts, writes, {formulaAt: fn});
		var files = DocSheetPatch.ods(parts, writes);
		// parts: {zipPath: xml} of the parts the patch needs (DocSheetPatch.xlsxParts(), "content.xml")
		// writes: [{sheet, r, c, text, value}]; text as typed ("12.5", "=A1*2", "TRUE", "" to clear),
		//         value the formula's result when known ({v, t} as from DocFormulas)
		//         [{sheet, r, c, valueOnly: true, value}]: a formula cell's result only (one that uses an edited cell)
		// files: {zipPath: new xml, or null to remove the part}

	formulaAt(sheet, r, c) gives a cell's formula in A1 syntax without the
	"=" (SheetJS has it for every cell): a cell that shares a formula with
	an edited one gets its own copy of it, since the edit removes the one
	they shared.
*/
(function(root) {
	var X = root.DocXml || (typeof require === "function" ? require("./Xml.js") : null);
	var Sheets = root.DocSheets || (typeof require === "function" ? require("./Sheets.js") : null);

	// --- Shared helpers --------------------------------------------------

	//* The writes, the last one per cell, grouped by sheet: {sheet: {r: {c: write}}}.
	function bySheet(inWrites) {
		var out = {};
		for (var i = 0; i < inWrites.length; i++) {
			var w = inWrites[i];
			var s = out[w.sheet] = out[w.sheet] || {};
			var row = s[w.r] = s[w.r] || {};
			row[w.c] = w;
		}
		return out;
	}

	function sortedKeys(inObj) {
		return Object.keys(inObj).map(Number).sort(function(a, b) {
			return a - b;
		});
	}

	//* The prefix (with ":", or "") the document gave the namespace whose canonical prefix is inCanon.
	function prefixOf(inScope, inCanon, inFallback) {
		// The scope's prototype chain holds the outer declarations: for..in walks them too.
		for (var k in inScope) {
			if (inScope[k] === inCanon) {
				return k ? k + ":" : "";
			}
		}
		return inFallback;
	}

	//* The raw prefix of an element's own name ("x:row" -> "x:", "row" -> "").
	function ownPrefix(inEl) {
		var i = inEl.raw.indexOf(":");
		return i < 0 ? "" : inEl.raw.slice(0, i + 1);
	}

	//* inTag (a start tag's text) with attribute inName set to inValue (null removes it).
	function setAttr(inTag, inName, inValue) {
		var esc = inName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
		var re = new RegExp("\\s" + esc + "\\s*=\\s*(\"[^\"]*\"|'[^']*')");
		var tag = inTag.replace(re, "");
		if (inValue === null || inValue === undefined) {
			return tag;
		}
		var close = /\/>$/.test(tag) ? 2 : 1;
		return tag.slice(0, tag.length - close) + " " + inName + "=\"" + X.escape(String(inValue)) + "\"" + tag.slice(tag.length - close);
	}

	function startTag(inXml, inEl) {
		return inXml.slice(inEl.p, inEl.ci);
	}

	//* A start tag as an empty element's tag, and an empty one as a start tag.
	function asEmpty(inTag) {
		return /\/>$/.test(inTag) ? inTag : inTag.replace(/\s*>$/, "/>");
	}

	function asOpen(inTag) {
		return inTag.replace(/\s*\/>$/, ">");
	}

	//* The cached result to write for {v, t}, or null when there is none to write.
	function cachedOf(inValue) {
		if (!inValue || inValue.name || inValue.t === "n" && !(typeof inValue.v === "number" && isFinite(inValue.v))) {
			// Unknown, or a function the calculator does not know: the reading program works it out.
			return null;
		}
		return inValue;
	}

	function numberText(inValue) {
		var s = String(inValue);
		return /^-?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s) ? s : "0";
	}

	// --- Office Open XML -------------------------------------------------

	//* The parts needed besides the sheets: the workbook, its relationships and the content types.
	function xlsxParts() {
		return ["xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"];
	}

	function resolve(inBase, inTarget) {
		if (inTarget.charAt(0) === "/") {
			return inTarget.slice(1);
		}
		var parts = inBase.split("/");
		parts.pop();
		inTarget.split("/").forEach(function(p) {
			if (p === "..") {
				parts.pop();
			} else if (p !== ".") {
				parts.push(p);
			}
		});
		return parts.join("/");
	}

	//* The zip paths of the workbook's sheets, in the workbook's order.
	function xlsxSheetPaths(inParts) {
		var wb = X.parse(inParts["xl/workbook.xml"] || "<x/>");
		var rels = {};
		if (inParts["xl/_rels/workbook.xml.rels"]) {
			X.children(X.parse(inParts["xl/_rels/workbook.xml.rels"])).forEach(function(r) {
				rels[r.a.Id] = resolve("xl/workbook.xml", r.a.Target || "");
			});
		}
		return X.children(X.child(wb, "x:sheets"), "x:sheet").map(function(sh) {
			return rels[sh.a["r:id"]] || null;
		});
	}

	//* The <c> element for one write; inOld the cell it replaces (its style is kept).
	function xlsxCell(inPrefix, inRef, inStyle, inWrite) {
		var p = Sheets.parseInput(inWrite.text);
		var attrs = " r=\"" + inRef + "\"" + (inStyle ? " s=\"" + X.escape(inStyle) + "\"" : "");
		var c = inPrefix + "c", v = inPrefix + "v";
		if (p.f) {
			var cv = xlsxCached(inPrefix, inWrite.value);
			return "<" + c + attrs + (cv.t ? " t=\"" + cv.t + "\"" : "") + "><" + inPrefix + "f>" + X.escape(p.f) + "</" + inPrefix + "f>" +
				cv.xml + "</" + c + ">";
		}
		if (p.t === "z") {
			// Cleared: the cell keeps its style, if it had one.
			return inStyle ? "<" + c + attrs + "/>" : "";
		}
		if (p.t === "n") {
			return "<" + c + attrs + "><" + v + ">" + numberText(p.v) + "</" + v + "></" + c + ">";
		}
		if (p.t === "b") {
			return "<" + c + attrs + " t=\"b\"><" + v + ">" + (p.v ? 1 : 0) + "</" + v + "></" + c + ">";
		}
		var t2 = inPrefix + "t";
		return "<" + c + attrs + " t=\"inlineStr\"><" + inPrefix + "is><" + t2 + " xml:space=\"preserve\">" + X.escape(p.v) +
			"</" + t2 + "></" + inPrefix + "is></" + c + ">";
	}

	//* A formula's cached result: its type attribute ("" for a number) and its <v>.
	function xlsxCached(inPrefix, inValue) {
		var c = cachedOf(inValue);
		var v = inPrefix + "v";
		if (!c) {
			return {t: "", xml: ""};
		}
		var text = c.t === "n" ? numberText(c.v) : c.t === "b" ? (c.v ? "1" : "0") : X.escape(String(c.v));
		return {t: c.t === "n" ? "" : c.t === "s" ? "str" : c.t, xml: "<" + v + ">" + text + "</" + v + ">"};
	}

	//* A formula cell with a new result; its formula (inFormula, the <f> to write) and style are kept.
	function xlsxResult(inXml, inPrefix, inOld, inValue, inFormula) {
		var cv = xlsxCached(inPrefix, inValue);
		var tag = asOpen(setAttr(startTag(inXml, inOld), "t", cv.t || null));
		return tag + inFormula + cv.xml + "</" + inOld.raw + ">";
	}

	function xlsxSheet(inXml, inSheet, inRows, inFormulaAt) {
		var rootEl = X.parse(inXml, {offsets: true});
		var sd = X.child(rootEl, "x:sheetData");
		if (!sd) {
			throw new Error("The sheet has no cell data.");
		}
		var px = ownPrefix(sd);
		var splices = [];
		var rowsByIndex = {}, rowList = [];
		var shared = {};       // si -> master cell, of the shared formulas
		var sharers = [];      // [{el, r, c, si}] cells using a shared formula
		var prevRow = -1;
		X.children(sd, "x:row").forEach(function(row) {
			var ri = row.a.r ? +row.a.r - 1 : prevRow + 1;
			prevRow = ri;
			var cells = [];
			var prevCol = -1;
			X.children(row, "x:c").forEach(function(cell) {
				var ref = cell.a.r ? Sheets.parseRef(cell.a.r) : null;
				var ci = ref ? ref.c : prevCol + 1;
				prevCol = ci;
				cells.push({el: cell, c: ci});
				var f = X.child(cell, "x:f");
				if (f && f.a.t === "shared" && f.a.si !== undefined) {
					if (f.a.ref) {
						shared[f.a.si] = {r: ri, c: ci};
					}
					sharers.push({f: f, r: ri, c: ci, si: f.a.si, master: !!f.a.ref});
				}
			});
			var entry = {el: row, r: ri, cells: cells};
			rowsByIndex[ri] = entry;
			rowList.push(entry);
		});
		var minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;
		var rowIdx = sortedKeys(inRows);
		// Shared formulas whose first cell is overwritten: the cells sharing it need it written out.
		var brokenShares = {};
		rowIdx.forEach(function(r) {
			var entry = rowsByIndex[r];
			(entry ? entry.cells : []).forEach(function(cell) {
				var w = inRows[r][cell.c];
				var f = w && !w.valueOnly && X.child(cell.el, "x:f");
				if (f && f.a.t === "shared" && f.a.ref) {
					brokenShares[f.a.si] = true;
				}
			});
		});
		//* The <f> of an existing formula cell as it is written: a broken share written out.
		var formulaXml = function(inF, inRow, inCol) {
			if (inF.a.t === "shared" && !inF.a.ref && brokenShares[inF.a.si]) {
				var formula = inFormulaAt ? inFormulaAt(inSheet, inRow, inCol) : null;
				if (!formula) {
					throw new Error("A formula shared with an edited cell could not be kept.");
				}
				return "<" + inF.raw + ">" + X.escape(formula) + "</" + inF.raw + ">";
			}
			return inXml.slice(inF.p, inF.e);
		};
		var newRows = [];      // rows to create: {r, xml}
		rowIdx.forEach(function(r) {
			var entry = rowsByIndex[r];
			var cols = sortedKeys(inRows[r]).filter(function(c) {
				return !inRows[r][c].valueOnly || entry;
			});
			if (!cols.length) {
				return;
			}
			var cellXml = function(inC, inOld) {
				var w = inRows[r][inC];
				minR = Math.min(minR, r);
				maxR = Math.max(maxR, r);
				minC = Math.min(minC, inC);
				maxC = Math.max(maxC, inC);
				if (w.valueOnly) {
					// A result only: there is nothing to change in a cell that is gone or holds no formula.
					var of = inOld && X.child(inOld, "x:f");
					return of ? xlsxResult(inXml, px, inOld, w.value, formulaXml(of, r, inC)) : null;
				}
				var style = inOld ? inOld.a.s : entry && entry.el.a.customFormat === "1" ? entry.el.a.s : null;
				return xlsxCell(px, Sheets.cellRef(r, inC), style, w);
			};
			if (!entry) {
				newRows.push({r: r, xml: "<" + px + "row r=\"" + (r + 1) + "\">" + cols.map(function(c) {
					return cellXml(c, null);
				}).join("") + "</" + px + "row>"});
				return;
			}
			var row = entry.el;
			var inserts = [];   // cells after the last one with a lower column
			cols.forEach(function(c) {
				var at = null, old = null;
				for (var i = 0; i < entry.cells.length; i++) {
					if (entry.cells[i].c === c) {
						old = entry.cells[i].el;
						break;
					}
					if (entry.cells[i].c > c) {
						at = entry.cells[i].el.p;
						break;
					}
				}
				if (old) {
					var nx = cellXml(c, old);
					if (nx !== null) {
						splices.push({s: old.p, e: old.e, v: nx});
					}
				} else if (!inRows[r][c].valueOnly) {
					inserts.push({at: at === null ? row.ce : at, xml: cellXml(c, null)});
				}
			});
			if (inserts.length) {
				if (row.ci === row.e) {
					// <row .../> becomes <row ...>cells</row>.
					splices.push({s: row.p, e: row.e, v: asOpen(startTag(inXml, row)) + inserts.map(function(x) {
						return x.xml;
					}).join("") + "</" + row.raw + ">"});
				} else {
					inserts.forEach(function(x) {
						splices.push({s: x.at, e: x.at, v: x.xml});
					});
				}
			}
		});
		if (newRows.length) {
			var groups = {};
			newRows.forEach(function(nr) {
				var at = sd.ce;
				for (var i = 0; i < rowList.length; i++) {
					if (rowList[i].r > nr.r) {
						at = rowList[i].el.p;
						break;
					}
				}
				(groups[at] = groups[at] || []).push(nr.xml);
			});
			if (sd.ci === sd.e) {
				splices.push({s: sd.p, e: sd.e, v: asOpen(startTag(inXml, sd)) + newRows.map(function(nr) {
					return nr.xml;
				}).join("") + "</" + sd.raw + ">"});
			} else {
				Object.keys(groups).forEach(function(at) {
					splices.push({s: +at, e: +at, v: groups[at].join("")});
				});
			}
		}
		// A shared formula whose first cell was edited: the other cells (not written above) get it written out.
		sharers.forEach(function(sh) {
			if (!brokenShares[sh.si] || sh.master || inRows[sh.r] && inRows[sh.r][sh.c]) {
				return;
			}
			splices.push({s: sh.f.p, e: sh.f.e, v: formulaXml(sh.f, sh.r, sh.c)});
		});
		// The used range grows to take new cells.
		var dim = X.child(rootEl, "x:dimension");
		if (dim && maxR >= 0) {
			var range = String(dim.a.ref || "A1").split(":");
			var a = Sheets.parseRef(range[0]) || {r: 0, c: 0};
			var b = Sheets.parseRef(range[1] || range[0]) || a;
			var ref = Sheets.cellRef(Math.min(a.r, minR), Math.min(a.c, minC)) + ":" +
				Sheets.cellRef(Math.max(b.r, maxR), Math.max(b.c, maxC));
			splices.push({s: dim.p, e: dim.ci, v: setAttr(startTag(inXml, dim), "ref", ref)});
		}
		return X.splice(inXml, splices);
	}

	/**
		The workbook's parts after the writes. Formulas are worked out again
		when the copy is opened (fullCalcOnLoad), and the calculation chain,
		which lists the formula cells, is dropped as it may no longer match.
	*/
	function xlsx(inParts, inWrites, inOptions) {
		var opts = inOptions || {};
		var out = {};
		var paths = xlsxSheetPaths(inParts);
		var groups = bySheet(inWrites);
		Object.keys(groups).forEach(function(s) {
			var p = paths[s];
			if (!p || inParts[p] === undefined) {
				throw new Error("The edited sheet was not found in the workbook.");
			}
			out[p] = xlsxSheet(inParts[p], +s, groups[s], opts.formulaAt);
		});
		var wbXml = inParts["xl/workbook.xml"];
		var wb = X.parse(wbXml, {offsets: true});
		var calc = X.child(wb, "x:calcPr");
		if (calc) {
			out["xl/workbook.xml"] = X.splice(wbXml, [{s: calc.p, e: calc.ci, v: setAttr(startTag(wbXml, calc), "fullCalcOnLoad", "1")}]);
		} else {
			// After the elements calcPr follows in the schema.
			var after = null;
			["x:sheets", "x:functionGroups", "x:externalReferences", "x:definedNames"].forEach(function(n) {
				var el = X.child(wb, n);
				if (el) {
					after = el.e;
				}
			});
			if (after !== null) {
				var sheetsEl = X.child(wb, "x:sheets");
				out["xl/workbook.xml"] = X.splice(wbXml, [{s: after, e: after, v: "<" + ownPrefix(sheetsEl) + "calcPr fullCalcOnLoad=\"1\"/>"}]);
			}
		}
		if (inParts["xl/_rels/workbook.xml.rels"] && /calcChain/.test(inParts["xl/_rels/workbook.xml.rels"])) {
			out["xl/_rels/workbook.xml.rels"] = inParts["xl/_rels/workbook.xml.rels"].replace(/<([\w]+:)?Relationship\b[^>]*\/calcChain"[^>]*\/>/g, "");
			out["xl/calcChain.xml"] = null;
			if (inParts["[Content_Types].xml"]) {
				out["[Content_Types].xml"] = inParts["[Content_Types].xml"].replace(/<([\w]+:)?Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, "");
			}
		}
		return out;
	}

	// --- OpenDocument ----------------------------------------------------

	/**
		A formula in A1 syntax ("SUM(A1:B2)*Sheet2!C3") in OpenFormula,
		as ODF stores it ("of:=SUM([.A1:.B2])*[$Sheet2.C3]").
	*/
	function toOpenFormula(inFormula) {
		var f = String(inFormula);
		var out = "";
		var re = /"(?:[^"]|"")*"|((?:'(?:[^']|'')+'|[A-Za-z_\u00c0-\uffff][\w.\u00c0-\uffff]*)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![\w(!])|,/g;
		var last = 0, m;
		while ((m = re.exec(f))) {
			var before = f.charAt(m.index - 1);
			if (m[2] && !m[1] && /[\w.$]/.test(before)) {
				// Part of a name ("LOG10(" is caught by the look-ahead, "A.B1" here).
				continue;
			}
			out += f.slice(last, m.index);
			last = m.index + m[0].length;
			if (m[0] === ",") {
				out += ";";
			} else if (m[0].charAt(0) === "\"") {
				out += m[0];
			} else {
				var sheet = m[1] ? m[1].slice(0, -1) : null;
				var ends = m[2].split(":");
				var first = (sheet ? "$" + sheet : "") + "." + ends[0];
				out += "[" + first + (ends[1] ? ":." + ends[1] : "") + "]";
			}
		}
		return "of:=" + out + f.slice(last);
	}

	//* The rows of a table, through the groups that hold them: [{el, start, count}].
	function odsRows(inTable) {
		var out = [];
		var at = 0;
		var walk = function(inEl) {
			for (var i = 0; i < inEl.c.length; i++) {
				var c = inEl.c[i];
				if (typeof c !== "object") {
					continue;
				}
				if (c.n === "table:table-row") {
					var n = Math.max(1, +(c.a["table:number-rows-repeated"] || 1) || 1);
					out.push({el: c, start: at, count: n});
					at += n;
				} else if (c.n === "table:table-header-rows" || c.n === "table:table-rows" || c.n === "table:table-row-group") {
					walk(c);
				}
			}
		};
		walk(inTable);
		return out;
	}

	function odsCells(inRow) {
		var out = [];
		var at = 0;
		for (var i = 0; i < inRow.c.length; i++) {
			var c = inRow.c[i];
			if (typeof c === "object" && (c.n === "table:table-cell" || c.n === "table:covered-table-cell")) {
				var n = Math.max(1, +(c.a["table:number-columns-repeated"] || 1) || 1);
				out.push({el: c, start: at, count: n});
				at += n;
			}
		}
		return out;
	}

	function Ods(inXml, inScope) {
		this.xml = inXml;
		this.t = prefixOf(inScope, "table", "table:");
		this.o = prefixOf(inScope, "office", "office:");
		this.x = prefixOf(inScope, "text", "text:");
	}

	//* The value attributes of a result ({v, t}) and the text shown for it.
	Ods.prototype.valueAttrs = function(inValue) {
		var o = this.o;
		var c = cachedOf(inValue);
		if (!c) {
			return {attrs: "", shown: ""};
		}
		if (c.t === "n") {
			return {attrs: " " + o + "value-type=\"float\" " + o + "value=\"" + numberText(c.v) + "\"", shown: String(c.v)};
		}
		if (c.t === "b") {
			return {attrs: " " + o + "value-type=\"boolean\" " + o + "boolean-value=\"" + !!c.v + "\"", shown: c.v ? "TRUE" : "FALSE"};
		}
		return {attrs: " " + o + "value-type=\"string\" " + o + "string-value=\"" + X.escape(String(c.v)) + "\"", shown: String(c.v)};
	};

	//* A formula cell with a new result: its formula and style are kept, its value and text replaced.
	Ods.prototype.result = function(inOld, inValue) {
		var self = this, o = this.o;
		var tag = setAttr(startTag(this.xml, inOld), this.t + "number-columns-repeated", null);
		["value-type", "value", "boolean-value", "string-value", "date-value", "time-value", "currency"].forEach(function(a) {
			tag = setAttr(tag, o + a, null);
		});
		tag = setAttr(tag, "calcext:value-type", null);
		var va = this.valueAttrs(inValue);
		tag = tag.replace(/\s*(\/?)>$/, va.attrs + "$1>");
		var keep = "";
		for (var i = 0; i < inOld.c.length; i++) {
			var c = inOld.c[i];
			if (typeof c === "object" && c.n !== "text:p") {
				keep += self.xml.slice(c.p, c.e);
			}
		}
		var content = keep + (va.shown === "" ? "" : "<" + this.x + "p>" + X.escape(va.shown) + "</" + this.x + "p>");
		return content ? asOpen(tag) + content + "</" + inOld.raw + ">" : asEmpty(tag);
	};

	//* The text of one cell written by inWrite; inOld the cell it replaces (its style and spans are kept).
	Ods.prototype.cell = function(inOld, inWrite) {
		var t = this.t, o = this.o;
		if (inWrite.valueOnly) {
			return inOld.a["table:formula"] ? this.result(inOld, inWrite.value) : this.repeated(inOld, "number-columns-repeated", 1);
		}
		var name = inOld ? inOld.raw : t + "table-cell";
		var tag = "<" + name;
		if (inOld) {
			[t + "style-name", t + "number-columns-spanned", t + "number-rows-spanned", t + "content-validation-name"].forEach(function(a) {
				var raw = inOld.a[a.replace(t, "table:")];
				if (raw !== undefined) {
					tag += " " + a + "=\"" + X.escape(raw) + "\"";
				}
			});
		}
		var p = Sheets.parseInput(inWrite.text);
		var shown = inWrite.text;
		if (p.f) {
			var va = this.valueAttrs(inWrite.value);
			tag += " " + t + "formula=\"" + X.escape(toOpenFormula(p.f)) + "\"" + va.attrs;
			shown = va.shown;
		} else if (p.t === "n") {
			tag += " " + o + "value-type=\"float\" " + o + "value=\"" + numberText(p.v) + "\"";
		} else if (p.t === "b") {
			tag += " " + o + "value-type=\"boolean\" " + o + "boolean-value=\"" + p.v + "\"";
			shown = p.v ? "TRUE" : "FALSE";
		} else if (p.t === "s") {
			tag += " " + o + "value-type=\"string\"";
		}
		// Comments and pictures in the cell stay with it; its paragraphs are the new value.
		var keep = "";
		if (inOld) {
			for (var i = 0; i < inOld.c.length; i++) {
				var c = inOld.c[i];
				if (typeof c === "object" && (c.n === "office:annotation" || c.n.indexOf("draw:") === 0)) {
					keep += this.xml.slice(c.p, c.e);
				}
			}
		}
		var content = keep + (shown === "" ? "" : "<" + this.x + "p>" + X.escape(shown) + "</" + this.x + "p>");
		return content ? tag + ">" + content + "</" + name + ">" : tag + "/>";
	};

	//* An element's text with its repeat count set (1 removes the attribute).
	Ods.prototype.repeated = function(inEl, inAttr, inCount) {
		var tag = setAttr(startTag(this.xml, inEl), this.t + inAttr, inCount > 1 ? inCount : null);
		return tag + this.xml.slice(inEl.ci, inEl.e);
	};

	//* A row's text with the writes in inCols ({c: write}) applied; inCount repeats of it. No inRow: a new row.
	Ods.prototype.row = function(inRow, inCols, inCount) {
		var self = this, t = this.t;
		var cells = inRow ? odsCells(inRow) : [];
		var cols = sortedKeys(inCols);
		var pieces = [];
		var at = inRow ? inRow.ci : 0;       // copied up to here
		var ci = 0;
		var end = 0;
		cells.forEach(function(cell) {
			var mine = [];
			while (ci < cols.length && cols[ci] < cell.start + cell.count) {
				mine.push(cols[ci++]);
			}
			end = cell.start + cell.count;
			if (!mine.length) {
				return;
			}
			pieces.push(self.xml.slice(at, cell.el.p));
			var from = cell.start;
			mine.forEach(function(c) {
				if (c > from) {
					pieces.push(self.repeated(cell.el, "number-columns-repeated", c - from));
				}
				pieces.push(self.cell(cell.el, inCols[c]));
				from = c + 1;
			});
			if (from < cell.start + cell.count) {
				pieces.push(self.repeated(cell.el, "number-columns-repeated", cell.start + cell.count - from));
			}
			at = cell.el.e;
		});
		var tail = "";
		for (; ci < cols.length; ci++) {
			var c = cols[ci];
			if (inCols[c].valueOnly) {
				continue;
			}
			if (c > end) {
				tail += "<" + t + "table-cell" + (c - end > 1 ? " " + t + "number-columns-repeated=\"" + (c - end) + "\"" : "") + "/>";
			}
			tail += self.cell(null, inCols[c]);
			end = c + 1;
		}
		if (!inRow) {
			return "<" + t + "table-row>" + tail + "</" + t + "table-row>";
		}
		var tag = setAttr(startTag(this.xml, inRow), t + "number-rows-repeated", inCount > 1 ? inCount : null);
		if (inRow.ci === inRow.e) {
			return asOpen(tag) + tail + "</" + inRow.raw + ">";
		}
		pieces.push(this.xml.slice(at, inRow.ce), tail);
		return tag + pieces.join("") + this.xml.slice(inRow.ce, inRow.e);
	};

	function ods(inParts, inWrites) {
		var xml = inParts["content.xml"];
		var doc = X.parse(xml, {offsets: true});
		var body = X.path(doc, "office:body", "office:spreadsheet");
		if (!body) {
			throw new Error("The workbook has no tables.");
		}
		var tables = X.children(body, "table:table");
		var o = new Ods(xml, body.ns);
		var t = o.t;
		var splices = [];
		var groups = bySheet(inWrites);
		Object.keys(groups).forEach(function(s) {
			var table = tables[s];
			if (!table) {
				throw new Error("The edited sheet was not found in the workbook.");
			}
			var rows = odsRows(table);
			var wanted = sortedKeys(groups[s]);
			var wi = 0;
			var end = 0;
			rows.forEach(function(row) {
				var mine = [];
				while (wi < wanted.length && wanted[wi] < row.start + row.count) {
					mine.push(wanted[wi++]);
				}
				end = row.start + row.count;
				if (!mine.length) {
					return;
				}
				var out = "";
				var from = row.start;
				mine.forEach(function(r) {
					if (r > from) {
						out += o.repeated(row.el, "number-rows-repeated", r - from);
					}
					out += o.row(row.el, groups[s][r], 1);
					from = r + 1;
				});
				if (from < row.start + row.count) {
					out += o.repeated(row.el, "number-rows-repeated", row.start + row.count - from);
				}
				splices.push({s: row.el.p, e: row.el.e, v: out});
			});
			if (wi < wanted.length) {
				// Past the last row: new rows after it.
				var add = "";
				for (; wi < wanted.length; wi++) {
					var r = wanted[wi];
					if (!sortedKeys(groups[s][r]).some(function(c) {
						return !groups[s][r][c].valueOnly;
					})) {
						continue;
					}
					if (r > end) {
						add += "<" + t + "table-row" + (r - end > 1 ? " " + t + "number-rows-repeated=\"" + (r - end) + "\"" : "") +
							"><" + t + "table-cell/></" + t + "table-row>";
					}
					add += o.row(null, groups[s][r], 1);
					end = r + 1;
				}
				var lastRow = rows.length ? rows[rows.length - 1].el : null;
				var at = lastRow ? lastRow.e : table.ce;
				splices.push({s: at, e: at, v: add});
			}
		});
		return {"content.xml": X.splice(xml, splices)};
	}

	var DocSheetPatch = {
		xlsx: xlsx,
		xlsxParts: xlsxParts,
		xlsxSheetPaths: xlsxSheetPaths,
		ods: ods,
		toOpenFormula: toOpenFormula,
		setAttr: setAttr
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocSheetPatch;
	} else {
		root.DocSheetPatch = DocSheetPatch;
	}
})(this);
