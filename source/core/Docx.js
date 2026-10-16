/**
	Word (.docx) to the document model of Odf.js (blocks of paragraphs,
	tables and pictures), for documents too large for docx-preview to lay
	out in reasonable time on a slow device. It runs in the worker like the
	ODF reader; the page then renders the blocks a slice at a time.

		var doc = DocDocx.parse(files);   // files: "word/document.xml", "word/styles.xml",
		                                  // "word/numbering.xml", "word/footnotes.xml", "word/_rels/document.xml.rels", ...
		doc.blocks, doc.page, doc.images, doc.base

	Kept: paragraph and character styles (bold, italic, underline, strike,
	colour, highlight, size, font, alignment, indents, spacing), headings,
	numbered and bulleted lists with their real labels, tables with spans
	and shading, pictures, hyperlinks, footnotes and endnotes, page and
	section breaks, the result text of fields. Not kept: headers and
	footers, text boxes' positions, columns, tracked deletions (dropped).
*/
(function(root) {
	var X = root.DocXml || (typeof require === "function" ? require("./Xml.js") : null);
	var Pptx = root.DocPptx || (typeof require === "function" ? require("./Pptx.js") : null);

	function twipsToPt(inTwips) {
		return Math.round(+inTwips / 20 * 10) / 10;
	}

	function val(inEl) {
		return inEl ? inEl.a["w:val"] : undefined;
	}

	//* An on/off property (<w:b/>, <w:b w:val="0"/>).
	function flag(inEl) {
		if (!inEl) {
			return undefined;
		}
		var v = inEl.a["w:val"];
		return !(v === "0" || v === "false" || v === "off" || v === "none");
	}

	var highlights = {yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff", red: "#ff0000",
		darkBlue: "#000080", darkCyan: "#008080", darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000",
		darkYellow: "#808000", darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000", white: "#ffffff"};

	function hex(inValue) {
		return /^[0-9a-fA-F]{6}$/.test(inValue || "") ? "#" + inValue.toLowerCase() : null;
	}

	//* Run properties (w:rPr) into CSS over inCss.
	function runCss(inRPr, inCss) {
		var css = inCss || {};
		if (!inRPr) {
			return css;
		}
		var kids = X.children(inRPr);
		for (var i = 0; i < kids.length; i++) {
			var k = kids[i];
			switch (k.n) {
			case "w:b":
				css.fontWeight = flag(k) ? "bold" : "normal";
				break;
			case "w:i":
				css.fontStyle = flag(k) ? "italic" : "normal";
				break;
			case "w:u":
				css.underline = val(k) !== "none";
				break;
			case "w:strike":
			case "w:dstrike":
				css.strike = flag(k);
				break;
			case "w:color":
				var c = hex(val(k));
				if (c) {
					css.color = c;
				}
				break;
			case "w:sz":
				if (+val(k)) {
					css.fontSize = (+val(k) / 2) + "pt";
				}
				break;
			case "w:highlight":
				if (highlights[val(k)]) {
					css.backgroundColor = highlights[val(k)];
				}
				break;
			case "w:shd":
				var f = hex(k.a["w:fill"]);
				if (f && f !== "#ffffff") {
					css.backgroundColor = f;
				}
				break;
			case "w:vertAlign":
				if (val(k) === "superscript" || val(k) === "subscript") {
					css.verticalAlign = val(k) === "superscript" ? "super" : "sub";
					css.fontSize = "0.7em";
				}
				break;
			case "w:caps":
				if (flag(k)) {
					css.textTransform = "uppercase";
				}
				break;
			case "w:smallCaps":
				if (flag(k)) {
					css.fontVariant = "small-caps";
				}
				break;
			case "w:rFonts":
				var font = k.a["w:ascii"] || k.a["w:hAnsi"];
				if (font && /^[-a-zA-Z0-9 ]+$/.test(font)) {
					css.fontFamily = "'" + font + "', sans-serif";
				}
				break;
			case "w:vanish":
				if (flag(k)) {
					css.hidden = true;
				}
				break;
			}
		}
		return css;
	}

	var aligns = {left: "left", start: "left", center: "center", right: "right", end: "right", both: "justify", distribute: "justify"};

	//* Paragraph properties (w:pPr) into CSS over inCss; also numbering and breaks in inExtra.
	function paraCss(inPPr, inCss, inExtra) {
		var css = inCss || {};
		if (!inPPr) {
			return css;
		}
		var kids = X.children(inPPr);
		for (var i = 0; i < kids.length; i++) {
			var k = kids[i];
			switch (k.n) {
			case "w:jc":
				if (aligns[val(k)]) {
					css.textAlign = aligns[val(k)];
				}
				break;
			case "w:ind":
				var left = k.a["w:left"] || k.a["w:start"];
				if (left !== undefined) {
					css.marginLeft = twipsToPt(left) + "pt";
				}
				var right = k.a["w:right"] || k.a["w:end"];
				if (right !== undefined) {
					css.marginRight = twipsToPt(right) + "pt";
				}
				if (k.a["w:hanging"] !== undefined) {
					css.textIndent = "-" + twipsToPt(k.a["w:hanging"]) + "pt";
				} else if (k.a["w:firstLine"] !== undefined) {
					css.textIndent = twipsToPt(k.a["w:firstLine"]) + "pt";
				}
				break;
			case "w:spacing":
				if (k.a["w:before"] !== undefined) {
					css.marginTop = twipsToPt(k.a["w:before"]) + "pt";
				}
				if (k.a["w:after"] !== undefined) {
					css.marginBottom = twipsToPt(k.a["w:after"]) + "pt";
				}
				if (k.a["w:line"] !== undefined && (!k.a["w:lineRule"] || k.a["w:lineRule"] === "auto")) {
					css.lineHeight = String(Math.round(+k.a["w:line"] / 240 * 1.15 * 100) / 100);
				}
				break;
			case "w:shd":
				var f = hex(k.a["w:fill"]);
				if (f && f !== "#ffffff") {
					css.backgroundColor = f;
				}
				break;
			case "w:pageBreakBefore":
				if (inExtra && flag(k)) {
					inExtra.pageBreak = true;
				}
				break;
			case "w:numPr":
				if (inExtra) {
					var numId = val(X.child(k, "w:numId")), ilvl = val(X.child(k, "w:ilvl"));
					if (numId !== undefined) {
						inExtra.numId = numId;
					}
					if (ilvl !== undefined) {
						inExtra.ilvl = +ilvl;
					}
				}
				break;
			case "w:outlineLvl":
				if (inExtra) {
					inExtra.outline = +val(k) + 1;
				}
				break;
			}
		}
		return css;
	}

	//* CSS a renderer can apply: the internal underline/strike flags become text-decoration.
	function finishCss(inCss) {
		var out = {};
		for (var k in inCss) {
			if (inCss.hasOwnProperty(k) && k !== "underline" && k !== "strike" && k !== "hidden") {
				out[k] = inCss[k];
			}
		}
		var deco = [];
		if (inCss.underline) {
			deco.push("underline");
		}
		if (inCss.strike) {
			deco.push("line-through");
		}
		if (deco.length) {
			out.textDecoration = deco.join(" ");
		}
		return out;
	}

	function merge(inA, inB) {
		var out = {}, k;
		for (k in inA) {
			if (inA.hasOwnProperty(k)) {
				out[k] = inA[k];
			}
		}
		for (k in inB) {
			if (inB.hasOwnProperty(k)) {
				out[k] = inB[k];
			}
		}
		return out;
	}

	function isEmpty(inObj) {
		for (var k in inObj) {
			if (inObj.hasOwnProperty(k)) {
				return false;
			}
		}
		return true;
	}

	// --- Styles and numbering ---------------------------------------------

	function Styles(inXml) {
		this.byId = {};
		this.defaultPara = null;
		this.defaultRun = {};
		this.defaultPPr = {};
		if (!inXml) {
			return;
		}
		var doc = X.parse(inXml);
		var defaults = X.child(doc, "w:docDefaults");
		if (defaults) {
			this.defaultRun = runCss(X.path(defaults, "w:rPrDefault", "w:rPr"), {});
			this.defaultPPr = paraCss(X.path(defaults, "w:pPrDefault", "w:pPr"), {});
		}
		var list = X.children(doc, "w:style");
		for (var i = 0; i < list.length; i++) {
			var s = list[i];
			var id = s.a["w:styleId"];
			var name = val(X.child(s, "w:name")) || "";
			var extra = {};
			this.byId[id] = {
				type: s.a["w:type"], basedOn: val(X.child(s, "w:basedOn")), name: name,
				run: runCss(X.child(s, "w:rPr"), {}),
				para: paraCss(X.child(s, "w:pPr"), {}, extra),
				extra: extra
			};
			if (s.a["w:type"] === "paragraph" && (s.a["w:default"] === "1" || s.a["w:default"] === "true")) {
				this.defaultPara = id;
			}
		}
	}

	//* {run, para, extra, heading} of a style with its bases.
	Styles.prototype.resolve = function(inId) {
		var chain = [];
		var id = inId, guard = 0;
		while (id && this.byId[id] && guard++ < 32) {
			chain.unshift(this.byId[id]);
			id = this.byId[id].basedOn;
		}
		var out = {run: {}, para: {}, extra: {}, heading: 0};
		for (var i = 0; i < chain.length; i++) {
			out.run = merge(out.run, chain[i].run);
			out.para = merge(out.para, chain[i].para);
			out.extra = merge(out.extra, chain[i].extra);
			var m = /^heading (\d)$/i.exec(chain[i].name);
			if (m) {
				out.heading = +m[1];
			} else if (/^title$/i.test(chain[i].name)) {
				out.heading = 1;
			}
		}
		if (!out.heading && out.extra.outline && out.extra.outline <= 6) {
			out.heading = out.extra.outline;
		}
		return out;
	};

	function Numbering(inXml) {
		this.abstracts = {};
		this.nums = {};
		this.counters = {};
		if (!inXml) {
			return;
		}
		var doc = X.parse(inXml);
		var abs = X.children(doc, "w:abstractNum");
		for (var i = 0; i < abs.length; i++) {
			var levels = {};
			var lvls = X.children(abs[i], "w:lvl");
			for (var l = 0; l < lvls.length; l++) {
				var lv = lvls[l];
				var extra = {};
				levels[+lv.a["w:ilvl"]] = {
					fmt: val(X.child(lv, "w:numFmt")) || "decimal",
					text: val(X.child(lv, "w:lvlText")),
					start: +(val(X.child(lv, "w:start")) || 1),
					para: paraCss(X.child(lv, "w:pPr"), {}, extra),
					run: runCss(X.child(lv, "w:rPr"), {})
				};
			}
			this.abstracts[abs[i].a["w:abstractNumId"]] = levels;
		}
		var nums = X.children(doc, "w:num");
		for (var n = 0; n < nums.length; n++) {
			var overrides = {};
			var ovs = X.children(nums[n], "w:lvlOverride");
			for (var o = 0; o < ovs.length; o++) {
				var so = X.child(ovs[o], "w:startOverride");
				if (so) {
					overrides[+ovs[o].a["w:ilvl"]] = +val(so);
				}
			}
			this.nums[nums[n].a["w:numId"]] = {abs: val(X.child(nums[n], "w:abstractNumId")), overrides: overrides};
		}
	}

	function formatNumber(inN, inFmt) {
		switch (inFmt) {
		case "lowerLetter":
			return Pptx.numberLabel(inN, "alphaLcPeriod").slice(0, -1);
		case "upperLetter":
			return Pptx.numberLabel(inN, "alphaUcPeriod").slice(0, -1);
		case "lowerRoman":
			return Pptx.numberLabel(inN, "romanLcPeriod").slice(0, -1);
		case "upperRoman":
			return Pptx.numberLabel(inN, "romanUcPeriod").slice(0, -1);
		case "decimalZero":
			return (inN < 10 ? "0" : "") + inN;
		default:
			return String(inN);
		}
	}

	/**
		The label of the next paragraph in list inNumId at level inLevel
		("3.", "b)", "1.2.", a bullet), and the level's indents.
	*/
	Numbering.prototype.next = function(inNumId, inLevel) {
		var num = this.nums[inNumId];
		if (!num || inNumId === "0") {
			return null;
		}
		var levels = this.abstracts[num.abs];
		var lvl = levels && levels[inLevel];
		if (!lvl) {
			return null;
		}
		// Counters are per abstract list, as Word continues a list across its nums.
		var key = num.abs;
		var counters = this.counters[key] || (this.counters[key] = {});
		var start = num.overrides.hasOwnProperty(inLevel) ? num.overrides[inLevel] : lvl.start;
		if (num.overrides.hasOwnProperty(inLevel) && !this.counters[key + ":" + inNumId + ":" + inLevel]) {
			this.counters[key + ":" + inNumId + ":" + inLevel] = true;
			counters[inLevel] = undefined;
		}
		counters[inLevel] = counters[inLevel] === undefined ? start : counters[inLevel] + 1;
		// Deeper levels start again.
		for (var d = inLevel + 1; d < 10; d++) {
			counters[d] = undefined;
		}
		var label;
		if (lvl.fmt === "bullet") {
			label = lvl.text && lvl.text.length === 1 && !/[-]/.test(lvl.text) && !/^[§·lØü]$/.test(lvl.text) ?
				lvl.text : "•";
			if (lvl.text === "·" || lvl.text === "") {
				label = "•";
			}
		} else if (lvl.fmt === "none") {
			label = "";
		} else {
			label = (lvl.text || "%" + (inLevel + 1) + ".").replace(/%(\d)/g, function(m, d) {
				var li = +d - 1;
				var lv = levels[li];
				var c = counters[li] === undefined ? (lv ? lv.start : 1) : counters[li];
				return formatNumber(c, lv ? lv.fmt : "decimal");
			});
		}
		return {label: label, para: lvl.para, run: lvl.run};
	};

	// --- Document --------------------------------------------------------

	function Builder(inFiles) {
		this.files = inFiles;
		this.styles = new Styles(inFiles["word/styles.xml"]);
		this.numbering = new Numbering(inFiles["word/numbering.xml"]);
		this.rels = this.readRels("word/document.xml");
		this.images = [];
		this.imageSet = {};
		this.notes = [];
		this.noteBodies = {footnote: this.readNotes("word/footnotes.xml", "w:footnote"),
			endnote: this.readNotes("word/endnotes.xml", "w:endnote")};
	}

	Builder.prototype.readRels = function(inPart) {
		var out = {};
		var dir = inPart.slice(0, inPart.lastIndexOf("/") + 1);
		var text = this.files[dir + "_rels/" + inPart.slice(dir.length) + ".rels"];
		if (!text) {
			return out;
		}
		var list = X.children(X.parse(text));
		for (var i = 0; i < list.length; i++) {
			var a = list[i].a;
			out[a.Id] = {target: a.TargetMode === "External" ? a.Target : Pptx.resolvePath(inPart, a.Target),
				external: a.TargetMode === "External"};
		}
		return out;
	};

	Builder.prototype.readNotes = function(inPart, inTag) {
		var out = {};
		if (!this.files[inPart]) {
			return out;
		}
		var list = X.children(X.parse(this.files[inPart]), inTag);
		for (var i = 0; i < list.length; i++) {
			out[list[i].a["w:id"]] = list[i];
		}
		return out;
	};

	Builder.prototype.useImage = function(inRid) {
		var rel = this.rels[inRid];
		if (!rel || rel.external) {
			return null;
		}
		if (!this.imageSet[rel.target]) {
			this.imageSet[rel.target] = true;
			this.images.push(rel.target);
		}
		return rel.target;
	};

	Builder.prototype.blocks = function(inParent, inOut) {
		var out = inOut || [];
		var kids = X.children(inParent);
		for (var i = 0; i < kids.length; i++) {
			var k = kids[i];
			if (k.n === "w:p") {
				this.paragraph(k, out);
			} else if (k.n === "w:tbl") {
				out.push(this.table(k));
			} else if (k.n === "w:sdt") {
				this.blocks(X.child(k, "w:sdtContent"), out);
			} else if (k.n === "w:customXml" || k.n === "w:ins" || k.n === "w:smartTag") {
				this.blocks(k, out);
			}
		}
		return out;
	};

	Builder.prototype.paragraph = function(inP, inOut) {
		var pPr = X.child(inP, "w:pPr");
		var styleId = val(X.child(pPr, "w:pStyle")) || this.styles.defaultPara;
		var st = this.styles.resolve(styleId);
		var extra = merge({}, st.extra);
		var own = paraCss(pPr, {}, extra);
		var para = merge(merge(this.styles.defaultPPr, st.para), own);
		var run = merge(merge(this.styles.defaultRun, st.run), {});
		var block = {t: "p", h: st.heading, s: null, r: []};
		if (extra.pageBreak && inOut.length) {
			inOut.push({t: "pagebreak"});
		}
		if (extra.numId !== undefined) {
			var n = this.numbering.next(extra.numId, extra.ilvl || 0);
			if (n) {
				// The level's indents, unless the paragraph sets its own.
				para = merge(merge(merge(this.styles.defaultPPr, st.para), n.para), own);
				if (n.label) {
					block.bu = n.label;
				}
			}
		}
		this.runs(inP, block.r, run);
		block.s = finishCss(merge(run, para));
		if (block.s.textIndent === "-0pt") {
			delete block.s.textIndent;
		}
		stripInherited(block.r, block.s);
		if (block.h) {
			// Headings keep their own size; the renderer's h1..h6 defaults must not add to it.
			block.s.fontSize = block.s.fontSize || ["", "20pt", "16pt", "14pt", "12pt", "11pt", "11pt"][block.h];
		}
		if (block.s.marginLeft && block.bu && block.s.textIndent) {
			// Hanging indent for the label: pad by the indent, pull the first line back.
			block.s.paddingLeft = block.s.marginLeft;
			delete block.s.marginLeft;
		}
		var breaks = block.pageBreaks;
		delete block.pageBreaks;
		inOut.push(block);
		if (breaks) {
			inOut.push({t: "pagebreak"});
		}
	};

	//* Runs of a paragraph (or a hyperlink, field, insertion) into inOut.
	Builder.prototype.runs = function(inParent, inOut, inBase, inState) {
		var state = inState || {field: 0};
		var kids = X.children(inParent);
		for (var i = 0; i < kids.length; i++) {
			var k = kids[i];
			switch (k.n) {
			case "w:r":
				this.run(k, inOut, inBase, state);
				break;
			case "w:hyperlink":
				var rel = k.a["r:id"] ? this.rels[k.a["r:id"]] : null;
				var href = rel && rel.external ? rel.target : k.a["w:anchor"] ? "#" + k.a["w:anchor"] : "";
				var link = {t: "a", href: href, r: []};
				this.runs(k, link.r, inBase, state);
				inOut.push(link);
				break;
			case "w:fldSimple":
			case "w:ins":
			case "w:smartTag":
			case "w:customXml":
			case "w:sdt":
			case "w:sdtContent":
			case "w:dir":
			case "w:bdo":
				this.runs(k, inOut, inBase, state);
				break;
			case "w:bookmarkStart":
				if (k.a["w:name"] && !/^_GoBack$/.test(k.a["w:name"])) {
					inOut.push({t: "anchor", id: k.a["w:name"]});
				}
				break;
			}
		}
	};

	Builder.prototype.run = function(inR, inOut, inBase, inState) {
		var rPr = X.child(inR, "w:rPr");
		var css = inBase;
		var rStyle = val(X.child(rPr, "w:rStyle"));
		if (rStyle) {
			css = merge(css, this.styles.resolve(rStyle).run);
		}
		css = runCss(rPr, merge(css, {}));
		if (css.hidden) {
			return;
		}
		var finished = finishCss(css);
		var kids = X.children(inR);
		for (var i = 0; i < kids.length; i++) {
			var k = kids[i];
			switch (k.n) {
			case "w:fldChar":
				var type = k.a["w:fldCharType"];
				if (type === "begin") {
					inState.field++;
					inState.code = true;
				} else if (type === "separate") {
					inState.code = false;
				} else if (type === "end") {
					inState.field = Math.max(0, inState.field - 1);
					inState.code = false;
				}
				break;
			case "w:instrText":
				break;
			case "w:t":
				if (!inState.code) {
					this.text(inOut, X.textOf(k), finished);
				}
				break;
			case "w:tab":
				if (!inState.code) {
					inOut.push({t: "tab"});
				}
				break;
			case "w:br":
			case "w:cr":
				if (k.a["w:type"] === "page") {
					inOut.push({t: "br"});
				} else {
					inOut.push({t: "br"});
				}
				break;
			case "w:noBreakHyphen":
				this.text(inOut, "‑", finished);
				break;
			case "w:softHyphen":
				this.text(inOut, "­", finished);
				break;
			case "w:sym":
				// A symbol-font character cannot be shown in another font; mark its place.
				this.text(inOut, "□", finished);
				break;
			case "w:drawing":
				this.drawing(k, inOut);
				break;
			case "w:pict":
			case "w:object":
				var vimg = X.findAll(k, "v:imagedata")[0];
				if (vimg && vimg.a["r:id"]) {
					var src = this.useImage(vimg.a["r:id"]);
					if (src) {
						var shape = X.findAll(k, "v:shape")[0];
						var style = shape ? shape.a.style || "" : "";
						var w = /width:\s*([0-9.]+(pt|in|cm|mm|px))/.exec(style), h = /height:\s*([0-9.]+(pt|in|cm|mm|px))/.exec(style);
						inOut.push({t: "img", src: src, w: w ? w[1] : "", h: h ? h[1] : "", alt: ""});
					}
				} else {
					// A text box: its paragraphs, inline.
					var box = X.findAll(k, "w:txbxContent")[0];
					if (box) {
						inOut.push({t: "frame", s: {}, b: this.blocks(box)});
					}
				}
				break;
			case "w:footnoteReference":
			case "w:endnoteReference":
				var kind = k.n === "w:footnoteReference" ? "footnote" : "endnote";
				var body = this.noteBodies[kind][k.a["w:id"]];
				var label = String(this.notes.length + 1);
				this.notes.push({label: label, b: body ? this.blocks(body) : []});
				inOut.push({t: "note", label: label});
				break;
			case "w:lastRenderedPageBreak":
				break;
			}
		}
	};

	Builder.prototype.text = function(inOut, inText, inCss) {
		if (!inText) {
			return;
		}
		var last = inOut[inOut.length - 1];
		var empty = isEmpty(inCss);
		if (last && last.t === "text" && (empty ? !last.s : last.s && JSON.stringify(last.s) === JSON.stringify(inCss))) {
			last.v += inText;
			return;
		}
		var run = {t: "text", v: inText};
		if (!empty) {
			run.s = inCss;
		}
		inOut.push(run);
	};

	Builder.prototype.drawing = function(inDrawing, inOut) {
		var holder = X.child(inDrawing, "wp:inline") || X.child(inDrawing, "wp:anchor");
		if (!holder) {
			return;
		}
		var blip = X.findAll(holder, "a:blip")[0];
		var ext = X.child(holder, "wp:extent");
		var doc = X.child(holder, "wp:docPr");
		if (blip && blip.a["r:embed"]) {
			var src = this.useImage(blip.a["r:embed"]);
			if (src) {
				inOut.push({t: "img", src: src, w: ext ? Math.round(+ext.a.cx / 12700 * 10) / 10 + "pt" : "",
					h: ext ? Math.round(+ext.a.cy / 12700 * 10) / 10 + "pt" : "", alt: doc ? doc.a.descr || doc.a.title || "" : ""});
				return;
			}
		}
		// A shape with text (wps:txbx) or a chart: its words, or a mark of its place.
		var box = X.findAll(holder, "w:txbxContent")[0];
		if (box) {
			inOut.push({t: "frame", s: {}, b: this.blocks(box)});
		} else if (X.findAll(holder, "c:chart").length) {
			inOut.push({t: "text", v: "[chart]", s: {fontStyle: "italic"}});
		}
	};

	Builder.prototype.table = function(inTbl) {
		var self = this;
		var cols = X.children(X.child(inTbl, "w:tblGrid"), "w:gridCol").map(function(c) {
			return c.a["w:w"] ? twipsToPt(c.a["w:w"]) + "pt" : "";
		});
		var tblPr = X.child(inTbl, "w:tblPr");
		var tableStyle = val(X.child(tblPr, "w:tblStyle"));
		var tst = tableStyle ? this.styles.resolve(tableStyle) : null;
		var borders = X.child(tblPr, "w:tblBorders") || (tableStyle && this.styles.byId[tableStyle] ? null : null);
		var hasBorders = !!(borders && X.children(borders).some(function(b) {
			return b.a["w:val"] && b.a["w:val"] !== "nil" && b.a["w:val"] !== "none";
		})) || /grid|table/i.test(tableStyle || "");
		var rows = [];
		// Vertically merged cells: the open merge per grid column.
		var openMerge = {};
		var trs = X.children(inTbl, "w:tr");
		for (var r = 0; r < trs.length; r++) {
			var tr = trs[r];
			var header = !!X.path(tr, "w:trPr", "w:tblHeader");
			var cells = [];
			var col = 0;
			var gridBefore = val(X.path(tr, "w:trPr", "w:gridBefore"));
			if (gridBefore) {
				col += +gridBefore;
			}
			var tcs = X.children(tr, "w:tc");
			for (var c = 0; c < tcs.length; c++) {
				var tc = tcs[c];
				var tcPr = X.child(tc, "w:tcPr");
				var span = +(val(X.child(tcPr, "w:gridSpan")) || 1);
				var vMerge = X.child(tcPr, "w:vMerge");
				if (vMerge && val(vMerge) !== "restart" && openMerge[col]) {
					openMerge[col].rs++;
					col += span;
					continue;
				}
				var css = {};
				if (hasBorders) {
					css.border = "1px solid #888888";
				}
				var shd = X.child(tcPr, "w:shd");
				var fill = shd ? hex(shd.a["w:fill"]) : null;
				if (fill && fill !== "#ffffff") {
					css.backgroundColor = fill;
				}
				var va = val(X.child(tcPr, "w:vAlign"));
				if (va === "center" || va === "bottom") {
					css.verticalAlign = va === "center" ? "middle" : "bottom";
				}
				var cell = {cs: span, rs: 1, s: css, b: self.blocks(tc)};
				if (vMerge && val(vMerge) === "restart") {
					openMerge[col] = cell;
				} else {
					for (var g = col; g < col + span; g++) {
						delete openMerge[g];
					}
				}
				cells.push(cell);
				col += span;
			}
			rows.push({header: header, s: {}, cells: cells});
		}
		var tableCss = {borderCollapse: "collapse"};
		if (tst && tst.run && tst.run.fontSize) {
			tableCss.fontSize = tst.run.fontSize;
		}
		return {t: "table", s: tableCss, cols: cols, rows: rows};
	};

	var inheritedProps = {fontFamily: 1, fontSize: 1, color: 1, fontWeight: 1, fontStyle: 1};

	/**
		Drops run properties the paragraph already has (the paragraph's own
		font, size and colour are inherited), then joins runs left equal:
		fewer spans to build and lay out.
	*/
	function stripInherited(inRuns, inPara) {
		var out = [];
		for (var i = 0; i < inRuns.length; i++) {
			var r = inRuns[i];
			if (r.t === "a") {
				stripInherited(r.r, inPara);
			}
			if (r.s) {
				for (var k in r.s) {
					if (r.s.hasOwnProperty(k) && inheritedProps[k] && inPara[k] === r.s[k]) {
						delete r.s[k];
					}
				}
				if (isEmpty(r.s)) {
					delete r.s;
				}
			}
			var last = out[out.length - 1];
			if (r.t === "text" && last && last.t === "text" && JSON.stringify(last.s || null) === JSON.stringify(r.s || null)) {
				last.v += r.v;
			} else {
				out.push(r);
			}
		}
		inRuns.length = 0;
		for (var j = 0; j < out.length; j++) {
			inRuns.push(out[j]);
		}
	}

	function pageOf(inSectPr) {
		var sz = X.child(inSectPr, "w:pgSz"), mar = X.child(inSectPr, "w:pgMar");
		var pt = function(inEl, inName, inDefault) {
			return (inEl && inEl.a[inName] !== undefined ? twipsToPt(inEl.a[inName]) : inDefault) + "pt";
		};
		return {
			width: pt(sz, "w:w", 595), height: pt(sz, "w:h", 842),
			marginTop: pt(mar, "w:top", 72), marginBottom: pt(mar, "w:bottom", 72),
			marginLeft: pt(mar, "w:left", 72), marginRight: pt(mar, "w:right", 72)
		};
	}

	function parse(inFiles) {
		var docXml = inFiles["word/document.xml"];
		if (!docXml) {
			throw new Error("DOCX: no word/document.xml");
		}
		var b = new Builder(inFiles);
		var body = X.path(X.parse(docXml), "w:body");
		var blocks = b.blocks(body);
		if (b.notes.length) {
			blocks.push({t: "notes", items: b.notes});
		}
		return {
			kind: "docx",
			blocks: blocks,
			base: finishCss(b.styles.defaultRun),
			page: pageOf(X.child(body, "w:sectPr")),
			images: b.images,
			meta: {}
		};
	}

	var DocDocx = {
		parse: parse,
		Numbering: Numbering,
		Styles: Styles
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocDocx;
	} else {
		root.DocDocx = DocDocx;
	}
})(this);
