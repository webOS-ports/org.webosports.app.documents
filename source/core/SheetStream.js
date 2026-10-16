/**
	Reads large spreadsheets as a stream: the sheet's XML is inflated and
	parsed a chunk at a time and only each cell's display text is kept, so
	a workbook whose XML is tens of megabytes never exists as one string,
	nor as a full workbook object, in a phone's memory. The grid gets rows
	in batches while the rest is still being read.

	Used for .xlsx and .ods files too large for the full reader (SheetJS),
	which is kept for smaller files because it knows every format and can
	write an edited copy. Streamed workbooks are read-only.

		var p = new DocSheetStream.XlsxSheetParser({sst: sharedStrings, styles: styleInfo, format: fn,
			onRow: function(r, texts, kinds) {...}});
		p.feed(chunk); ... p.end();
		p.merges, p.colWidths, p.rowHeights

		var o = new DocSheetStream.OdsParser({onSheet, onRow, onSheetEnd});
		o.feed(chunk); ... o.end();

	Pure JavaScript, tested in node with chunks of fixture files.
*/
(function(root) {
	var X = root.DocXml || (typeof require === "function" ? require("./Xml.js") : null);
	var decode = X.decodeEntities;

	//* "AB12" -> column index 27 (0-based); the row part is ignored.
	function colOfRef(inRef) {
		var c = 0;
		for (var i = 0; i < inRef.length; i++) {
			var ch = inRef.charCodeAt(i);
			if (ch >= 65 && ch <= 90) {
				c = c * 26 + (ch - 64);
			} else if (ch >= 97 && ch <= 122) {
				c = c * 26 + (ch - 96);
			} else {
				break;
			}
		}
		return c - 1;
	}

	function rangeOf(inRef) {
		var parts = String(inRef).split(":");
		var a = parts[0], b = parts[1] || parts[0];
		var rowA = +a.replace(/^[A-Za-z$]+/, "").replace(/\$/g, ""), rowB = +b.replace(/^[A-Za-z$]+/, "").replace(/\$/g, "");
		return {r: rowA - 1, c: colOfRef(a.replace(/\$/g, "")), r1: rowB - 1, c1: colOfRef(b.replace(/\$/g, ""))};
	}

	function attr(inText, inName) {
		var re = new RegExp("\\s" + inName + "=\"([^\"]*)\"");
		var m = re.exec(inText);
		return m ? m[1] : null;
	}

	// --- Shared strings ----------------------------------------------------

	/**
		xl/sharedStrings.xml, streamed: every <si> becomes one string (the
		text of its runs; phonetic hints dropped).
	*/
	function SharedStrings() {
		this.list = [];
		this.buf = "";
	}

	SharedStrings.prototype.feed = function(inChunk) {
		var s = this.buf + inChunk;
		var pos = 0;
		for (;;) {
			var start = s.indexOf("<si", pos);
			if (start < 0) {
				break;
			}
			var end = s.indexOf("</si>", start);
			// <si/> is an empty string.
			var selfClose = s.indexOf("/>", start);
			var gt = s.indexOf(">", start);
			if (gt >= 0 && selfClose >= 0 && selfClose + 1 === gt && (end < 0 || selfClose < end)) {
				this.list.push("");
				pos = gt + 1;
				continue;
			}
			if (end < 0) {
				break;
			}
			this.list.push(siText(s.slice(start, end)));
			pos = end + 5;
		}
		this.buf = s.slice(pos);
	};

	SharedStrings.prototype.end = function() {
		this.buf = "";
		return this.list;
	};

	function siText(inSi) {
		// Drop phonetic runs, then join every <t>.
		var si = inSi.indexOf("<rPh") >= 0 ? inSi.replace(/<rPh[\s\S]*?<\/rPh>/g, "") : inSi;
		var out = "";
		var re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
		var m;
		while ((m = re.exec(si))) {
			out += m[1];
		}
		return decode(out);
	}

	// --- Number formats ----------------------------------------------------

	/**
		The number format of every cell style (xl/styles.xml): cellXfs[i] ->
		format code. Built-in format ids are left to the formatter.
	*/
	function readStyles(inXml) {
		var out = {numFmts: {}, xfFmt: []};
		if (!inXml) {
			return out;
		}
		var doc = X.parse(inXml);
		var fmts = X.children(X.child(doc, "x:numFmts"), "x:numFmt");
		for (var i = 0; i < fmts.length; i++) {
			out.numFmts[fmts[i].a.numFmtId] = fmts[i].a.formatCode;
		}
		var xfs = X.children(X.child(doc, "x:cellXfs"), "x:xf");
		for (var j = 0; j < xfs.length; j++) {
			var id = +(xfs[j].a.numFmtId || 0);
			out.xfFmt.push(out.numFmts.hasOwnProperty(id) ? out.numFmts[id] : id);
		}
		return out;
	}

	// --- xlsx sheet --------------------------------------------------------

	/**
		Parses a worksheet's XML fed in chunks. inOptions:
			sst      shared strings (array)
			styles   readStyles() result
			format   function(value, formatCodeOrId) -> text (SheetJS's SSF), optional
			maxCells stop after this many cells (truncated is set)
			onRow    function(rowIndex, texts[], kinds[]) for every row with content
	*/
	function XlsxSheetParser(inOptions) {
		this.o = inOptions;
		this.buf = "";
		this.cells = 0;
		this.truncated = false;
		this.done = false;
		this.merges = [];
		this.colWidths = [];
		this.rowHeights = {};
		this.maxCol = -1;
		this.lastRow = -1;
		this.headParsed = false;
		this.inTail = false;
		this.tail = "";
	}

	XlsxSheetParser.prototype.feed = function(inChunk) {
		if (this.inTail) {
			// After </sheetData>: merged cells and the like, kept whole (they are small).
			this.tail += inChunk;
			return;
		}
		if (this.done) {
			return;
		}
		var s = this.buf + inChunk;
		var pos = 0;
		if (!this.headParsed) {
			var sd = s.indexOf("<sheetData");
			if (sd < 0) {
				// Keep everything until <sheetData> shows up (the head is small).
				this.buf = s;
				return;
			}
			this.parseHead(s.slice(0, sd));
			this.headParsed = true;
			pos = sd;
		}
		for (;;) {
			var start = s.indexOf("<row", pos);
			var dataEnd = s.indexOf("</sheetData>", pos);
			if (dataEnd >= 0 && (start < 0 || dataEnd < start)) {
				this.inTail = true;
				this.tail = s.slice(dataEnd);
				this.buf = "";
				return;
			}
			if (start < 0) {
				break;
			}
			var gt = s.indexOf(">", start);
			if (gt < 0) {
				break;
			}
			if (s.charAt(gt - 1) === "/") {
				// <row .../>: an empty row (maybe with a height).
				this.rowAttrs(s.slice(start, gt));
				pos = gt + 1;
				continue;
			}
			var end = s.indexOf("</row>", gt);
			if (end < 0) {
				break;
			}
			this.parseRow(s.slice(start, gt + 1), s.slice(gt + 1, end));
			pos = end + 6;
			if (this.done) {
				break;
			}
		}
		this.buf = this.done ? "" : s.slice(pos);
	};

	XlsxSheetParser.prototype.parseTail = function(inText) {
		var re = /<mergeCell\s+ref="([^"]+)"/g;
		var m;
		while ((m = re.exec(inText))) {
			var r = rangeOf(m[1]);
			this.merges.push({r: r.r, c: r.c, rs: r.r1 - r.r + 1, cs: r.c1 - r.c + 1});
		}
	};

	XlsxSheetParser.prototype.parseHead = function(inHead) {
		var re = /<col\s([^>]*)\/?>/g;
		var m;
		while ((m = re.exec(inHead))) {
			var a = m[1];
			var min = +(attr(" " + a, "min") || 1), max = Math.min(+(attr(" " + a, "max") || min), 16384);
			var width = +(attr(" " + a, "width") || 0);
			var hidden = attr(" " + a, "hidden") === "1";
			var px = hidden ? 0 : width ? Math.round(width * 7 + 5) : 80;
			for (var c = min - 1; c < max && c < 1024; c++) {
				this.colWidths[c] = Math.min(600, px);
			}
		}
	};

	XlsxSheetParser.prototype.rowAttrs = function(inOpen) {
		var r = attr(inOpen, "r");
		var rowIndex = r ? +r - 1 : this.lastRow + 1;
		this.lastRow = rowIndex;
		var ht = attr(inOpen, "ht");
		var hidden = attr(inOpen, "hidden") === "1";
		if (hidden) {
			this.rowHeights[rowIndex] = 0;
		} else if (ht && attr(inOpen, "customHeight") === "1") {
			this.rowHeights[rowIndex] = Math.min(400, Math.round(+ht * 96 / 72));
		}
		return rowIndex;
	};

	var cellRe = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;

	XlsxSheetParser.prototype.parseRow = function(inOpen, inBody) {
		var rowIndex = this.rowAttrs(inOpen);
		var texts = null, kinds = null;
		var col = -1;
		cellRe.lastIndex = 0;
		var m;
		while ((m = cellRe.exec(inBody))) {
			var a = m[1];
			var ref = attr(a, "r");
			col = ref ? colOfRef(ref) : col + 1;
			if (col >= 1024) {
				continue;
			}
			var inner = m[3];
			if (!inner) {
				continue;
			}
			var t = attr(a, "t") || "n";
			var vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
			var v = vm ? vm[1] : null;
			var text, kind;
			if (t === "s") {
				text = v !== null ? this.o.sst[+v] : "";
				kind = "s";
			} else if (t === "inlineStr") {
				text = siText(inner);
				kind = "s";
			} else if (t === "str") {
				text = v !== null ? decode(v) : "";
				kind = "s";
			} else if (t === "b") {
				text = v === "1" ? "TRUE" : "FALSE";
				kind = "b";
			} else if (t === "e") {
				text = v !== null ? decode(v) : "#ERR";
				kind = "e";
			} else if (t === "d") {
				text = v !== null ? decode(v) : "";
				kind = "n";
			} else {
				if (v === null) {
					continue;
				}
				text = this.formatNumber(+v, +(attr(a, "s") || 0));
				kind = "n";
			}
			if (text === undefined || text === null || text === "") {
				continue;
			}
			if (!texts) {
				texts = [];
				kinds = [];
			}
			texts[col] = text;
			kinds[col] = kind;
			if (col > this.maxCol) {
				this.maxCol = col;
			}
			this.cells++;
		}
		if (texts) {
			this.o.onRow(rowIndex, texts, kinds);
		}
		if (this.o.maxCells && this.cells >= this.o.maxCells) {
			this.truncated = true;
			this.done = true;
		}
	};

	XlsxSheetParser.prototype.formatNumber = function(inValue, inStyle) {
		var fmt = this.o.styles && this.o.styles.xfFmt[inStyle];
		if (fmt === undefined || fmt === 0 || fmt === "General" || !this.o.format) {
			return generalNumber(inValue);
		}
		try {
			return this.o.format(fmt, inValue);
		} catch (e) {
			return generalNumber(inValue);
		}
	};

	XlsxSheetParser.prototype.end = function() {
		if (this.tail) {
			this.parseTail(this.tail);
			this.tail = "";
		}
		this.buf = "";
		this.done = true;
	};

	/**
		A number as the "General" format shows it: up to 11 significant
		digits, no trailing zeros, exponent for very large or small values.
	*/
	function generalNumber(inValue) {
		if (!isFinite(inValue)) {
			return String(inValue);
		}
		if (inValue === Math.floor(inValue) && Math.abs(inValue) < 1e11) {
			return String(inValue);
		}
		var abs = Math.abs(inValue);
		if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) {
			return inValue.toExponential(5).replace(/\.?0+e/, "E").replace("e", "E").replace(/E\+?(-?)(\d)$/, "E$1+0$2")
				.replace("E+-", "E-");
		}
		var s = inValue.toPrecision(11);
		if (s.indexOf(".") >= 0 && s.indexOf("e") < 0) {
			s = s.replace(/0+$/, "").replace(/\.$/, "");
		}
		return s;
	}

	// --- ods -----------------------------------------------------------

	/**
		Parses an ODS content.xml fed in chunks. ODF stores each cell's text
		as shown (formatted), so no number formats are needed. inOptions:
			maxCells, onSheet(name), onRow(row, texts, kinds), onSheetEnd(info)
		info: {merges, colWidths, rowHeights, truncated}
	*/
	function OdsParser(inOptions) {
		this.o = inOptions;
		this.buf = "";
		this.cells = 0;
		this.styles = null;
		this.sheet = null;
		this.done = false;
	}

	OdsParser.prototype.feed = function(inChunk) {
		if (this.done) {
			return;
		}
		var s = this.buf + inChunk;
		var pos = 0;
		if (!this.styles) {
			var body = s.indexOf("<office:body");
			if (body < 0) {
				this.buf = s;
				return;
			}
			this.readStyles(s.slice(0, body));
			pos = body;
		}
		for (;;) {
			if (!this.sheet) {
				var ts = s.indexOf("<table:table ", pos);
				if (ts < 0) {
					break;
				}
				var tgt = s.indexOf(">", ts);
				if (tgt < 0) {
					break;
				}
				this.startSheet(s.slice(ts, tgt + 1));
				pos = tgt + 1;
				continue;
			}
			var nextRow = s.indexOf("<table:table-row", pos);
			var nextCol = s.indexOf("<table:table-column", pos);
			var sheetEnd = s.indexOf("</table:table>", pos);
			// Column definitions come before the rows.
			if (nextCol >= 0 && (nextRow < 0 || nextCol < nextRow) && (sheetEnd < 0 || nextCol < sheetEnd)) {
				var cgt = s.indexOf(">", nextCol);
				if (cgt < 0) {
					break;
				}
				var tag = s.slice(nextCol, cgt + 1);
				if (/^<table:table-column[\s\/>]/.test(tag)) {
					this.column(tag);
				}
				pos = cgt + 1;
				continue;
			}
			if (sheetEnd >= 0 && (nextRow < 0 || sheetEnd < nextRow)) {
				this.endSheet();
				pos = sheetEnd + 14;
				continue;
			}
			if (nextRow < 0) {
				break;
			}
			var rgt = s.indexOf(">", nextRow);
			if (rgt < 0) {
				break;
			}
			var open = s.slice(nextRow, rgt + 1);
			if (!/^<table:table-row[\s\/>]/.test(open)) {
				// <table:table-rows> or <table:table-row-group>: a container, step in.
				pos = rgt + 1;
				continue;
			}
			if (open.charAt(open.length - 2) === "/") {
				this.row(open, "");
				pos = rgt + 1;
				continue;
			}
			var rend = s.indexOf("</table:table-row>", rgt);
			if (rend < 0) {
				break;
			}
			this.row(open, s.slice(rgt + 1, rend));
			pos = rend + 18;
			if (this.done) {
				break;
			}
		}
		this.buf = s.slice(pos);
	};

	OdsParser.prototype.readStyles = function(inHead) {
		this.styles = {cols: {}, rows: {}};
		var re = /<style:style\s([^>]*)>([\s\S]*?)<\/style:style>/g;
		var m;
		while ((m = re.exec(inHead))) {
			var name = attr(" " + m[1], "style:name");
			var family = attr(" " + m[1], "style:family");
			if (family === "table-column") {
				var w = attr(m[2], "style:column-width");
				if (w) {
					this.styles.cols[name] = w;
				}
			} else if (family === "table-row") {
				var h = attr(m[2], "style:row-height");
				var opt = attr(m[2], "style:use-optimal-row-height");
				if (h && opt !== "true") {
					this.styles.rows[name] = h;
				}
			}
		}
	};

	var perUnit = {px: 1, pt: 96 / 72, pc: 16, "in": 96, cm: 96 / 2.54, mm: 96 / 25.4};

	function toPx(inLength) {
		var m = /^\s*(-?[0-9.]+)\s*(px|pt|pc|in|cm|mm)?\s*$/.exec(String(inLength || ""));
		return m ? parseFloat(m[1]) * perUnit[m[2] || "px"] : NaN;
	}

	OdsParser.prototype.startSheet = function(inOpen) {
		this.sheet = {name: decode(attr(inOpen, "table:name") || ""), row: 0, colWidths: [], col: 0, merges: [], rowHeights: {},
			covered: {}, truncated: false, maxCol: -1};
		this.o.onSheet(this.sheet.name);
	};

	OdsParser.prototype.column = function(inTag) {
		var sh = this.sheet;
		var n = Math.min(+(attr(inTag, "table:number-columns-repeated") || 1), 1024);
		var style = attr(inTag, "table:style-name");
		var w = style && this.styles.cols[style] ? Math.round(toPx(this.styles.cols[style])) : 80;
		var hidden = attr(inTag, "table:visibility") === "collapse";
		for (var i = 0; i < n && sh.col < 1024; i++) {
			sh.colWidths[sh.col++] = hidden ? 0 : Math.min(600, w);
		}
	};

	var odsCellRe = /<table:(covered-table-cell|table-cell)\b([^>]*?)(\/>|>([\s\S]*?)<\/table:\1>)/g;

	OdsParser.prototype.row = function(inOpen, inBody) {
		var sh = this.sheet;
		var reps = +(attr(inOpen, "table:number-rows-repeated") || 1);
		var style = attr(inOpen, "table:style-name");
		var hidden = /table:visibility="(collapse|filter)"/.test(inOpen);
		var texts = null, kinds = null;
		var col = 0;
		odsCellRe.lastIndex = 0;
		var m;
		while ((m = odsCellRe.exec(inBody))) {
			var a = m[2];
			var n = +(attr(a, "table:number-columns-repeated") || 1);
			if (m[1] === "covered-table-cell") {
				col += n;
				continue;
			}
			var inner = m[4] || "";
			var text = inner ? odsText(inner) : "";
			var type = attr(a, "office:value-type");
			var cs = +(attr(a, "table:number-columns-spanned") || 1), rs = +(attr(a, "table:number-rows-spanned") || 1);
			if ((cs > 1 || rs > 1) && col < 1024) {
				sh.merges.push({r: sh.row, c: col, rs: rs, cs: cs});
			}
			if (text !== "") {
				// A repeated non-empty cell is rare and never wide; cap it anyway.
				var count = Math.min(n, 1024 - col);
				for (var k = 0; k < count; k++) {
					if (!texts) {
						texts = [];
						kinds = [];
					}
					texts[col + k] = text;
					kinds[col + k] = type === "float" || type === "percentage" || type === "currency" || type === "date" ||
						type === "time" ? "n" : type === "boolean" ? "b" : "s";
					this.cells++;
				}
				if (col + count - 1 > sh.maxCol) {
					sh.maxCol = col + count - 1;
				}
			}
			col += n;
		}
		if (texts) {
			// Rows of content repeat only a few times in practice (a copied row).
			var copies = Math.min(reps, 1000);
			for (var i = 0; i < copies; i++) {
				this.o.onRow(sh.row + i, texts, kinds);
			}
			this.cells += (copies - 1) * texts.length;
			if (style && this.styles.rows[style]) {
				for (var j = 0; j < copies; j++) {
					sh.rowHeights[sh.row + j] = Math.min(400, Math.round(toPx(this.styles.rows[style])));
				}
			}
			if (hidden) {
				for (var hi = 0; hi < copies; hi++) {
					sh.rowHeights[sh.row + hi] = 0;
				}
			}
		}
		sh.row += reps;
		if (this.o.maxCells && this.cells >= this.o.maxCells) {
			sh.truncated = true;
			this.endSheet();
			this.done = true;
		}
	};

	function odsText(inInner) {
		// The text:p paragraphs of a cell, one per line; annotations dropped.
		var body = inInner.indexOf("<office:annotation") >= 0 ? inInner.replace(/<office:annotation[\s\S]*?<\/office:annotation>/g, "") : inInner;
		var paras = [];
		var re = /<text:p(?:\s[^>]*)?>([\s\S]*?)<\/text:p>|<text:p[^>]*\/>/g;
		var m;
		while ((m = re.exec(body))) {
			var t = m[1] || "";
			t = t.replace(/<text:s(?:\s+text:c="(\d+)")?\s*\/>/g, function(x, n) {
				return new Array((+n || 1) + 1).join(" ");
			}).replace(/<text:tab\s*\/>/g, "\t").replace(/<text:line-break\s*\/>/g, "\n").replace(/<[^>]+>/g, "");
			paras.push(decode(t));
		}
		return paras.join("\n");
	}

	OdsParser.prototype.endSheet = function() {
		var sh = this.sheet;
		if (!sh) {
			return;
		}
		this.sheet = null;
		this.o.onSheetEnd({merges: sh.merges, colWidths: sh.colWidths, rowHeights: sh.rowHeights, truncated: sh.truncated,
			maxCol: sh.maxCol});
	};

	OdsParser.prototype.end = function() {
		this.endSheet();
		this.buf = "";
		this.done = true;
	};

	var DocSheetStream = {
		SharedStrings: SharedStrings,
		XlsxSheetParser: XlsxSheetParser,
		OdsParser: OdsParser,
		readStyles: readStyles,
		generalNumber: generalNumber,
		colOfRef: colOfRef,
		rangeOf: rangeOf,
		odsText: odsText
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocSheetStream;
	} else {
		root.DocSheetStream = DocSheetStream;
	}
})(this);
