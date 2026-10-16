/**
	Spreadsheets as the grid shows them: per sheet a dense array of the
	cells' display strings, their kinds (for alignment), formulas, merged
	ranges and column widths. Built from a SheetJS workbook (xlsx, xls,
	ods) or from CSV rows; edits go back into the workbook for "Save a
	copy".

		var model = DocSheets.fromWorkbook(XLSX, workbook);
		model.sheets[0]   // {name, rows, cols, text: [[...]], kinds: [[...]], formulas: {"r,c": "=A1+1"},
		                  //  merges: [{r, c, rs, cs}], colWidths: [px], rowHeights: {r: px}, truncated}
		DocSheets.setCell(XLSX, workbook, 0, 4, 2, "12.5");   // into the workbook
		var bytes = DocSheets.write(XLSX, workbook, "ods");     // Uint8Array/ArrayBuffer of the copy

	XLSX is passed in, so node tests can use the same bundled SheetJS file.
*/
(function(root) {
	//* Cells per sheet the grid takes; a larger sheet is cut at a row boundary and says so.
	var MAX_CELLS = 2000000;
	//* Columns past this are never shown (the format allows 16384).
	var MAX_COLS = 1024;

	function colName(inIndex) {
		var s = "";
		var n = inIndex + 1;
		while (n > 0) {
			var m = (n - 1) % 26;
			s = String.fromCharCode(65 + m) + s;
			n = Math.floor((n - 1) / 26);
		}
		return s;
	}

	function cellRef(inRow, inCol) {
		return colName(inCol) + (inRow + 1);
	}

	//* "B12" -> {r: 11, c: 1}; null when not a cell reference.
	function parseRef(inRef) {
		var m = /^\s*\$?([A-Za-z]{1,3})\$?([0-9]{1,7})\s*$/.exec(String(inRef || ""));
		if (!m) {
			return null;
		}
		var letters = m[1].toUpperCase();
		var c = 0;
		for (var i = 0; i < letters.length; i++) {
			c = c * 26 + (letters.charCodeAt(i) - 64);
		}
		return {r: +m[2] - 1, c: c - 1};
	}

	function displayOf(inXLSX, inCell) {
		if (!inCell) {
			return "";
		}
		if (inCell.w !== undefined && inCell.w !== null) {
			return String(inCell.w);
		}
		if (inCell.t === "e") {
			return inCell.w || "#ERR";
		}
		if (inCell.v === undefined || inCell.v === null) {
			return "";
		}
		if (inCell.t === "b") {
			return inCell.v ? "TRUE" : "FALSE";
		}
		try {
			return inXLSX.utils.format_cell(inCell);
		} catch (e) {
			return String(inCell.v);
		}
	}

	/**
		One SheetJS worksheet (dense or sparse) as grid data. Excel column
		widths come as wpx/wch; heights as hpx/hpt.
	*/
	function fromSheet(inXLSX, inName, inWs) {
		var ref = inWs["!ref"];
		var range = ref ? inXLSX.utils.decode_range(ref) : {s: {r: 0, c: 0}, e: {r: -1, c: -1}};
		var rows = range.e.r + 1, cols = Math.min(range.e.c + 1, MAX_COLS);
		var truncated = false;
		if (cols > 0 && rows * cols > MAX_CELLS) {
			rows = Math.floor(MAX_CELLS / cols);
			truncated = true;
		}
		var dense = inWs["!data"] || (Array.isArray(inWs) ? inWs : null);
		var text = new Array(rows);
		var kinds = new Array(rows);
		var formulas = {};
		var usedRows = 0, usedCols = 0;
		for (var r = 0; r < rows; r++) {
			var rowText = null, rowKinds = null;
			var src = dense ? dense[r] : null;
			for (var c = 0; c < cols; c++) {
				var cell = dense ? (src ? src[c] : null) : inWs[cellRef(r, c)];
				if (!cell) {
					continue;
				}
				var s = displayOf(inXLSX, cell);
				if (cell.f) {
					formulas[r + "," + c] = "=" + cell.f;
				}
				if (s === "" && !cell.f) {
					continue;
				}
				if (!rowText) {
					rowText = [];
					rowKinds = [];
				}
				rowText[c] = s;
				rowKinds[c] = cell.t === "n" || cell.t === "d" ? "n" : cell.t === "b" ? "b" : cell.t === "e" ? "e" : "s";
				usedRows = r + 1;
				usedCols = Math.max(usedCols, c + 1);
			}
			text[r] = rowText;
			kinds[r] = rowKinds;
		}
		var merges = (inWs["!merges"] || []).filter(function(m) {
			return m.s.r < rows && m.s.c < cols;
		}).map(function(m) {
			usedRows = Math.max(usedRows, Math.min(m.e.r + 1, rows));
			usedCols = Math.max(usedCols, Math.min(m.e.c + 1, cols));
			return {r: m.s.r, c: m.s.c, rs: m.e.r - m.s.r + 1, cs: m.e.c - m.s.c + 1};
		});
		// The range of a sheet often runs far past its content (formatted empty cells).
		text.length = usedRows;
		kinds.length = usedRows;
		var colWidths = [];
		var colInfo = inWs["!cols"] || [];
		for (var ci = 0; ci < usedCols; ci++) {
			var info = colInfo[ci];
			var w = 80;
			if (info) {
				if (info.hidden) {
					w = 0;
				} else if (info.wpx) {
					w = info.wpx;
				} else if (info.wch) {
					w = Math.round(info.wch * 7 + 5);
				} else if (info.width) {
					w = Math.round(info.width * 7 + 5);
				}
			}
			colWidths.push(Math.max(0, Math.min(600, w)));
		}
		var rowHeights = {};
		var rowInfo = inWs["!rows"] || [];
		for (var ri = 0; ri < Math.min(rowInfo.length, usedRows); ri++) {
			var rinfo = rowInfo[ri];
			if (rinfo) {
				var h = rinfo.hidden ? 0 : rinfo.hpx || (rinfo.hpt ? rinfo.hpt * 96 / 72 : 0);
				if (h || rinfo.hidden) {
					rowHeights[ri] = Math.max(0, Math.min(400, Math.round(h)));
				}
			}
		}
		return {
			name: inName, rows: usedRows, cols: usedCols, text: text, kinds: kinds, formulas: formulas,
			merges: merges, colWidths: colWidths, rowHeights: rowHeights, truncated: truncated
		};
	}

	function fromWorkbook(inXLSX, inWb) {
		var sheets = [];
		var hidden = (inWb.Workbook && inWb.Workbook.Sheets) || [];
		for (var i = 0; i < inWb.SheetNames.length; i++) {
			var name = inWb.SheetNames[i];
			var ws = inWb.Sheets[name];
			if (!ws) {
				continue;
			}
			var s = fromSheet(inXLSX, name, ws);
			s.hidden = !!(hidden[i] && hidden[i].Hidden);
			s.index = i;
			sheets.push(s);
		}
		return {sheets: sheets};
	}

	//* CSV rows as one sheet (cells that look like numbers are aligned as numbers).
	function fromRows(inRows, inName) {
		var cols = 0;
		var text = [], kinds = [];
		var rows = inRows.length;
		for (var r = 0; r < rows; r++) {
			var row = inRows[r];
			cols = Math.max(cols, Math.min(row.length, MAX_COLS));
		}
		var truncated = false;
		if (cols && rows * cols > MAX_CELLS) {
			rows = Math.floor(MAX_CELLS / cols);
			truncated = true;
		}
		for (var r2 = 0; r2 < rows; r2++) {
			var src = inRows[r2];
			var t = [], k = [];
			for (var c = 0; c < Math.min(src.length, MAX_COLS); c++) {
				t[c] = src[c];
				k[c] = isNumeric(src[c]) ? "n" : "s";
			}
			text.push(t);
			kinds.push(k);
		}
		var colWidths = [];
		for (var ci = 0; ci < cols; ci++) {
			colWidths.push(100);
		}
		return {sheets: [{name: inName || "Sheet1", rows: rows, cols: cols, text: text, kinds: kinds, formulas: {},
			merges: [], colWidths: colWidths, rowHeights: {}, truncated: truncated, index: 0}]};
	}

	function isNumeric(inText) {
		return /^\s*[-+]?(\d+([.,]\d+)?|\d{1,3}([,.]\d{3})+([.,]\d+)?)\s*%?\s*$/.test(String(inText || ""));
	}

	/**
		What a typed value becomes: {t: "n", v} for a plain number (point
		or comma as decimal separator), {t: "b"} for TRUE/FALSE, {f} for a
		formula ("=..."), {t: "s", v} otherwise, {t: "z"} for empty.
	*/
	function parseInput(inText) {
		var s = String(inText === undefined || inText === null ? "" : inText);
		if (s === "") {
			return {t: "z"};
		}
		if (s.charAt(0) === "=" && s.length > 1) {
			return {f: s.slice(1)};
		}
		if (/^\s*[-+]?\d+([.,]\d+)?([eE][-+]?\d+)?\s*$/.test(s)) {
			return {t: "n", v: parseFloat(s.replace(",", "."))};
		}
		if (/^(true|false)$/i.test(s.trim())) {
			return {t: "b", v: /^true$/i.test(s.trim())};
		}
		return {t: "s", v: s};
	}

	//* Writes one typed value into a SheetJS worksheet (dense or sparse), growing its range.
	function setCell(inXLSX, inWb, inSheetIndex, inRow, inCol, inText) {
		var name = inWb.SheetNames[inSheetIndex];
		var ws = inWb.Sheets[name];
		var p = parseInput(inText);
		var cell = null;
		if (p.f) {
			cell = {t: "n", f: p.f};
		} else if (p.t !== "z") {
			cell = {t: p.t, v: p.v};
		}
		var dense = ws["!data"];
		if (dense) {
			// SheetJS's xlsx writer expects every row of a dense sheet to exist, empty ones too.
			for (var r = 0; r <= inRow; r++) {
				if (!dense[r]) {
					dense[r] = [];
				}
			}
			if (cell) {
				dense[inRow][inCol] = cell;
			} else {
				delete dense[inRow][inCol];
			}
		} else {
			var ref = cellRef(inRow, inCol);
			if (cell) {
				ws[ref] = cell;
			} else {
				delete ws[ref];
			}
		}
		if (cell) {
			var range = ws["!ref"] ? inXLSX.utils.decode_range(ws["!ref"]) : {s: {r: inRow, c: inCol}, e: {r: inRow, c: inCol}};
			range.s.r = Math.min(range.s.r, inRow);
			range.s.c = Math.min(range.s.c, inCol);
			range.e.r = Math.max(range.e.r, inRow);
			range.e.c = Math.max(range.e.c, inCol);
			ws["!ref"] = inXLSX.utils.encode_range(range);
		}
		return {text: p.f ? "=" + p.f : p.t === "z" ? "" : p.t === "b" ? (p.v ? "TRUE" : "FALSE") : String(p.v),
			kind: p.f ? "s" : p.t === "n" ? "n" : p.t === "b" ? "b" : "s", formula: p.f ? "=" + p.f : null};
	}

	//* The bytes of a copy of the workbook in inFormat ("xlsx", "ods", "csv"); csv takes the first sheet.
	function write(inXLSX, inWb, inFormat, inSheetIndex) {
		var opts = {type: "array", bookType: inFormat, compression: true};
		if (inFormat === "csv") {
			opts.sheet = inWb.SheetNames[inSheetIndex || 0];
		}
		return inXLSX.write(inWb, opts);
	}

	var DocSheets = {
		MAX_CELLS: MAX_CELLS,
		colName: colName,
		cellRef: cellRef,
		parseRef: parseRef,
		fromSheet: fromSheet,
		fromWorkbook: fromWorkbook,
		fromRows: fromRows,
		isNumeric: isNumeric,
		parseInput: parseInput,
		setCell: setCell,
		write: write
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocSheets;
	} else {
		root.DocSheets = DocSheets;
	}
})(this);
