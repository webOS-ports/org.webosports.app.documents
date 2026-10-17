/**
	Works formulas out again while a workbook is edited, with HyperFormula
	(GPLv3): a changed cell gives new values to every formula that uses it,
	shown at once and written into the saved copy.

		var calc = DocFormulas.create(HyperFormula, XLSX, workbook);
		var changed = calc.set(0, 4, 2, "12.5");   // [{sheet, r, c, v, t, text}]
		calc.values()                              // {"sheet,r,c": {v, t}}, every value changed so far

	t is "n", "s", "b" or "e" (an error such as "#DIV/0!", in v). text is
	what the grid shows, the number formatted as the cell says.

	The engine is built on the first edit only: HyperFormula is large, and
	works out the whole workbook once when built.
*/
(function(root) {
	var Sheets = root.DocSheets || (typeof require === "function" ? require("./Sheets.js") : null);

	//* A SheetJS cell as HyperFormula takes it.
	function contentOf(inCell) {
		if (!inCell) {
			return null;
		}
		if (inCell.f) {
			return "=" + inCell.f;
		}
		if (inCell.t === "n" || inCell.t === "b") {
			return inCell.v;
		}
		if (inCell.t === "e") {
			return inCell.w || "#N/A";
		}
		if (inCell.v === undefined || inCell.v === null) {
			return null;
		}
		var s = String(inCell.v);
		// Text that would read as a formula or a number stays text.
		return /^[=+\-']/.test(s) || /^\s*[-+]?\d/.test(s) ? "'" + s : s;
	}

	function Calc(inHF, inXLSX, inWb) {
		this.XLSX = inXLSX;
		this.wb = inWb;
		this.changed = {};
		var data = {};
		for (var i = 0; i < inWb.SheetNames.length; i++) {
			var ws = inWb.Sheets[inWb.SheetNames[i]];
			var dense = (ws && ws["!data"]) || [];
			var rows = [];
			for (var r = 0; r < dense.length; r++) {
				var src = dense[r];
				var row = [];
				if (src) {
					for (var c = 0; c < src.length; c++) {
						row[c] = contentOf(src[c]);
					}
				}
				for (var k = 0; k < row.length; k++) {
					if (row[k] === undefined) {
						row[k] = null;
					}
				}
				rows.push(row);
			}
			data[inWb.SheetNames[i]] = rows;
		}
		this.hf = inHF.HyperFormula.buildFromSheets(data, {licenseKey: "gpl-v3", useArrayArithmetic: true});
		this.sheetIds = inWb.SheetNames.map(function(n) {
			return this.hf.getSheetId(n);
		}, this);
	}

	//* A value from HyperFormula as {v, t}.
	function plain(inValue) {
		if (inValue === null || inValue === undefined) {
			return {v: "", t: "s"};
		}
		if (typeof inValue === "number") {
			return {v: inValue, t: "n"};
		}
		if (typeof inValue === "boolean") {
			return {v: inValue, t: "b"};
		}
		if (typeof inValue === "object" && inValue.value !== undefined) {
			// A CellError: "#DIV/0!", "#NAME?" (a function HyperFormula does not know), ...
			return {v: String(inValue.value), t: "e", name: inValue.type === "NAME"};
		}
		return {v: String(inValue), t: "s"};
	}

	//* What the grid shows for a value in a cell, by the cell's number format.
	Calc.prototype.text = function(inSheet, inRow, inCol, inValue) {
		if (inValue.t === "b") {
			return inValue.v ? "TRUE" : "FALSE";
		}
		if (inValue.t !== "n") {
			return String(inValue.v);
		}
		var ws = this.wb.Sheets[this.wb.SheetNames[inSheet]];
		var row = ws && ws["!data"] && ws["!data"][inRow];
		var cell = row && row[inCol];
		var fmt = cell && cell.z || "General";
		try {
			return this.XLSX.SSF.format(fmt, inValue.v);
		} catch (e) {
			return String(inValue.v);
		}
	};

	/**
		Sets a cell from what was typed (as Sheets.parseInput reads it) and
		returns every cell whose value changed, the cell itself first.
	*/
	Calc.prototype.set = function(inSheet, inRow, inCol, inText) {
		var p = Sheets.parseInput(inText);
		var content = p.f ? "=" + p.f : p.t === "z" ? null : p.t === "s" ? contentOf({t: "s", v: p.v}) : p.v;
		var id = this.sheetIds[inSheet];
		var changes = this.hf.setCellContents({sheet: id, row: inRow, col: inCol}, [[content]]);
		var out = [];
		var self = this;
		var own = null;
		changes.forEach(function(ch) {
			if (!ch.address || ch.address.row === undefined) {
				return;
			}
			var s = self.sheetIds.indexOf(ch.address.sheet);
			if (s < 0) {
				return;
			}
			var v = plain(ch.newValue);
			var item = {sheet: s, r: ch.address.row, c: ch.address.col, v: v.v, t: v.t};
			item.text = self.text(s, item.r, item.c, v);
			self.changed[s + "," + item.r + "," + item.c] = v;
			if (s === inSheet && item.r === inRow && item.c === inCol) {
				own = item;
			} else {
				out.push(item);
			}
		});
		if (!own && p.f) {
			// A formula whose value did not change still has one to show.
			var v0 = plain(this.hf.getCellValue({sheet: id, row: inRow, col: inCol}));
			own = {sheet: inSheet, r: inRow, c: inCol, v: v0.v, t: v0.t, text: this.text(inSheet, inRow, inCol, v0)};
			this.changed[inSheet + "," + inRow + "," + inCol] = v0;
		}
		if (own) {
			out.unshift(own);
		}
		return out;
	};

	//* Every value changed so far: {"sheet,r,c": {v, t}}.
	Calc.prototype.values = function() {
		return this.changed;
	};

	Calc.prototype.destroy = function() {
		if (this.hf) {
			this.hf.destroy();
			this.hf = null;
		}
	};

	var DocFormulas = {
		create: function(inHF, inXLSX, inWb) {
			return new Calc(inHF, inXLSX, inWb);
		},
		contentOf: contentOf,
		plain: plain
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocFormulas;
	} else {
		root.DocFormulas = DocFormulas;
	}
})(this);
