/**
	A spreadsheet grid that only builds the cells near the screen, so a
	sheet of a million rows scrolls like one of ten. Column letters and row
	numbers stay in place, and the first row can be frozen under the
	letters as a header.

		{kind: "SheetGrid", onSelect: "cellSelected"}
		grid.setSheet(model.sheets[0]);     // DocSheets sheet model
		grid.select(4, 2);                   // row, column (0-based)
		grid.scrollToCell(1000, 0);

	The scroller is enyo's, for its flick and momentum; the cells live in a
	"window" of rows and columns around the visible part, rebuilt (as one
	innerHTML string, every text escaped) when scrolling leaves it.
*/
enyo.kind({
	name: "SheetGrid",
	kind: enyo.Control,
	className: "sheet-grid",
	published: {
		zoom: 1,
		freezeHeader: true
	},
	events: {
		onSelect: "",
		onActivate: ""
	},
	//* Geometry at zoom 1, in CSS px.
	rowHeight: 24,
	headerHeight: 22,
	headerWidth: 44,
	extraRows: 20,
	extraCols: 5,
	maxRows: 1000000,
	components: [
		{name: "corner", className: "sheet-corner"},
		{name: "colHead", className: "sheet-colhead", components: [
			{name: "colInner", className: "sheet-colinner"}
		]},
		{name: "rowHead", className: "sheet-rowhead", components: [
			{name: "rowInner", className: "sheet-rowinner"}
		]},
		{name: "scroller", kind: "Scroller", className: "sheet-body", autoHorizontal: false, horizontal: true,
			onScroll: "scrolled", onScrollStop: "scrolled", components: [
			{name: "space", className: "sheet-space", onclick: "spaceClick", components: [
				{name: "cells", className: "sheet-cells"},
				{name: "selection", className: "sheet-selection"}
			]}
		]}
	],

	create: function() {
		this.inherited(arguments);
		this.sheet = null;
		this.sel = null;
		this.win = null;
		this.highlight = null;
	},

	rendered: function() {
		this.inherited(arguments);
		if (this.sheet) {
			this.measure();
		}
	},

	//* inSheet: a DocSheets sheet ({rows, cols, text, kinds, merges, colWidths, rowHeights}).
	setSheet: function(inSheet) {
		this.sheet = inSheet;
		this.sel = null;
		this.highlight = null;
		this.prepareMerges();
		if (this.hasNode()) {
			this.measure();
			this.$.scroller.setScrollPositionDirect(0, 0);
			this.scrolled();
		}
	},

	zoomChanged: function() {
		if (this.sheet && this.hasNode()) {
			var s = this.$.scroller;
			var oldZoom = this.lastZoom || 1;
			var left = s.scrollLeft / oldZoom, top = s.scrollTop / oldZoom;
			this.measure();
			this.setScroll(left * this.zoom, top * this.zoom);
		}
	},

	freezeHeaderChanged: function() {
		if (this.sheet && this.hasNode()) {
			this.measure();
			this.scrolled();
		}
	},

	prepareMerges: function() {
		this.covered = {};
		this.mergeAt = {};
		var merges = (this.sheet && this.sheet.merges) || [];
		for (var i = 0; i < merges.length; i++) {
			var m = merges[i];
			this.mergeAt[m.r + "," + m.c] = m;
			// Big merges are rare; cap the bookkeeping anyway.
			if (m.rs * m.cs > 100000) {
				continue;
			}
			for (var r = m.r; r < m.r + m.rs; r++) {
				for (var c = m.c; c < m.c + m.cs; c++) {
					if (r !== m.r || c !== m.c) {
						this.covered[r + "," + c] = m;
					}
				}
			}
		}
	},

	// --- Geometry --------------------------------------------------------

	measure: function() {
		var s = this.sheet;
		var z = this.zoom;
		this.lastZoom = z;
		this.nRows = Math.min(s.rows + this.extraRows, this.maxRows);
		//* Rows that held content when measured (the rest are the empty rows after them).
		this.drawnRows = s.rows;
		this.nCols = s.cols + this.extraCols;
		var colX = this.colX = new Float64Array(this.nCols + 1);
		for (var c = 0; c < this.nCols; c++) {
			var w = c < s.colWidths.length ? s.colWidths[c] : 80;
			colX[c + 1] = colX[c] + Math.round(w * z);
		}
		var first = this.bodyFirst();
		var rowY = this.rowY = new Float64Array(this.nRows - first + 1);
		var rh = Math.round(this.rowHeight * z);
		var custom = s.rowHeights || {};
		for (var r = first; r < this.nRows; r++) {
			var h = custom.hasOwnProperty(r) ? Math.round(custom[r] * z) : rh;
			rowY[r - first + 1] = rowY[r - first] + h;
		}
		this.frozenHeight = this.freezeHeader && s.rows > 1 ? (custom.hasOwnProperty(0) ? Math.round(custom[0] * z) : rh) : 0;
		var hh = Math.round(this.headerHeight * Math.min(z, 1.5));
		var hw = Math.round(this.headerWidth * Math.min(z, 1.5) + (String(this.nRows).length > 5 ? 12 : 0));
		this.headTop = hh + this.frozenHeight;
		this.headLeft = hw;
		this.applyBox(this.$.corner, 0, 0, hw, this.headTop);
		this.$.colHead.applyStyle("left", hw + "px");
		this.$.colHead.applyStyle("height", this.headTop + "px");
		this.$.rowHead.applyStyle("top", this.headTop + "px");
		this.$.rowHead.applyStyle("width", hw + "px");
		this.$.scroller.applyStyle("left", hw + "px");
		this.$.scroller.applyStyle("top", this.headTop + "px");
		var space = this.$.space.hasNode();
		if (space) {
			space.style.width = colX[this.nCols] + "px";
			space.style.height = rowY[rowY.length - 1] + "px";
			space.style.fontSize = Math.round(13 * z) + "px";
		}
		var ch = this.$.colHead.hasNode(), rhd = this.$.rowHead.hasNode();
		if (ch) {
			ch.style.fontSize = Math.round(12 * Math.min(z, 1.5)) + "px";
		}
		if (rhd) {
			rhd.style.fontSize = Math.round(12 * Math.min(z, 1.5)) + "px";
		}
		this.win = null;
		this.headerWin = null;
	},

	applyBox: function(inControl, inLeft, inTop, inWidth, inHeight) {
		inControl.applyStyle("left", inLeft + "px");
		inControl.applyStyle("top", inTop + "px");
		inControl.applyStyle("width", inWidth + "px");
		inControl.applyStyle("height", inHeight + "px");
	},

	bodyFirst: function() {
		return this.freezeHeader && this.sheet && this.sheet.rows > 1 ? 1 : 0;
	},

	//* Index of the last prefix entry <= inValue.
	find: function(inPrefix, inValue) {
		var lo = 0, hi = inPrefix.length - 2;
		if (hi < 0) {
			return 0;
		}
		while (lo < hi) {
			var mid = (lo + hi + 1) >> 1;
			if (inPrefix[mid] <= inValue) {
				lo = mid;
			} else {
				hi = mid - 1;
			}
		}
		return lo;
	},

	/**
		The visible part. Its size is measured once and again only on a
		resize: reading it on every scroll step would force the engine to lay
		out the cells just written, which costs half a second a time on a
		slow device.
	*/
	viewport: function() {
		var s = this.$.scroller;
		if (!this.viewSize || !this.viewSize.width) {
			var n = s.hasNode();
			this.viewSize = {width: n ? n.clientWidth : 0, height: n ? n.clientHeight : 0};
		}
		return {
			left: Math.max(0, s.scrollLeft || 0), top: Math.max(0, s.scrollTop || 0),
			width: this.viewSize.width, height: this.viewSize.height
		};
	},

	// --- Drawing ---------------------------------------------------------

	scrolled: function() {
		if (!this.sheet || !this.colX) {
			return;
		}
		var v = this.viewport();
		var s = this.$.scroller;
		var colInner = this.$.colInner.hasNode(), rowInner = this.$.rowInner.hasNode();
		if (colInner) {
			colInner.style.transform = "translate3d(" + (-(s.scrollLeft || 0)) + "px,0,0)";
		}
		if (rowInner) {
			rowInner.style.transform = "translate3d(0," + (-(s.scrollTop || 0)) + "px,0)";
		}
		var first = this.bodyFirst();
		var r0 = this.find(this.rowY, v.top) + first, r1 = this.find(this.rowY, v.top + v.height) + first;
		var c0 = this.find(this.colX, v.left), c1 = this.find(this.colX, v.left + v.width);
		var w = this.win;
		if (!w || r0 < w.r0 || r1 > w.r1 || c0 < w.c0 || c1 > w.c1) {
			// A margin around the visible cells, so short scrolls need no rebuild; small,
			// as every cell built is laid out (slow on old devices).
			var overR = Math.max(8, Math.ceil((r1 - r0) / 2)), overC = 2;
			this.win = {
				r0: Math.max(first, r0 - overR), r1: Math.min(this.nRows - 1, r1 + overR),
				c0: Math.max(0, c0 - overC), c1: Math.min(this.nCols - 1, c1 + overC)
			};
			this.drawBody();
			this.drawHeaders();
		}
	},

	escape: function(inText) {
		return String(inText).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
	},

	rowTop: function(inRow) {
		return this.rowY[inRow - this.bodyFirst()];
	},

	rowBottom: function(inRow) {
		return this.rowY[inRow - this.bodyFirst() + 1];
	},

	drawBody: function() {
		var node = this.$.cells.hasNode();
		if (!node) {
			return;
		}
		var s = this.sheet, w = this.win, colX = this.colX;
		var html = [];
		var top = this.rowTop(w.r0), bottom = this.rowBottom(w.r1);
		var left = colX[w.c0], right = colX[w.c1 + 1];
		// Grid lines: one per row and column in the window.
		for (var r = w.r0; r <= w.r1 + 1 && r - this.bodyFirst() < this.rowY.length; r++) {
			html.push("<div class=\"gl-h\" style=\"top:" + (this.rowTop(r) - 1) + "px;left:" + left + "px;width:" + (right - left) + "px\"></div>");
		}
		for (var c = w.c0; c <= w.c1 + 1 && c < colX.length; c++) {
			html.push("<div class=\"gl-v\" style=\"left:" + (colX[c] - 1) + "px;top:" + top + "px;height:" + (bottom - top) + "px\"></div>");
		}
		var drawnMerges = {};
		for (var r2 = w.r0; r2 <= w.r1; r2++) {
			var rowText = s.text[r2], rowKinds = s.kinds[r2];
			for (var c2 = w.c0; c2 <= w.c1; c2++) {
				var key = r2 + "," + c2;
				var m = this.covered[key];
				if (m) {
					// A merge that starts outside the window is drawn from its first visible cell.
					var mk = m.r + "," + m.c;
					if (!drawnMerges[mk] && (m.r < w.r0 || m.c < w.c0)) {
						drawnMerges[mk] = true;
						html.push(this.cellHtml(m.r, m.c, m));
					}
					continue;
				}
				var merge = this.mergeAt[key];
				if (merge) {
					drawnMerges[key] = true;
					html.push(this.cellHtml(r2, c2, merge));
					continue;
				}
				var t = rowText ? rowText[c2] : null;
				if (t === undefined || t === null || t === "") {
					continue;
				}
				html.push(this.cellHtml(r2, c2, null, t, rowKinds && rowKinds[c2]));
			}
		}
		if (this.highlight) {
			html.push(this.highlightHtml());
		}
		node.innerHTML = html.join("");
		this.placeSelection();
	},

	cellHtml: function(inRow, inCol, inMerge, inText, inKind) {
		var s = this.sheet;
		var text = inText;
		var kind = inKind;
		if (text === undefined) {
			text = s.text[inRow] ? s.text[inRow][inCol] : "";
			kind = s.kinds[inRow] ? s.kinds[inRow][inCol] : "s";
		}
		var r1 = inMerge ? Math.min(inRow + inMerge.rs - 1, this.nRows - 1) : inRow;
		var c1 = inMerge ? Math.min(inCol + inMerge.cs - 1, this.nCols - 1) : inCol;
		var top = this.rowTop(inRow), left = this.colX[inCol];
		var height = this.rowBottom(r1) - top, width = this.colX[c1 + 1] - left;
		var cls = "c" + (kind === "n" ? " n" : kind === "b" || kind === "e" ? " b" : "") + (inMerge ? " m" : "");
		return "<div class=\"" + cls + "\" style=\"left:" + left + "px;top:" + top + "px;width:" + (width - 1) +
			"px;height:" + (height - 1) + "px;line-height:" + (height - 1) + "px\">" +
			(text === undefined || text === null ? "" : this.escape(text)) + "</div>";
	},

	drawHeaders: function() {
		var w = this.win, colX = this.colX;
		var cols = [], rows = [];
		var hh = this.headTop - this.frozenHeight;
		for (var c = w.c0; c <= w.c1; c++) {
			cols.push("<div class=\"ch\" style=\"left:" + colX[c] + "px;width:" + (colX[c + 1] - colX[c] - 1) + "px;height:" + hh +
				"px;line-height:" + hh + "px\">" + DocSheets.colName(c) + "</div>");
		}
		if (this.frozenHeight) {
			// The frozen first row, under the letters.
			var s = this.sheet;
			for (var c2 = w.c0; c2 <= w.c1; c2++) {
				var m = this.mergeAt["0," + c2];
				if (this.covered["0," + c2]) {
					continue;
				}
				var c3 = m ? Math.min(c2 + m.cs - 1, this.nCols - 1) : c2;
				var t = s.text[0] ? s.text[0][c2] : "";
				var k = s.kinds[0] ? s.kinds[0][c2] : "s";
				cols.push("<div class=\"fz" + (k === "n" ? " n" : "") + "\" style=\"top:" + hh + "px;left:" + colX[c2] + "px;width:" +
					(colX[c3 + 1] - colX[c2] - 1) + "px;height:" + (this.frozenHeight - 1) + "px;line-height:" +
					(this.frozenHeight - 1) + "px\">" + (t === undefined || t === null ? "" : this.escape(t)) + "</div>");
			}
		}
		for (var r = w.r0; r <= w.r1; r++) {
			var top = this.rowTop(r), h = this.rowBottom(r) - top;
			rows.push("<div class=\"rh\" style=\"top:" + top + "px;height:" + (h - 1) + "px;line-height:" + (h - 1) + "px\">" +
				(r + 1) + "</div>");
		}
		var colInner = this.$.colInner.hasNode(), rowInner = this.$.rowInner.hasNode(), corner = this.$.corner.hasNode();
		if (colInner) {
			colInner.innerHTML = cols.join("");
		}
		if (rowInner) {
			rowInner.innerHTML = rows.join("");
		}
		if (corner) {
			corner.innerHTML = this.frozenHeight ? "<div class=\"fz-label\" style=\"top:" + hh + "px;height:" + (this.frozenHeight - 1) +
				"px;line-height:" + (this.frozenHeight - 1) + "px\">1</div>" : "";
		}
	},

	// --- Selection -------------------------------------------------------

	spaceClick: function(inSender, inEvent) {
		if (!this.sheet || this.$.scroller.isScrolling && this.$.scroller.isScrolling() && this.wasDragging) {
			return;
		}
		var node = this.$.space.hasNode();
		var rect = node.getBoundingClientRect();
		var x = inEvent.clientX - rect.left, y = inEvent.clientY - rect.top;
		var c = this.find(this.colX, x);
		var r = this.find(this.rowY, y) + this.bodyFirst();
		var m = this.covered[r + "," + c];
		if (m) {
			r = m.r;
			c = m.c;
		}
		var again = this.sel && this.sel.r === r && this.sel.c === c;
		this.select(r, c);
		if (again) {
			this.doActivate(r, c);
		}
		return true;
	},

	//* Clicks on the frozen header row (in the column header strip).
	colHeadClick: function(inEvent) {
		var node = this.$.colInner.hasNode();
		var rect = node.getBoundingClientRect();
		var x = inEvent.clientX - rect.left, y = inEvent.clientY - rect.top;
		if (!this.frozenHeight || y < this.headTop - this.frozenHeight) {
			return;
		}
		var c = this.find(this.colX, x);
		var again = this.sel && this.sel.r === 0 && this.sel.c === c;
		this.select(0, c);
		if (again) {
			this.doActivate(0, c);
		}
	},

	//* Clicks that bubble up to the grid itself: only the header strip's are ours.
	clickHandler: function(inSender, inEvent) {
		var ch = this.$.colHead.hasNode();
		if (ch && ch.contains(inEvent.target)) {
			this.colHeadClick(inEvent);
			return true;
		}
		return false;
	},

	select: function(inRow, inCol) {
		this.sel = {r: inRow, c: inCol};
		this.placeSelection();
		this.doSelect(inRow, inCol);
	},

	placeSelection: function() {
		var n = this.$.selection.hasNode();
		if (!n) {
			return;
		}
		if (!this.sel || !this.colX) {
			n.style.display = "none";
			return;
		}
		var r = this.sel.r, c = this.sel.c;
		var m = this.mergeAt[r + "," + c];
		var c1 = m ? Math.min(c + m.cs - 1, this.nCols - 1) : c;
		var r1 = m ? Math.min(r + m.rs - 1, this.nRows - 1) : r;
		if (r < this.bodyFirst()) {
			// The frozen row: the selection is drawn in the header strip.
			n.style.display = "none";
			this.markFrozen(c, c1);
			return;
		}
		this.markFrozen(-1);
		n.style.display = "block";
		n.style.left = (this.colX[c] - 1) + "px";
		n.style.top = (this.rowTop(r) - 1) + "px";
		n.style.width = (this.colX[c1 + 1] - this.colX[c] - 1) + "px";
		n.style.height = (this.rowBottom(r1) - this.rowTop(r) - 1) + "px";
	},

	markFrozen: function(inCol, inCol1) {
		var inner = this.$.colInner.hasNode();
		if (!inner) {
			return;
		}
		var old = inner.querySelector(".sheet-frozen-sel");
		if (old) {
			old.parentNode.removeChild(old);
		}
		if (inCol < 0) {
			return;
		}
		var d = document.createElement("div");
		d.className = "sheet-frozen-sel";
		d.style.left = (this.colX[inCol] - 1) + "px";
		d.style.top = (this.headTop - this.frozenHeight - 1) + "px";
		d.style.width = (this.colX[inCol1 + 1] - this.colX[inCol] - 1) + "px";
		d.style.height = (this.frozenHeight - 1) + "px";
		inner.appendChild(d);
	},

	//* A find match: a box drawn around the cell (kept until cleared with null).
	setHighlight: function(inRow, inCol) {
		this.highlight = inRow === null || inRow === undefined ? null : {r: inRow, c: inCol};
		this.win = null;
		this.scrolled();
	},

	highlightHtml: function() {
		var h = this.highlight;
		if (h.r < this.bodyFirst()) {
			return "";
		}
		var top = this.rowTop(h.r), left = this.colX[h.c];
		return "<div class=\"sheet-hit\" style=\"left:" + left + "px;top:" + top + "px;width:" + (this.colX[h.c + 1] - left - 1) +
			"px;height:" + (this.rowBottom(h.r) - top - 1) + "px\"></div>";
	},

	setScroll: function(inLeft, inTop) {
		var v = this.viewport();
		var maxL = Math.max(0, this.colX[this.nCols] - v.width), maxT = Math.max(0, this.rowY[this.rowY.length - 1] - v.height);
		this.$.scroller.setScrollPositionDirect(Math.max(0, Math.min(maxL, inLeft)), Math.max(0, Math.min(maxT, inTop)));
		this.scrolled();
	},

	//* Scrolls the least needed to show the cell.
	scrollToCell: function(inRow, inCol) {
		if (!this.colX) {
			return;
		}
		var v = this.viewport();
		var left = v.left, top = v.top;
		var x0 = this.colX[inCol], x1 = this.colX[Math.min(inCol + 1, this.nCols)];
		if (x0 < left) {
			left = x0;
		} else if (x1 > left + v.width) {
			left = x1 - v.width;
		}
		if (inRow >= this.bodyFirst()) {
			var y0 = this.rowTop(inRow), y1 = this.rowBottom(inRow);
			if (y0 < top) {
				top = y0;
			} else if (y1 > top + v.height) {
				top = y1 - v.height + 4;
			}
		}
		this.setScroll(left, top);
	},

	/**
		More rows (or the final widths and merges) of a streamed sheet:
		measure again, keeping the scroll position; inFull also re-reads the
		merges.
	*/
	grow: function(inFull) {
		if (!this.sheet || !this.hasNode()) {
			return;
		}
		if (inFull) {
			this.prepareMerges();
		}
		var s = this.$.scroller;
		var left = s.scrollLeft, top = s.scrollTop;
		var w = this.win, drawnRows = this.drawnRows || 0, drawnCols = this.nCols;
		this.measure();
		// New rows below the drawn window change only the size of the sheet: keep what is drawn.
		if (!inFull && w && w.r1 < drawnRows && this.nCols === drawnCols) {
			this.win = w;
		}
		this.setScroll(left, top);
		this.placeSelection();
	},

	//* After a cell's text changed: redraw it.
	refresh: function() {
		this.win = null;
		this.scrolled();
	},

	resizeHandler: function() {
		this.viewSize = null;
		this.inherited(arguments);
		if (this.sheet) {
			this.win = null;
			this.scrolled();
		}
	}
});
