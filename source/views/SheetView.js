/**
	A workbook: the cell bar (where the selected cell's value or formula is
	shown and edited), the grid, and a strip of sheet tabs.

	Editing is light: a value, a number or a formula per cell. Edits are
	kept as a list; "Save a copy" (in the viewer) writes them into a new
	file and never over the original.
*/
enyo.kind({
	name: "SheetView",
	kind: enyo.VFlexBox,
	className: "sheet-view",
	events: {
		onDirty: "",
		onMessage: ""
	},
	published: {
		editable: true
	},
	components: [
		{name: "cellBar", kind: enyo.HFlexBox, className: "sheet-cellbar", align: "center", components: [
			{name: "ref", className: "sheet-ref", content: ""},
			{name: "value", flex: 1, className: "sheet-value", content: ""},
			{name: "editor", kind: "Input", flex: 1, className: "sheet-editor", showing: false,
				autoCapitalize: "lowercase", autocorrect: false, spellcheck: false, autoWordComplete: false, onkeypress: "editorKey"},
			{name: "editButton", kind: "Button", className: "enyo-button-dark sheet-edit-button", caption: $L("Edit"), onclick: "startEdit"},
			{name: "doneButton", kind: "Button", className: "enyo-button-affirmative sheet-edit-button", caption: $L("Done"),
				showing: false, onclick: "commitEdit"}
		]},
		{name: "grid", kind: "SheetGrid", flex: 1, onSelect: "cellSelected", onActivate: "startEdit"},
		{name: "notice", className: "sheet-notice", showing: false},
		{name: "tabs", kind: "Scroller", className: "sheet-tabs", vertical: false, autoVertical: false, autoHorizontal: true,
			components: [{name: "tabStrip", className: "sheet-tabstrip"}]}
	],

	create: function() {
		this.inherited(arguments);
		this.model = null;
		this.sheetIndex = 0;
		this.edits = [];
		this.editing = false;
	},

	//* inModel: DocSheets model; inEditable: false for formats that cannot be saved.
	setModel: function(inModel) {
		this.model = inModel;
		this.edits = [];
		this.editing = false;
		this.buildTabs();
		var first = 0;
		for (var i = 0; i < inModel.sheets.length; i++) {
			if (!inModel.sheets[i].hidden) {
				first = i;
				break;
			}
		}
		this.showSheet(first);
	},

	editableChanged: function() {
		this.$.editButton.setShowing(this.editable && !this.editing);
		this.$.cellBar.addRemoveClass("sheet-cellbar-readonly", !this.editable);
	},

	buildTabs: function() {
		this.$.tabStrip.destroyControls();
		var sheets = this.model.sheets;
		for (var i = 0; i < sheets.length; i++) {
			this.$.tabStrip.createComponent({kind: enyo.Control, className: "sheet-tab" + (sheets[i].hidden ? " sheet-tab-hidden" : ""),
				content: enyo.string.escapeHtml(sheets[i].name || ("Sheet " + (i + 1))), sheetIndex: i, onclick: "tabClick", owner: this});
		}
		this.$.tabStrip.render();
		var show = sheets.length > 1;
		if (show !== this.$.tabs.showing) {
			this.$.tabs.setShowing(show);
			this.resized();
		}
	},

	tabClick: function(inSender) {
		this.commitEdit();
		this.showSheet(inSender.sheetIndex);
	},

	showSheet: function(inIndex) {
		this.sheetIndex = inIndex;
		var sheet = this.model.sheets[inIndex];
		var tabs = this.$.tabStrip.getControls();
		for (var i = 0; i < tabs.length; i++) {
			tabs[i].addRemoveClass("sheet-tab-current", i === inIndex);
		}
		this.$.grid.setSheet(sheet);
		this.showNotice(sheet);
		this.$.ref.setContent("");
		this.$.value.setContent("");
		this.selected = null;
		this.$.editButton.setShowing(this.editable);
		this.$.editButton.setDisabled(!this.selected);
	},

	currentSheet: function() {
		return this.model ? this.model.sheets[this.sheetIndex] : null;
	},

	cellSelected: function(inSender, inRow, inCol) {
		if (this.editing) {
			this.commitEdit();
		}
		this.selected = {r: inRow, c: inCol};
		var sheet = this.currentSheet();
		this.$.ref.setContent(DocSheets.cellRef(inRow, inCol));
		this.$.value.setContent(enyo.string.escapeHtml(this.cellSource(sheet, inRow, inCol)));
		this.$.editButton.setDisabled(false);
	},

	//* What the cell bar shows: the formula when there is one, else the text.
	cellSource: function(inSheet, inRow, inCol) {
		var f = inSheet.formulas[inRow + "," + inCol];
		if (f) {
			return f;
		}
		var row = inSheet.text[inRow];
		var t = row ? row[inCol] : "";
		return t === undefined || t === null ? "" : String(t);
	},

	startEdit: function() {
		if (!this.editable || !this.selected) {
			return;
		}
		var sheet = this.currentSheet();
		this.editing = true;
		this.$.value.hide();
		this.$.editButton.hide();
		this.$.editor.show();
		this.$.doneButton.show();
		this.$.editor.setValue(this.cellSource(sheet, this.selected.r, this.selected.c));
		this.$.editor.forceFocus();
	},

	editorKey: function(inSender, inEvent) {
		if (inEvent.keyCode === 13) {
			this.commitEdit();
			// Enter moves down, as in other spreadsheets.
			if (this.selected) {
				var r = this.selected.r + 1;
				this.$.grid.select(r, this.selected.c);
				this.$.grid.scrollToCell(r, this.selected.c);
			}
			return true;
		}
		if (inEvent.keyCode === 27) {
			this.cancelEdit();
			return true;
		}
	},

	cancelEdit: function() {
		if (!this.editing) {
			return false;
		}
		this.editing = false;
		this.$.editor.hide();
		this.$.doneButton.hide();
		this.$.value.show();
		this.$.editButton.setShowing(this.editable);
		return true;
	},

	commitEdit: function() {
		if (!this.editing) {
			return;
		}
		var text = this.$.editor.getValue();
		var sel = this.selected;
		var sheet = this.currentSheet();
		this.cancelEdit();
		if (!sel || text === this.cellSource(sheet, sel.r, sel.c)) {
			return;
		}
		var p = DocSheets.parseInput(text);
		if (!sheet.text[sel.r]) {
			sheet.text[sel.r] = [];
			sheet.kinds[sel.r] = [];
		}
		// Shown as typed: there is no formula engine, so a formula shows its formula until saved and reopened.
		sheet.text[sel.r][sel.c] = p.t === "z" ? "" : text;
		sheet.kinds[sel.r][sel.c] = p.t === "n" ? "n" : "s";
		if (p.f) {
			sheet.formulas[sel.r + "," + sel.c] = text;
		} else {
			delete sheet.formulas[sel.r + "," + sel.c];
		}
		if (sel.r >= sheet.rows) {
			sheet.rows = sel.r + 1;
		}
		if (sel.c >= sheet.cols) {
			sheet.cols = sel.c + 1;
		}
		this.edits.push({sheet: sheet.index, r: sel.r, c: sel.c, text: text});
		this.$.grid.refresh();
		this.$.value.setContent(enyo.string.escapeHtml(text));
		this.doDirty(this.edits.length);
	},

	// --- Streaming -------------------------------------------------------

	//* A sheet found while streaming (ods names its sheets as it goes).
	addSheet: function(inSheet) {
		this.model.sheets.push(inSheet);
		this.buildTabs();
		var tabs = this.$.tabStrip.getControls();
		for (var i = 0; i < tabs.length; i++) {
			tabs[i].addRemoveClass("sheet-tab-current", i === this.sheetIndex);
		}
	},

	appendRows: function(inSheetIndex, inRows) {
		var sheet = this.model.sheets[inSheetIndex];
		if (!sheet) {
			return;
		}
		for (var i = 0; i < inRows.length; i++) {
			var r = inRows[i][0];
			sheet.text[r] = inRows[i][1];
			sheet.kinds[r] = inRows[i][2];
			if (r + 1 > sheet.rows) {
				sheet.rows = r + 1;
			}
			if (inRows[i][1].length > sheet.cols) {
				sheet.cols = inRows[i][1].length;
			}
		}
		this.scheduleGrow(inSheetIndex);
	},

	sheetDone: function(inSheetIndex, inInfo) {
		var sheet = this.model.sheets[inSheetIndex];
		if (!sheet) {
			return;
		}
		DocEngine.applyInfo(sheet, inInfo || {});
		if (inSheetIndex === this.sheetIndex) {
			enyo.job.stop("documents-grow");
			this.$.grid.grow(true);
			this.showNotice(sheet);
		}
	},

	allDone: function(inError) {
		var sheet = this.currentSheet();
		if (sheet) {
			this.showNotice(sheet, inError);
		}
	},

	//* The grid takes new rows at most a few times a second.
	scheduleGrow: function(inSheetIndex) {
		if (inSheetIndex !== this.sheetIndex) {
			return;
		}
		var grid = this.$.grid;
		enyo.job("documents-grow", function() {
			grid.grow(false);
		}, 1000);
	},

	showNotice: function(inSheet, inError) {
		var notes = [];
		if (inError) {
			notes.push($L("Not all of this workbook could be read: ") + inError);
		}
		if (inSheet.truncated) {
			notes.push($L("This workbook is larger than Documents can hold on this device: the first #{rows} rows are shown.")
				.replace("#{rows}", String(inSheet.rows)));
		}
		if (!this.editable && this.model && this.model.streaming) {
			notes.push($L("Large workbook: shown read-only."));
		}
		var text = enyo.string.escapeHtml(notes.join(" "));
		if (text !== this.noticeText) {
			this.noticeText = text;
			this.$.notice.setContent(text);
		}
		// Every resize makes the page lay out again: only when the notice comes or goes.
		if ((notes.length > 0) !== this.$.notice.showing) {
			this.$.notice.setShowing(notes.length > 0);
			this.resized();
		}
	},

	// --- Find ------------------------------------------------------------

	//* Every cell of the current sheet whose text matches (case and accents ignored).
	search: function(inQuery) {
		var q = DocFinder.foldString(String(inQuery || "").replace(/^\s+|\s+$/g, ""));
		this.matches = [];
		var sheet = this.currentSheet();
		if (!q || !sheet) {
			this.$.grid.setHighlight(null);
			return 0;
		}
		for (var r = 0; r < sheet.text.length && this.matches.length < 10000; r++) {
			var row = sheet.text[r];
			if (!row) {
				continue;
			}
			for (var c = 0; c < row.length; c++) {
				var t = row[c];
				if (t && DocFinder.foldString(String(t)).indexOf(q) >= 0) {
					this.matches.push({r: r, c: c});
				}
			}
		}
		return this.matches.length;
	},

	showMatch: function(inIndex) {
		var m = this.matches && this.matches[inIndex];
		if (!m) {
			return;
		}
		this.$.grid.setHighlight(m.r, m.c);
		this.$.grid.scrollToCell(m.r, m.c);
		this.$.grid.select(m.r, m.c);
	},

	clearSearch: function() {
		this.matches = [];
		this.$.grid.setHighlight(null);
	},

	setZoom: function(inZoom) {
		if (this.$.grid.zoom !== inZoom) {
			this.$.grid.setZoom(inZoom);
		}
	},

	setFreezeHeader: function(inFreeze) {
		if (this.$.grid.freezeHeader !== !!inFreeze) {
			this.$.grid.setFreezeHeader(!!inFreeze);
		}
	}
});
