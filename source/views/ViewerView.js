/**
	Shows one document: text documents as pages (or reflowed to the screen
	on a phone), presentations as slides one under the other, workbooks in
	a grid. Zoom, find, night reading and document info for all of them.
	Editing, with "Save a copy": cells of workbooks; paragraphs of
	OpenDocument text (tap one in edit mode); all of a plain text or
	Markdown file; Word documents in the .docx editor (DocxEditor.js).

		viewer.openPath("/media/internal/Report.odt");
		viewer.timings   // {read, parse, render, total} in ms of the last open, for the device tests
*/
enyo.kind({
	name: "ViewerView",
	kind: enyo.VFlexBox,
	className: "viewer",
	events: {
		onOpened: "",
		onClose: "",
		onOpenLink: "",
		onHandOff: "",
		onSaveCopy: "",
		onNight: ""
	},
	zoomSteps: [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4],
	components: [
		{name: "header", kind: enyo.HFlexBox, className: "viewer-header", align: "center", components: [
			{name: "titleBox", flex: 1, className: "viewer-titlebox", components: [
				{name: "title", className: "viewer-title", content: ""},
				{name: "subtitle", className: "viewer-subtitle", content: ""}
			]},
			{name: "saveButton", kind: "Button", className: "enyo-button-affirmative viewer-save", caption: $L("Save a Copy"),
				showing: false, onclick: "saveCopy"}
		]},
		{name: "findBar", kind: enyo.HFlexBox, className: "viewer-findbar", align: "center", showing: false, components: [
			{name: "findInput", kind: "Input", flex: 1, hint: $L("Find in document"), changeOnInput: true, autoCapitalize: "lowercase",
				autocorrect: false, spellcheck: false, autoWordComplete: false, onchange: "findChanged", onkeypress: "findKey"},
			{name: "findCount", className: "viewer-findcount", content: ""},
			{kind: "ToolButton", name: "findPrev", className: "viewer-tool", icon: "images/prev.svg", onclick: "findPrev"},
			{kind: "ToolButton", name: "findNext", className: "viewer-tool", icon: "images/next.svg", onclick: "findNext"},
			{kind: "ToolButton", name: "findClose", className: "viewer-tool", icon: "images/close.svg", onclick: "closeFind"}
		]},
		{name: "body", flex: 1, className: "viewer-body", components: [
			{name: "scroller", kind: "Scroller", className: "viewer-scroller", autoHorizontal: true, horizontal: true,
				components: [{name: "content", className: "viewer-content", onclick: "contentClick"}]},
			{name: "sheet", kind: "SheetView", className: "viewer-sheet", showing: false, onDirty: "sheetDirty"},
			{name: "docx", kind: "DocxEditor", className: "viewer-docx", showing: false, onChanged: "markEdited",
				onZoomChanged: "updateZoomLabel"},
			{name: "lok", kind: "LokEditor", className: "viewer-docx", showing: false, onChanged: "lokChanged",
				onClosed: "lokClosed"},
			{name: "status", className: "viewer-status", showing: false, components: [
				{name: "spinner", kind: "SpinnerLarge"},
				{name: "statusText", className: "viewer-status-text", content: ""}
			]}
		]},
		{name: "toolbar", kind: "Toolbar", className: "viewer-toolbar", components: [
			{kind: "ToolButton", name: "zoomOut", className: "viewer-tool", icon: "images/zoom-out.svg", onclick: "zoomOut"},
			{kind: "ToolButton", name: "zoomLabel", className: "viewer-tool viewer-zoomlabel", caption: "100%", onclick: "zoomFit"},
			{kind: "ToolButton", name: "zoomIn", className: "viewer-tool", icon: "images/zoom-in.svg", onclick: "zoomIn"},
			{kind: "ToolButton", name: "layoutButton", className: "viewer-tool", icon: "images/reflow.svg", onclick: "toggleLayout"},
			{kind: "ToolButton", name: "findButton", className: "viewer-tool", icon: "images/find.svg", onclick: "openFind"},
			{kind: "ToolButton", name: "nightButton", className: "viewer-tool", icon: "images/night.svg", onclick: "toggleNight"},
			{kind: "ToolButton", name: "infoButton", className: "viewer-tool", icon: "images/info.svg", onclick: "showInfo"},
			{kind: "ToolButton", name: "editButton", className: "viewer-tool", icon: "images/edit.svg", onclick: "toggleEdit", showing: false}
		]},
		// Laid out as the Wi-Fi settings' network info: the name, a line under it, the details
		// in groups that scroll within the screen, and Done below them.
		{name: "infoDialog", kind: "ModalDialog", className: "enyo-popup enyo-modaldialog viewer-info-dialog", components: [
			{name: "infoSubtitle", className: "viewer-info-subtitle", content: ""},
			{name: "infoScroller", kind: "Scroller", className: "viewer-info-scroller", horizontal: false, autoHorizontal: false,
				components: [
					{name: "infoGroups", className: "viewer-info"}
				]},
			{kind: "Button", caption: $L("Done"), onclick: "closeInfo"}
		]},
		{name: "linkDialog", kind: "ModalDialog", caption: $L("Open Link?"), components: [
			{name: "linkText", className: "viewer-linktext", content: ""},
			{kind: enyo.HFlexBox, components: [
				{kind: "Button", flex: 1, caption: $L("Cancel"), onclick: "closeLink"},
				{kind: "Button", flex: 1, className: "enyo-button-affirmative", caption: $L("Open"), onclick: "openLink"}
			]}
		]}
	],

	create: function() {
		this.inherited(arguments);
		this.zoom = 1;
		this.fitMode = true;
		this.reflow = null;
		this.night = false;
		this.doc = null;
		this.objectUrls = [];
		this.timings = {};
		this.finder = null;
		this.findIndex = -1;
		this.findTotal = 0;
		this.loadToken = 0;
		this.editMode = false;
		this.canEdit = false;
		this.paraBlocks = {};
		this.paraEditor = null;
	},

	rendered: function() {
		this.inherited(arguments);
		var self = this;
		var body = this.$.body.hasNode();
		if (body && !this.wheelHooked) {
			this.wheelHooked = true;
			// Ctrl + wheel (and the touchpad pinch, which browsers send as one) zooms.
			body.addEventListener("wheel", function(e) {
				if (e.ctrlKey && self.doc) {
					e.preventDefault();
					self.stepZoom(e.deltaY < 0 ? 1 : -1);
				}
			}, {passive: false});
			this.hookPinch(body);
		}
	},

	// --- Opening ---------------------------------------------------------

	openPath: function(inPath) {
		var token = ++this.loadToken;
		var self = this;
		var t0 = this.openStarted = Date.now();
		this.reset();
		this.timings = {};
		this.path = inPath;
		this.lastError = null;
		// Not "name": that is the component's own name, which the Pane goes by.
		this.fileName = DocLibrary.baseName(inPath);
		this.convertedFrom = null;
		this.$.title.setContent(enyo.string.escapeHtml(this.fileName));
		this.$.subtitle.setContent(enyo.string.escapeHtml(DocLibrary.dirName(inPath).replace(/^\/media\/internal/, $L("USB storage"))));
		this.showStatus($L("Opening\u2026"), true);
		DocLoader.read(inPath, function(inErr, inBuffer) {
			if (token !== self.loadToken) {
				return;
			}
			if (inErr) {
				self.showError(inErr.message);
				return;
			}
			self.size = inBuffer.byteLength;
			self.timings = {read: Date.now() - t0};
			self.openBuffer(self.fileName, inBuffer, token, t0);
		});
	},

	openBuffer: function(inName, inBuffer, inToken, inT0) {
		var self = this;
		var t1 = Date.now();
		// docx-preview needs the bytes in the page; the engine gets its own copy.
		var keep = FileTypes.kindFromName(inName) === "docx" || FileTypes.sniff(new Uint8Array(inBuffer, 0, Math.min(8, inBuffer.byteLength))) === "zip" ?
			inBuffer : null;
		var forEngine = keep ? inBuffer.slice(0) : inBuffer;
		// Content another app shows, under a name that said otherwise: hand it on.
		var handOff = FileTypes.handOffMime(FileTypes.resolve(inName, new Uint8Array(inBuffer, 0, Math.min(8, inBuffer.byteLength))).kind);
		if (handOff && !FileTypes.handOffMime(FileTypes.kindFromName(inName))) {
			this.pendingHandOff = {name: inName, buffer: forEngine, token: inToken};
			this.doHandOff(this.path, handOff);
			return;
		}
		var onProgress = function(inUpdate) {
			if (inToken === self.loadToken && self.doc && self.doc.family === "sheet") {
				self.sheetProgress(inUpdate);
			}
		};
		DocLoader.open(inName, forEngine, function(inErr, inResult) {
			if (inToken !== self.loadToken) {
				return;
			}
			self.timings.parse = Date.now() - t1;
			if (inErr) {
				var to = !self.convertedFrom && inErr.code === "unsupported" && FileTypes.convertsTo(inErr.kind);
				if (to) {
					self.convertAndOpen(inErr, to, inToken, inT0);
					return;
				}
				self.showError(inErr.message, inErr.code);
				return;
			}
			self.doc = inResult;
			var t2 = Date.now();
			var done = function(inRenderErr) {
				if (inToken !== self.loadToken) {
					return;
				}
				if (inRenderErr) {
					self.showError(inRenderErr.message || String(inRenderErr));
					return;
				}
				self.timings.render = Date.now() - t2;
				self.timings.total = Date.now() - inT0;
				self.hideStatus();
				self.updateTools();
				self.doOpened(self.path, inResult.kind);
			};
			try {
				self.renderResult(inResult, keep, done);
			} catch (e) {
				enyo.warn("render failed: " + (e && e.stack || e));
				done(e);
			}
		}, onProgress);
	},

	/**
		A document the app does not read: LibreOffice converts it to
		OpenDocument, and that is shown under the document's own name; a copy
		is saved next to it. Without the service, inErr (why it was refused)
		is shown.
	*/
	convertAndOpen: function(inErr, inFormat, inToken, inT0) {
		var self = this;
		var kind = inErr.kind;
		this.showStatus($L("Converting with LibreOffice\u2026"), true);
		FileService.convert(this.path, inFormat, function(inErr2, inResult) {
			if (inToken !== self.loadToken) {
				return;
			}
			if (inErr2 || !inResult || !inResult.target) {
				// No service (LibreOffice is not installed): why the app itself refused the document.
				var missing = !inErr2 || !inErr2.code || /does not exist|not found|unknown service/i.test(inErr2.message);
				self.showError(missing ? inErr.message : $L("LibreOffice could not convert this document: ") + inErr2.message, inErr.code);
				return;
			}
			DocLoader.read(inResult.target, function(inErr3, inBuffer) {
				if (inToken !== self.loadToken) {
					return;
				}
				if (inErr3) {
					self.showError(inErr3.message);
					return;
				}
				// The size shown stays the document's own.
				self.convertedFrom = kind;
				self.openBuffer(DocLibrary.baseName(inResult.target), inBuffer, inToken, inT0);
			});
		});
	},

	//* No app took the file onHandOff offered: say so, the way any refused file is.
	handOffFailed: function() {
		var p = this.pendingHandOff;
		this.pendingHandOff = null;
		if (!p || p.token !== this.loadToken) {
			return;
		}
		var self = this;
		DocLoader.open(p.name, p.buffer, function(inErr) {
			if (p.token === self.loadToken) {
				self.showError(inErr ? inErr.message : $L("This document could not be opened."), inErr && inErr.code);
			}
		});
	},

	//* More rows of a streamed workbook.
	sheetProgress: function(inUpdate) {
		var sheetView = this.$.sheet;
		this.streaming = !inUpdate.finished;
		if (inUpdate.newSheet) {
			sheetView.addSheet(inUpdate.newSheet);
		}
		if (inUpdate.rows) {
			sheetView.appendRows(inUpdate.sheet, inUpdate.rows);
		}
		if (inUpdate.done) {
			sheetView.sheetDone(inUpdate.sheet, inUpdate.info);
		}
		if (inUpdate.finished) {
			this.timings.complete = Date.now() - this.openStarted;
			this.streaming = false;
			sheetView.allDone(inUpdate.error);
			this.$.subtitle.setContent(enyo.string.escapeHtml(this.describe()));
			if (this.$.findBar.showing && this.$.findInput.getValue()) {
				this.runFind();
			}
		} else if (!this.readingShown) {
			this.readingShown = true;
			this.$.subtitle.setContent(enyo.string.escapeHtml(this.describe() + " \u00b7 " + $L("reading\u2026")));
		}
	},

	reset: function() {
		this.pendingHandOff = null;
		this.closeFind();
		if (this.progressive) {
			this.progressive.cancel();
			this.progressive = null;
		}
		this.renderingAll = false;
		this.finder = null;
		this.doc = null;
		this.edited = false;
		this.$.saveButton.hide();
		this.dropEditorHost(this.paraEditor && this.paraEditor.host);
		this.dropEditorHost(this.textHost);
		this.textHost = null;
		this.paraEditor = null;
		this.paraBlocks = {};
		this.editMode = false;
		this.canEdit = false;
		this.textInput = null;
		this.docBytes = null;
		this.$.docx.close();
		this.$.docx.hide();
		this.$.lok.close();
		this.$.lok.hide();
		this.lokEdited = false;
		this.$.editButton.hide();
		this.$.editButton.addRemoveClass("viewer-tool-on", false);
		this.revokeUrls();
		var c = this.$.content.hasNode();
		if (c) {
			c.innerHTML = "";
			c.className = "viewer-content";
			c.style.width = "";
			c.style.zoom = "";
		}
		this.$.sheet.hide();
		this.$.scroller.show();
		this.$.scroller.setScrollPositionDirect(0, 0);
		this.slideBoxes = null;
		this.zoom = 1;
		this.fitMode = true;
	},

	revokeUrls: function() {
		for (var i = 0; i < this.objectUrls.length; i++) {
			URL.revokeObjectURL(this.objectUrls[i]);
		}
		this.objectUrls = [];
	},

	imageUrls: function(inImages) {
		var out = {};
		for (var k in inImages) {
			if (inImages.hasOwnProperty(k) && inImages[k].bytes) {
				var url = URL.createObjectURL(new Blob([inImages[k].bytes], {type: inImages[k].mime}));
				this.objectUrls.push(url);
				out[k] = url;
			}
		}
		return out;
	},

	isPhone: function() {
		var n = this.hasNode();
		return n ? n.clientWidth < 700 : false;
	},

	renderResult: function(inResult, inBuffer, inDone) {
		var content = this.$.content.hasNode();
		var family = inResult.family;
		this.kind = inResult.kind;
		this.$.content.addRemoveClass("night-invert", false);
		if (family === "sheet") {
			this.$.scroller.hide();
			this.$.sheet.show();
			// A streamed (large) workbook is read-only: the full reader would not fit in memory.
			// Editing needs this app's service to save the copy; it is optional in a developer install.
			this.streaming = !!inResult.streamed;
			this.$.sheet.setEditable(false);
			this.readingShown = false;
			this.$.sheet.setFreezeHeader(this.freezeHeader !== false);
			this.$.sheet.setModel(inResult.model);
			if (!inResult.streamed) {
				var sheetView = this.$.sheet, token = this.loadToken, self = this;
				FileService.serviceAvailable(function(inOk) {
					if (token === self.loadToken) {
						sheetView.setEditable(inOk);
					}
				});
			}
			inDone();
			return;
		}
		if (this.reflow === null) {
			// Phones read reflowed text; tablets see the page.
			this.reflow = this.isPhone();
		}
		if (inResult.kind === "docx") {
			// The file, for the .docx editor.
			this.docBytes = inBuffer;
			this.offerEditing();
		}
		if (inResult.kind === "docx" && !inResult.model) {
			this.renderDocx(content, inBuffer, inDone);
			return;
		}
		var ctx = this.renderCtx = {images: this.imageUrls(inResult.images || {})};
		if (inResult.kind === "txt" || inResult.kind === "md") {
			this.renderPlain(inResult.kind, inResult.model);
			this.offerEditing();
		} else if (family === "slides") {
			content.className = "viewer-content doc-slides";
			this.slideBoxes = DocRender.slides(content, inResult.model.slides, ctx);
			if (inResult.editable) {
				this.indexSlides(inResult.model.slides);
				this.offerEditing();
			}
		} else {
			// ODF text, or a large Word document read by Docx.js: pages of blocks, a slice at a time.
			content.className = "viewer-content doc-odf" + (inResult.simplified ? " doc-simplified" : "");
			var page = DocRender.el("div", "doc-page", content);
			var p = inResult.model.page;
			this.pageWidth = p.width;
			page.style.width = p.width;
			page.style.padding = [p.marginTop, p.marginRight, p.marginBottom, p.marginLeft].join(" ");
			DocRender.apply(page, inResult.model.base);
			var viewer = this;
			if (inResult.editable) {
				this.indexParagraphs(inResult.model.blocks);
				this.offerEditing();
			}
			this.renderingAll = true;
			this.progressive = DocRender.progressive(page, inResult.model.blocks, ctx, function() {
				viewer.renderingAll = false;
				viewer.timings.complete = Date.now() - viewer.openStarted;
				viewer.$.scroller.calcBoundaries();
				// A search typed while the rest was still coming in: search the whole document now.
				if (viewer.$.findBar.showing && viewer.$.findInput.getValue()) {
					viewer.runFind();
				}
			});
		}
		this.applyLayout();
		inDone();
	},

	//* A plain text or Markdown file: its text, or (Markdown) the page made from it.
	renderPlain: function(inKind, inModel) {
		var content = this.$.content.hasNode();
		content.innerHTML = "";
		if (inKind === "txt") {
			content.className = "viewer-content doc-plain";
			var pre = DocRender.el("div", "doc-plaintext", content);
			pre.appendChild(document.createTextNode(inModel.text));
		} else {
			content.className = "viewer-content doc-markdown";
			var md = DocRender.el("div", "doc-markdown-body", content);
			// DocText.markdown escapes all of the input; its own tags are the only markup.
			md.innerHTML = inModel.html;
			DocRender.neutraliseLinks(md);
		}
	},

	renderDocx: function(inContent, inBuffer, inDone) {
		var self = this;
		DocLoader.loadScripts(["lib/jszip/jszip.min.js", "lib/docx-preview/docx-preview.min.js"], function(inErr) {
			if (inErr) {
				inDone(inErr);
				return;
			}
			inContent.className = "viewer-content doc-docx";
			var styleHost = DocRender.el("div", "doc-docx-styles", inContent);
			var bodyHost = DocRender.el("div", "doc-docx-body", inContent);
			window.docx.renderAsync(inBuffer, bodyHost, styleHost, {
				className: "docx",
				inWrapper: true,
				ignoreWidth: false,
				ignoreHeight: false,
				ignoreFonts: false,
				breakPages: true,
				ignoreLastRenderedPageBreak: true,
				experimental: false,
				trimXmlDeclaration: true,
				useBase64URL: false,
				renderHeaders: true,
				renderFooters: true,
				renderFootnotes: true,
				renderEndnotes: true,
				renderComments: false,
				debug: false
			}).then(function() {
				DocRender.neutraliseLinks(bodyHost);
				// Blob URLs docx-preview made for pictures: collect them to revoke on close.
				var imgs = bodyHost.querySelectorAll("img[src^='blob:']");
				for (var i = 0; i < imgs.length; i++) {
					self.objectUrls.push(imgs[i].src);
				}
				self.applyLayout();
				inDone();
			}, function(e) {
				inDone(new Error($L("This Word document could not be shown: ") + (e && e.message || e)));
			});
		});
	},

	showStatus: function(inText, inBusy) {
		this.$.status.show();
		this.$.statusText.setContent(enyo.string.escapeHtml(inText));
		this.$.spinner.setShowing(!!inBusy);
		if (inBusy) {
			this.$.spinner.show();
		}
		this.$.status.addRemoveClass("viewer-status-error", !inBusy);
	},

	hideStatus: function() {
		this.$.spinner.hide();
		this.$.status.hide();
	},

	showError: function(inMessage, inCode) {
		this.doc = null;
		this.updateTools();
		this.$.scroller.show();
		this.$.sheet.hide();
		this.showStatus(inMessage || $L("This document could not be opened."), false);
		this.lastError = {message: inMessage, code: inCode || ""};
	},

	updateTools: function() {
		var has = !!this.doc;
		// The .docx editor lays out and searches its own pages; the zoom keys work its zoom.
		var own = has && !this.$.docx.isOpen() && !this.$.lok.isOpen();
		var reflowable = own && (this.doc.kind === "docx" || this.doc.kind === "odt");
		this.$.zoomOut.setDisabled(!has);
		this.$.zoomIn.setDisabled(!has);
		this.$.zoomLabel.setDisabled(!has);
		this.$.findButton.setDisabled(!own);
		this.$.nightButton.setDisabled(!own);
		this.$.infoButton.setDisabled(!has);
		this.$.layoutButton.setShowing(reflowable);
		this.$.layoutButton.setIcon(this.reflow ? "images/page.svg" : "images/reflow.svg");
		this.updateZoomLabel();
		if (has) {
			this.$.subtitle.setContent(enyo.string.escapeHtml(this.describe()));
		}
	},

	describe: function() {
		var d = this.doc;
		var label = FileTypes.kinds[d.kind] ? FileTypes.kinds[d.kind].label : d.kind;
		if (this.convertedFrom && FileTypes.kinds[this.convertedFrom]) {
			label = FileTypes.kinds[this.convertedFrom].label + ", " + $L("converted by LibreOffice");
		}
		var parts = [label];
		if (d.family === "slides") {
			var n = d.model.slides.length;
			parts.push(n === 1 ? $L("1 slide") : $L("#{n} slides").replace("#{n}", n));
		} else if (d.family === "sheet") {
			var s = d.model.sheets.length;
			parts.push(s === 1 ? $L("1 sheet") : $L("#{n} sheets").replace("#{n}", s));
		}
		parts.push(DocLibrary.formatSize(this.size));
		return parts.join(" \u00b7 ");
	},

	// --- Zoom and layout -------------------------------------------------

	availableWidth: function() {
		var n = this.$.body.hasNode();
		return n ? n.clientWidth : 480;
	},

	//* Lays the content out for the current zoom, fit mode and reflow.
	applyLayout: function() {
		if (!this.doc) {
			return;
		}
		if (this.doc.family === "sheet") {
			this.$.sheet.setZoom(this.fitMode ? 1 : this.zoom);
			if (this.fitMode) {
				this.zoom = 1;
			}
			this.updateZoomLabel();
			return;
		}
		var c = this.$.content.hasNode();
		if (!c) {
			return;
		}
		var avail = this.availableWidth() - 2;
		var kind = this.doc.kind;
		if (this.slideBoxes) {
			// Every slide fitted to the width (fit) or at the zoom.
			var maxW = 0;
			for (var i = 0; i < this.slideBoxes.length; i++) {
				maxW = Math.max(maxW, this.slideBoxes[i].width);
			}
			if (this.fitMode) {
				this.zoom = Math.max(0.1, (avail - 16) / maxW);
			}
			for (var k = 0; k < this.slideBoxes.length; k++) {
				var b = this.slideBoxes[k];
				b.frame.style.width = Math.round(b.width * this.zoom) + "px";
				b.frame.style.height = Math.round(b.height * this.zoom) + "px";
				b.slide.style.transform = "scale(" + this.zoom + ")";
			}
			c.style.width = this.fitMode ? "" : Math.round(maxW * this.zoom + 16) + "px";
		} else if (kind === "txt" || kind === "md" || this.reflow) {
			// Reflowed: the text wraps at the screen width at any zoom.
			if (this.fitMode) {
				this.zoom = 1;
			}
			c.classList.add("doc-reflow");
			c.style.zoom = this.zoom;
			c.style.width = Math.floor(avail / this.zoom) + "px";
		} else {
			c.classList.remove("doc-reflow");
			c.style.zoom = 1;
			c.style.width = "";
			var natural = c.scrollWidth || avail;
			if (this.fitMode) {
				this.zoom = Math.min(2, Math.max(0.2, avail / natural));
			}
			c.style.zoom = this.zoom;
		}
		this.updateZoomLabel();
		this.$.scroller.calcBoundaries();
		this.$.scroller.stabilize();
	},

	updateZoomLabel: function() {
		var docx = this.$.lok.isOpen() ? this.$.lok.zoomPercent() : this.$.docx.isOpen() ? this.$.docx.zoomPercent() : 0;
		this.$.zoomLabel.setCaption((docx || Math.round(this.zoom * 100)) + "%");
	},

	stepZoom: function(inDirection) {
		var steps = this.zoomSteps;
		if (this.$.lok.isOpen()) {
			this.$.lok.zoomBy(inDirection > 0 ? 1.25 : 0.8);
			this.updateZoomLabel();
			return;
		}
		if (this.$.docx.isOpen()) {
			// The .docx editor's own zoom, by the same steps.
			this.$.docx.stepZoom(inDirection, steps);
			this.updateZoomLabel();
			return;
		}
		var z = this.zoom, next = z;
		if (inDirection > 0) {
			for (var i = 0; i < steps.length; i++) {
				if (steps[i] > z + 0.001) {
					next = steps[i];
					break;
				}
			}
		} else {
			for (var j = steps.length - 1; j >= 0; j--) {
				if (steps[j] < z - 0.001) {
					next = steps[j];
					break;
				}
			}
		}
		this.setZoom(next);
	},

	setZoom: function(inZoom) {
		if (!this.doc) {
			return;
		}
		var s = this.$.scroller;
		var old = this.zoom || 1;
		var v = s.hasNode();
		// Keep the middle of the screen where it was.
		var cx = (s.scrollLeft + (v ? v.clientWidth / 2 : 0)) / old, cy = (s.scrollTop + (v ? v.clientHeight / 2 : 0)) / old;
		this.zoom = Math.max(0.1, Math.min(5, inZoom));
		this.fitMode = false;
		this.applyLayout();
		if (this.doc.family !== "sheet" && v) {
			var b = s.getBoundaries();
			var left = Math.max(0, Math.min(b.right, cx * this.zoom - v.clientWidth / 2));
			var top = Math.max(0, Math.min(b.bottom, cy * this.zoom - v.clientHeight / 2));
			s.setScrollPositionDirect(left, top);
		}
	},

	zoomIn: function() {
		this.stepZoom(1);
	},

	zoomOut: function() {
		this.stepZoom(-1);
	},

	//* The zoom label: back to fit.
	zoomFit: function() {
		if (this.$.lok.isOpen()) {
			this.$.lok.zoomBy(0);
			this.updateZoomLabel();
			return;
		}
		if (this.$.docx.isOpen()) {
			this.$.docx.fitWidth();
			this.updateZoomLabel();
			return;
		}
		this.fitMode = true;
		this.applyLayout();
	},

	toggleLayout: function() {
		this.reflow = !this.reflow;
		this.fitMode = true;
		this.updateTools();
		this.applyLayout();
		this.$.scroller.setScrollPositionDirect(0, 0);
	},

	hookPinch: function(inNode) {
		var self = this;
		var start = null;
		var dist = function(e) {
			var a = e.touches[0], b = e.touches[1];
			var dx = a.clientX - b.clientX, dy = a.clientY - b.clientY;
			return Math.sqrt(dx * dx + dy * dy);
		};
		inNode.addEventListener("touchstart", function(e) {
			if (e.touches.length === 2 && self.doc) {
				start = {d: dist(e), zoom: self.zoom};
			}
		}, {passive: true});
		inNode.addEventListener("touchmove", function(e) {
			if (start && e.touches.length === 2) {
				start.scale = dist(e) / start.d;
			}
		}, {passive: true});
		inNode.addEventListener("touchend", function(e) {
			if (start && e.touches.length < 2) {
				var scale = start.scale || 1;
				var z = start.zoom;
				start = null;
				if (Math.abs(scale - 1) > 0.08) {
					self.setZoom(z * scale);
				}
			}
		}, {passive: true});
	},

	resizeHandler: function() {
		this.inherited(arguments);
		if (this.doc && this.fitMode) {
			this.applyLayout();
		}
		// A notification or the keyboard changes the room the info has
		if (this.$.infoDialog.isOpen) {
			this.fitInfo();
		}
	},

	// --- Night -----------------------------------------------------------

	setNight: function(inNight) {
		this.night = !!inNight;
		this.addRemoveClass("night", this.night);
		this.$.nightButton.addRemoveClass("viewer-tool-on", this.night);
	},

	toggleNight: function() {
		this.setNight(!this.night);
		this.doNight(this.night);
	},

	// --- Find ------------------------------------------------------------

	openFind: function() {
		if (!this.doc) {
			return;
		}
		if (this.editMode) {
			var self = this;
			this.commitParagraph(function() {
				self.leaveEdit();
				self.openFind();
			});
			return;
		}
		this.$.findBar.show();
		this.resized();
		this.$.findInput.forceFocus();
		if (this.$.findInput.getValue()) {
			this.runFind();
		}
	},

	closeFind: function() {
		if (!this.$.findBar.showing) {
			return false;
		}
		this.$.findBar.hide();
		if (this.finder) {
			this.finder.clear();
		}
		this.$.sheet.clearSearch();
		this.findTotal = 0;
		this.findIndex = -1;
		this.$.findCount.setContent("");
		this.resized();
		return true;
	},

	findChanged: function() {
		var self = this;
		// Typing fast should not search on every letter in a long document.
		enyo.job("documents-find", function() {
			self.runFind();
		}, 250);
	},

	findKey: function(inSender, inEvent) {
		if (inEvent.keyCode === 13) {
			enyo.job.stop("documents-find");
			if (this.lastQuery !== this.$.findInput.getValue()) {
				this.runFind();
			} else {
				this.findNext();
			}
			return true;
		}
	},

	runFind: function() {
		var q = this.$.findInput.getValue();
		this.lastQuery = q;
		if (!this.doc) {
			return;
		}
		if (this.doc.family === "sheet") {
			this.findTotal = this.$.sheet.search(q);
		} else {
			if (!this.finder) {
				this.finder = new DocFinder(this.$.content.hasNode());
			}
			this.findTotal = this.finder.search(q);
		}
		this.findIndex = -1;
		if (this.findTotal) {
			this.showFind(0);
		} else {
			this.$.findCount.setContent(q ? $L("No matches") : "");
		}
	},

	showFind: function(inIndex) {
		var n = this.findTotal;
		if (!n) {
			return;
		}
		this.findIndex = ((inIndex % n) + n) % n;
		this.$.findCount.setContent($L("#{i} of #{n}").replace("#{i}", this.findIndex + 1).replace("#{n}", n));
		if (this.doc.family === "sheet") {
			this.$.sheet.showMatch(this.findIndex);
			return;
		}
		var mark = this.finder.select(this.findIndex);
		this.scrollToNode(mark);
	},

	/**
		inNode brought into view, a third down the screen; or (inAtBottom) its
		bottom a sixth of the screen above the bottom, for a box under what it
		edits: clear of the toolbar and the gesture area, with as much as fits of
		what it edits above it.
	*/
	scrollToNode: function(inNode, inAtBottom) {
		if (!inNode) {
			return;
		}
		var s = this.$.scroller;
		var view = s.hasNode();
		var content = this.$.content.hasNode();
		var r = inNode.getBoundingClientRect(), c = content.getBoundingClientRect(), v = view.getBoundingClientRect();
		var y = r.top - c.top, x = r.left - c.left;
		// The boundaries are the scroll positions' limits (Enyo's BasicScroller), not the content's size.
		var b = s.getBoundaries();
		var top = Math.max(0, Math.min(b.bottom, inAtBottom ? y + r.height + Math.max(24, v.height / 6) - v.height : y - v.height / 3));
		var left = s.scrollLeft;
		if (x < left || x + r.width > left + v.width) {
			left = Math.max(0, Math.min(b.right, x - v.width / 3));
		}
		s.setScrollPositionDirect(left, top);
	},

	findNext: function() {
		this.showFind(this.findIndex + 1);
	},

	findPrev: function() {
		this.showFind(this.findIndex - 1);
	},

	// --- Links -----------------------------------------------------------

	contentClick: function(inSender, inEvent) {
		var t = inEvent.target;
		if (this.editMode) {
			// In edit mode a tap edits the paragraph; links are not followed.
			while (t && t !== this.$.content.hasNode()) {
				if (t.classList && (t.classList.contains("doc-para-editor") || t.classList.contains("doc-text-host"))) {
					return true;
				}
				if (t.getAttribute && t.getAttribute("data-pid") !== null) {
					this.startParagraph(t);
					return true;
				}
				t = t.parentNode;
			}
			return true;
		}
		while (t && t !== this.$.content.hasNode()) {
			if (t.getAttribute && t.getAttribute("data-href") !== null) {
				this.linkTapped(t.getAttribute("data-href"));
				return true;
			}
			t = t.parentNode;
		}
	},

	linkTapped: function(inHref) {
		var href = String(inHref || "");
		if (href.charAt(0) === "#") {
			// A link inside the document (a bookmark or heading).
			var id = href.slice(1);
			var target = null;
			try {
				var safe = id.replace(/["\\]/g, "");
				target = this.$.content.hasNode().querySelector("[data-anchor=\"" + safe + "\"], [id=\"" + safe + "\"], a[name=\"" +
					safe + "\"]");
			} catch (e) {}
			if (target) {
				this.scrollToNode(target);
			}
			return;
		}
		if (!/^(https?:\/\/|mailto:)/i.test(href)) {
			return;
		}
		this.pendingLink = href;
		this.$.linkDialog.openAtCenter();
		this.$.linkText.setContent(enyo.string.escapeHtml(href));
	},

	openLink: function() {
		this.$.linkDialog.close();
		if (this.pendingLink) {
			this.doOpenLink(this.pendingLink);
		}
		this.pendingLink = null;
	},

	closeLink: function() {
		this.$.linkDialog.close();
		this.pendingLink = null;
	},

	// --- Info ------------------------------------------------------------

	showInfo: function() {
		if (!this.doc) {
			return;
		}
		var d = this.doc;
		var meta = d.meta || {};
		var groups = [
			[$L("File"), [
				[$L("Folder"), DocLibrary.dirName(this.path).replace(/^\/media\/internal/, $L("USB storage"))],
				[$L("Size"), DocLibrary.formatSize(this.size)],
				[$L("Opened in"), this.timings.total ? this.timings.total + " ms" : ""]
			]],
			[$L("Document"), [
				[$L("Title"), meta.title],
				[$L("Subject"), meta.subject],
				[$L("Description"), meta.description],
				[$L("Author"), meta.author],
				[$L("Created by"), meta.createdBy],
				[$L("Created"), this.infoDate(meta.created)],
				[$L("Modified by"), meta.modifiedBy],
				[$L("Modified"), this.infoDate(meta.modified)],
				[$L("Pages"), meta.pages],
				[$L("Words"), meta.words],
				[$L("Encoding"), meta.encoding],
				[$L("Separator"), meta.delimiter],
				[$L("Made with"), this.infoGenerator(meta.generator)]
			]]
		];
		var dialog = this.$.infoDialog;
		// Filled in before it opens: it is centred on the size it has then.
		dialog.validateComponents();
		dialog.setCaption(this.fileName);
		this.$.infoSubtitle.setContent(this.describe());
		this.$.infoGroups.destroyControls();
		for (var g = 0; g < groups.length; g++) {
			var rows = [];
			for (var i = 0; i < groups[g][1].length; i++) {
				var v = groups[g][1][i][1];
				if (v !== undefined && v !== null && v !== "") {
					rows.push({kind: enyo.HFlexBox, className: "viewer-info-row", components: [
						{className: "viewer-info-label", content: groups[g][1][i][0]},
						{flex: 1, className: "viewer-info-value", content: String(v)}
					]});
				}
			}
			if (rows.length) {
				this.$.infoGroups.createComponent({kind: "RowGroup", caption: groups[g][0], components: rows});
			}
		}
		this.$.infoGroups.render();
		this.$.infoScroller.applyStyle("height", null);
		dialog.openAtCenter();
		this.fitInfo();
	},

	//* The groups as tall as they are, or as fit the screen with the rest of the dialog; centred again
	fitInfo: function() {
		DialogFit.fit(this.$.infoDialog, this.$.infoScroller, this.$.infoGroups);
	},

	//* "LibreOffice/26.2.6.3$Linux_X86_64 LibreOffice_project/..." as "LibreOffice 26.2.6.3"
	infoGenerator: function(inValue) {
		var m = /^([^\/$]+)\/([^\s$]+)/.exec(inValue || "");
		return m ? m[1] + " " + m[2] : inValue;
	},

	//* A date as the document stores it (ISO 8601), in the device's own words; anything else as it is
	infoDate: function(inValue) {
		var t = inValue ? Date.parse(inValue) : NaN;
		return isNaN(t) ? inValue : new Date(t).toLocaleString();
	},

	closeInfo: function() {
		this.$.infoDialog.close();
	},

	// --- Editing ---------------------------------------------------------

	sheetDirty: function(inSender, inCount) {
		this.edited = inCount > 0;
		this.$.saveButton.setShowing(this.edited);
		this.resized();
	},

	//* A text document was changed: it can be saved as a copy.
	markEdited: function() {
		if (!this.edited) {
			this.edited = true;
			this.$.saveButton.show();
			this.resized();
		}
		// Find's marks were made on the old text.
		this.finder = null;
	},

	saveCopy: function() {
		if (this.doc && this.doc.family === "sheet") {
			this.$.sheet.commitEdit();
			this.doSaveCopy();
			return;
		}
		var self = this;
		this.commitParagraph(function() {
			self.doSaveCopy();
		});
	},

	//* The edits and the format a copy is written in (the document's own; xls as xlsx).
	saveRequest: function() {
		var kind = this.doc ? this.doc.kind : "";
		if (this.$.lok.isOpen()) {
			return {format: this.$.lok.copyFormat(), lok: true};
		}
		if (kind === "odt" || kind === "docx" || kind === "odp" || kind === "pptx") {
			return {format: kind};
		}
		if (kind === "txt" || kind === "md") {
			return {format: kind, text: this.currentText()};
		}
		var format = kind === "ods" ? "ods" : kind === "csv" ? "csv" : "xlsx";
		return {edits: this.$.sheet.edits.slice(0), format: format, sheet: this.$.sheet.currentSheet() ? this.$.sheet.currentSheet().index : 0};
	},

	//* The edit tool, when this app's service (which writes the copy) is there.
	offerEditing: function() {
		var token = this.loadToken, self = this;
		FileService.serviceAvailable(function(inOk) {
			if (token === self.loadToken && inOk) {
				self.canEdit = true;
				self.$.editButton.show();
				self.resized();
			}
		});
	},

	//* The paragraphs of a presentation's slides that can be edited.
	indexSlides: function(inSlides) {
		for (var i = 0; i < inSlides.length; i++) {
			for (var k = 0; k < inSlides[i].items.length; k++) {
				this.indexParagraphs(inSlides[i].items[k].b);
			}
		}
	},

	//* Every paragraph of the model that can be edited, by id.
	indexParagraphs: function(inBlocks) {
		var map = this.paraBlocks;
		var walk = function(inList) {
			for (var i = 0; i < (inList || []).length; i++) {
				var b = inList[i];
				if (b.t === "p" && b.id !== undefined) {
					map[b.id] = b;
				}
				if (b.t === "list") {
					for (var j = 0; j < b.items.length; j++) {
						walk(b.items[j]);
					}
				} else if (b.t === "table") {
					for (var r = 0; r < b.rows.length; r++) {
						for (var c = 0; c < b.rows[r].cells.length; c++) {
							walk(b.rows[r].cells[c].b);
						}
					}
				} else if (b.t === "frame") {
					walk(b.b);
				} else if (b.t === "notes") {
					for (var k = 0; k < b.items.length; k++) {
						walk(b.items[k].b);
					}
				}
				if (b.r) {
					// Text boxes in a paragraph hold paragraphs too.
					for (var q = 0; q < b.r.length; q++) {
						if (b.r[q].t === "frame") {
							walk(b.r[q].b);
						}
					}
				}
			}
		};
		walk(inBlocks);
	},

	toggleEdit: function() {
		if (!this.canEdit || !this.doc) {
			return;
		}
		var self = this;
		if (this.doc.kind === "docx") {
			this.toggleDocxEdit();
			return;
		}
		if (this.editMode) {
			this.commitParagraph(function() {
				self.leaveEdit();
			});
			return;
		}
		this.closeFind();
		this.editMode = true;
		this.$.editButton.addRemoveClass("viewer-tool-on", true);
		this.$.content.addRemoveClass("doc-editing", true);
		if (this.doc.kind === "txt" || this.doc.kind === "md") {
			this.startTextEdit();
		}
	},

	//* A Word document: the .docx editor in place of the page (loaded the first time), editing or read only.
	toggleDocxEdit: function() {
		var self = this, token = this.loadToken;
		var docx = this.$.docx;
		if (docx.isOpen()) {
			this.setDocxEditing(!this.editMode);
			return;
		}
		this.closeFind();
		this.showStatus($L("Opening the editor\u2026"), true);
		docx.open(this.docBytes, function(inErr) {
			if (token !== self.loadToken) {
				return;
			}
			self.hideStatus();
			if (inErr) {
				// The document stays as it was shown; the message goes after a while.
				docx.close();
				self.showStatus(inErr.message, false);
				setTimeout(function() {
					if (token === self.loadToken) {
						self.hideStatus();
					}
				}, 5000);
				return;
			}
			self.$.scroller.hide();
			docx.show();
			self.setDocxEditing(true);
			self.updateTools();
		});
	},

	setDocxEditing: function(inEditing) {
		this.editMode = inEditing;
		this.$.editButton.addRemoveClass("viewer-tool-on", inEditing);
		this.$.docx.setEditing(inEditing);
	},

	//* The bytes of the copy: from the .docx editor, else from the engine.
	makeCopy: function(inRequest, inCallback) {
		if (inRequest.lok) {
			this.$.lok.copyBytes(inCallback);
			return;
		}
		if (inRequest.format === "docx") {
			this.$.docx.copyBytes(inCallback);
			return;
		}
		DocLoader.save(inRequest, inCallback);
	},

	// --- Editing in LibreOffice (LokEditor.js) -----------------------------

	/**
		The document in a LibreOffice editing session, in place of the page:
		any format LibreOffice reads, presentations included. Offered (in
		the app menu) when the office service has sessions.
	*/
	editWithLibreOffice: function() {
		var self = this, token = this.loadToken;
		if (!this.doc || !this.path || this.$.lok.isOpen()) {
			return;
		}
		this.closeFind();
		this.showStatus($L("Opening in LibreOffice\u2026"), true);
		this.$.lok.open(this.path, function(inErr) {
			if (token !== self.loadToken) {
				return;
			}
			self.hideStatus();
			if (inErr) {
				self.$.lok.close();
				self.showStatus(inErr.message, false);
				setTimeout(function() {
					if (token === self.loadToken) {
						self.hideStatus();
					}
				}, 5000);
				return;
			}
			self.$.scroller.hide();
			self.$.sheet.hide();
			self.$.docx.hide();
			self.$.lok.show();
			self.$.lok.showPart();
			self.updateTools();
		});
	},

	lokChanged: function() {
		this.lokEdited = true;
		this.markEdited();
	},

	//* Back to the document as it was shown; LibreOffice's edits close with the session
	closeLok: function() {
		this.$.lok.close();
		this.lokClosed();
	},

	lokClosed: function() {
		this.$.lok.hide();
		if (this.lokEdited) {
			this.lokEdited = false;
			this.edited = false;
			this.$.saveButton.hide();
			this.resized();
		}
		var sheet = this.doc && this.doc.family === "sheet";
		this.$.sheet.setShowing(!!sheet);
		this.$.scroller.setShowing(!sheet && !this.$.docx.isOpen());
		this.$.docx.setShowing(this.$.docx.isOpen());
		this.updateTools();
	},

	leaveEdit: function() {
		this.editMode = false;
		this.$.editButton.addRemoveClass("viewer-tool-on", false);
		this.$.content.addRemoveClass("doc-editing", false);
		if (this.textInput) {
			var model = this.doc.model;
			model.text = this.currentText();
			if (this.doc.kind === "md") {
				model.html = DocText.markdown(model.text);
			}
			this.textInput = null;
			this.dropEditorHost(this.textHost);
			this.textHost = null;
			this.renderPlain(this.doc.kind, model);
			this.applyLayout();
		}
	},

	//* The text of a plain text or Markdown file as it is now.
	currentText: function() {
		if (this.textInput) {
			return this.textInput.value.replace(/\r\n?/g, "\n");
		}
		return this.doc && this.doc.model ? String(this.doc.model.text || "") : "";
	},

	//* All of a plain text or Markdown file in one text box.
	/**
		A node for a text box, inside an Enyo control of its own (as Enyo's
		own inputs: requiresDomMousedown off on a device, where Enyo focuses
		the box itself; on in a desktop browser). The box goes in inParent,
		before inBefore (or at its end).
	*/
	editorHost: function(inParent, inBefore, inClass) {
		var place = document.createElement("div");
		inParent.insertBefore(place, inBefore || null);
		var control = this.createComponent({kind: enyo.Control, className: inClass, requiresDomMousedown: !window.PalmSystem, owner: this});
		control.renderInto(place);
		return {control: control, place: place, node: control.hasNode()};
	},

	//* Removes a text box made by editorHost.
	dropEditorHost: function(inHost) {
		if (!inHost) {
			return;
		}
		inHost.control.destroy();
		if (inHost.place.parentNode) {
			inHost.place.parentNode.removeChild(inHost.place);
		}
	},

	startTextEdit: function() {
		var content = this.$.content.hasNode();
		var self = this;
		content.innerHTML = "";
		content.className = "viewer-content doc-plain doc-editing";
		this.dropEditorHost(this.textHost);
		this.textHost = this.editorHost(content, null, "doc-text-host");
		var ta = DocRender.el("textarea", "doc-text-input doc-plaintext", this.textHost.node);
		ta.value = String(this.doc.model.text || "");
		ta.setAttribute("spellcheck", "false");
		ta.addEventListener("input", function() {
			self.markEdited();
			self.fitTextArea(ta);
		});
		this.caretOnTap(ta);
		this.textInput = ta;
		this.fitTextArea(ta);
		ta.focus();
	},

	//* A text box as tall as its text, so the page scrolls rather than the box.
	/**
		inFn on a tap of a DOM button. On a device a touch makes no DOM click
		(Enyo turns the browser's own off and dispatches its own, which a
		DOM listener does not hear), so the touch itself counts; in a desktop
		browser the click does. One tap is one call.
	*/
	onTap: function(inNode, inFn) {
		var last = 0;
		var fire = function() {
			var now = Date.now();
			if (now - last > 500) {
				last = now;
				inFn();
			}
		};
		inNode.addEventListener("touchend", function(e) {
			e.preventDefault();
			fire();
		});
		inNode.addEventListener("click", fire);
	},

	/**
		A tap in a text box puts the cursor under the finger. Enyo focuses the
		box and leaves placing the cursor to PalmSystem.simulateMouseClick,
		which on LuneOS's WAM sends no events: the cursor would stay where it
		was (at the end).
	*/
	caretOnTap: function(inArea) {
		var self = this;
		var start = null;
		inArea.addEventListener("touchstart", function(e) {
			var t = e.touches && e.touches[0];
			start = t ? {x: t.clientX, y: t.clientY} : null;
		});
		inArea.addEventListener("touchend", function(e) {
			var t = e.changedTouches && e.changedTouches[0];
			// A tap, not a drag (which scrolls).
			if (!t || !start || Math.abs(t.clientX - start.x) > 8 || Math.abs(t.clientY - start.y) > 8) {
				return;
			}
			var x = t.clientX, y = t.clientY;
			setTimeout(function() {
				var at = self.offsetAtPoint(inArea, x, y);
				if (at !== null) {
					if (document.activeElement !== inArea) {
						inArea.focus();
					}
					inArea.setSelectionRange(at, at);
				}
			}, 0);
		});
	},

	/**
		The character offset in a text box at a point of the screen: a copy
		of its text, laid out as the box lays it out, is put over it for a
		moment and asked (caretRangeFromPoint cannot look into a textarea).
	*/
	offsetAtPoint: function(inArea, inX, inY) {
		if (!document.caretRangeFromPoint || !inArea.parentNode) {
			return null;
		}
		var cs = window.getComputedStyle(inArea);
		var copy = document.createElement("div");
		["fontFamily", "fontSize", "fontWeight", "fontStyle", "letterSpacing", "lineHeight", "textAlign", "textIndent",
			"paddingTop", "paddingRight", "paddingBottom", "paddingLeft", "borderTopWidth", "borderRightWidth",
			"borderBottomWidth", "borderLeftWidth", "boxSizing", "wordSpacing", "tabSize"].forEach(function(p) {
			copy.style[p] = cs[p];
		});
		copy.style.position = "absolute";
		copy.style.left = inArea.offsetLeft + "px";
		copy.style.top = (inArea.offsetTop - inArea.scrollTop) + "px";
		copy.style.width = inArea.offsetWidth + "px";
		copy.style.borderStyle = "solid";
		copy.style.borderColor = "transparent";
		copy.style.whiteSpace = "pre-wrap";
		copy.style.overflowWrap = "break-word";
		copy.style.wordWrap = "break-word";
		copy.style.zIndex = "100000";
		copy.style.color = "transparent";
		copy.style.background = "transparent";
		copy.appendChild(document.createTextNode(inArea.value + "\u200b"));
		var host = inArea.offsetParent || inArea.parentNode;
		host.appendChild(copy);
		var r = document.caretRangeFromPoint(inX, inY);
		var at = r && r.startContainer === copy.firstChild ? Math.min(r.startOffset, inArea.value.length) : null;
		host.removeChild(copy);
		return at;
	},

	fitTextArea: function(inArea) {
		inArea.style.height = "auto";
		inArea.style.height = (inArea.scrollHeight + 4) + "px";
		this.$.scroller.calcBoundaries && this.$.scroller.calcBoundaries();
	},

	//* A tap on a paragraph in edit mode: its text in a box in its place.
	startParagraph: function(inNode) {
		var id = +inNode.getAttribute("data-pid");
		var block = this.paraBlocks[id];
		if (!block) {
			return;
		}
		var self = this;
		this.commitParagraph(function() {
			if (!inNode.parentNode) {
				return;
			}
			// On a slide (drawn smaller or larger) the box goes under the slide, at the page's own size; the
			// paragraph stays in view, marked.
			var frame = self.slideBoxes ? self.closestClass(inNode, "doc-slide-frame") : null;
			var host = frame ? self.editorHost(frame.parentNode, frame.nextSibling, "doc-para-editor doc-slide-editor") :
				self.editorHost(inNode.parentNode, inNode, "doc-para-editor");
			var box = host.node;
			var ta = DocRender.el("textarea", "doc-para-input", box);
			ta.value = block.pt;
			ta.setAttribute("spellcheck", "false");
			var cs = window.getComputedStyle(inNode);
			ta.style.fontSize = cs.fontSize;
			ta.style.fontFamily = cs.fontFamily;
			ta.style.fontWeight = cs.fontWeight;
			ta.style.textAlign = cs.textAlign;
			ta.addEventListener("input", function() {
				self.fitTextArea(ta);
			});
			self.caretOnTap(ta);
			var bar = DocRender.el("div", "doc-para-bar", box);
			var err = DocRender.el("div", "doc-para-error", bar);
			var cancel = DocRender.el("button", "doc-para-button", bar);
			cancel.appendChild(document.createTextNode($L("Cancel")));
			var done = DocRender.el("button", "doc-para-button doc-para-done", bar);
			done.appendChild(document.createTextNode($L("Done")));
			self.onTap(cancel, function() {
				self.cancelParagraph();
			});
			self.onTap(done, function() {
				self.commitParagraph();
			});
			if (frame) {
				inNode.classList.add("doc-para-editing");
			} else {
				inNode.style.display = "none";
			}
			self.paraEditor = {id: id, node: inNode, host: host, input: ta, error: err, frame: frame};
			self.fitTextArea(ta);
			if (frame) {
				// Under the slide it may be off the screen: the box at the bottom, as much of its slide as fits above.
				self.scrollToNode(host.node, true);
			}
			ta.focus();
		});
	},

	closestClass: function(inNode, inClass) {
		for (var n = inNode; n && n.classList; n = n.parentNode) {
			if (n.classList.contains(inClass)) {
				return n;
			}
		}
		return null;
	},

	//* Slide inIndex drawn again from inModel (after an edit), its paragraphs indexed again.
	redrawSlide: function(inIndex, inModel) {
		var box = this.slideBoxes && this.slideBoxes[inIndex];
		if (!box) {
			return;
		}
		var holder = document.createElement("div");
		var fresh = DocRender.slides(holder, [inModel], this.renderCtx || {images: {}}, inIndex)[0];
		box.frame.parentNode.replaceChild(fresh.frame, box.frame);
		this.slideBoxes[inIndex] = fresh;
		this.indexSlides([inModel]);
		this.applyLayout();
	},

	cancelParagraph: function() {
		var ed = this.paraEditor;
		if (!ed) {
			return false;
		}
		this.paraEditor = null;
		ed.node.style.display = "";
		ed.node.classList.remove("doc-para-editing");
		this.dropEditorHost(ed.host);
		return true;
	},

	/**
		Ends the paragraph being edited: its new text goes to the engine,
		and the paragraph it gives back takes the old one's place. Then
		inDone, also when no paragraph was open; not when it failed (the
		box stays, with the error).
	*/
	commitParagraph: function(inDone) {
		var ed = this.paraEditor;
		var done = inDone || function() {};
		if (!ed) {
			done();
			return;
		}
		var text = ed.input.value.replace(/\r\n?/g, "\n");
		if (text === this.paraBlocks[ed.id].pt) {
			this.cancelParagraph();
			done();
			return;
		}
		var self = this, token = this.loadToken;
		ed.input.disabled = true;
		DocLoader.edit({para: {id: ed.id, text: text}}, function(inErr, inResult) {
			if (token !== self.loadToken || self.paraEditor !== ed) {
				return;
			}
			ed.input.disabled = false;
			if (inErr) {
				ed.error.textContent = inErr.message;
				return;
			}
			var block = inResult && inResult.block;
			self.paraEditor = null;
			ed.node.classList.remove("doc-para-editing");
			if (inResult && inResult.slide) {
				// A .pptx slide is read again whole: drawn again in its place.
				self.redrawSlide(inResult.slide.index, inResult.slide.model);
				self.markEdited();
			} else if (block) {
				self.paraBlocks[ed.id] = block;
				var node = DocRender.paragraph(block, self.renderCtx || {images: {}});
				ed.node.parentNode.replaceChild(node, ed.node);
				self.markEdited();
			} else {
				ed.node.style.display = "";
			}
			self.dropEditorHost(ed.host);
			self.$.scroller.calcBoundaries();
			done();
		});
	},

	//* Back: closes what is open in the viewer first; false when nothing was.
	back: function() {
		if (this.$.linkDialog.isOpen || this.$.infoDialog.isOpen) {
			this.$.linkDialog.close();
			this.$.infoDialog.close();
			return true;
		}
		if (this.$.sheet.cancelEdit()) {
			return true;
		}
		if (this.cancelParagraph()) {
			return true;
		}
		if (this.$.lok.isOpen()) {
			this.closeLok();
			return true;
		}
		if (this.editMode && this.$.docx.isOpen()) {
			this.setDocxEditing(false);
			return true;
		}
		if (this.editMode) {
			this.leaveEdit();
			return true;
		}
		return this.closeFind();
	},

	close: function() {
		this.loadToken++;
		this.reset();
		this.hideStatus();
	}
});
