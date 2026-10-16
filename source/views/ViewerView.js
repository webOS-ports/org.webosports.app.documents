/**
	Shows one document: text documents as pages (or reflowed to the screen
	on a phone), presentations as slides one under the other, workbooks in
	a grid. Zoom, find, night reading and document info for all of them;
	light cell editing and "Save a copy" for workbooks.

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
			{kind: "ToolButton", name: "infoButton", className: "viewer-tool", icon: "images/info.svg", onclick: "showInfo"}
		]},
		{name: "infoDialog", kind: "ModalDialog", caption: $L("Document Info"), components: [
			{name: "infoText", className: "viewer-info", allowHtml: true, content: ""},
			{kind: "Button", caption: $L("Close"), onclick: "closeInfo"}
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
		if (inResult.kind === "docx" && !inResult.model) {
			this.renderDocx(content, inBuffer, inDone);
			return;
		}
		var ctx = {images: this.imageUrls(inResult.images || {})};
		if (inResult.kind === "txt") {
			content.className = "viewer-content doc-plain";
			var pre = DocRender.el("div", "doc-plaintext", content);
			pre.appendChild(document.createTextNode(inResult.model.text));
		} else if (inResult.kind === "md") {
			content.className = "viewer-content doc-markdown";
			var md = DocRender.el("div", "doc-markdown-body", content);
			// DocText.markdown escapes all of the input; its own tags are the only markup.
			md.innerHTML = inResult.model.html;
			DocRender.neutraliseLinks(md);
		} else if (family === "slides") {
			content.className = "viewer-content doc-slides";
			this.slideBoxes = DocRender.slides(content, inResult.model.slides, ctx);
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
		var reflowable = has && (this.doc.kind === "docx" || this.doc.kind === "odt");
		this.$.zoomOut.setDisabled(!has);
		this.$.zoomIn.setDisabled(!has);
		this.$.zoomLabel.setDisabled(!has);
		this.$.findButton.setDisabled(!has);
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
		this.$.zoomLabel.setCaption(Math.round(this.zoom * 100) + "%");
	},

	stepZoom: function(inDirection) {
		var steps = this.zoomSteps;
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
			var left = Math.max(0, Math.min(b.right - v.clientWidth, cx * this.zoom - v.clientWidth / 2));
			var top = Math.max(0, Math.min(b.bottom - v.clientHeight, cy * this.zoom - v.clientHeight / 2));
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

	scrollToNode: function(inNode) {
		if (!inNode) {
			return;
		}
		var s = this.$.scroller;
		var view = s.hasNode();
		var content = this.$.content.hasNode();
		var r = inNode.getBoundingClientRect(), c = content.getBoundingClientRect(), v = view.getBoundingClientRect();
		var y = r.top - c.top, x = r.left - c.left;
		var b = s.getBoundaries();
		var top = Math.max(0, Math.min(b.bottom - v.height, y - v.height / 3));
		var left = s.scrollLeft;
		if (x < left || x + r.width > left + v.width) {
			left = Math.max(0, Math.min(b.right - v.width, x - v.width / 3));
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
		var rows = [
			[$L("File"), this.fileName],
			[$L("Folder"), DocLibrary.dirName(this.path)],
			[$L("Type"), FileTypes.kinds[d.kind] ? FileTypes.kinds[d.kind].label : d.kind],
			[$L("Size"), DocLibrary.formatSize(this.size)],
			[$L("Title"), meta.title],
			[$L("Author"), meta.author],
			[$L("Subject"), meta.subject],
			[$L("Created"), meta.created],
			[$L("Modified"), meta.modified],
			[$L("Pages"), meta.pages],
			[$L("Words"), meta.words],
			[$L("Encoding"), meta.encoding],
			[$L("Separator"), meta.delimiter],
			[$L("Opened in"), this.timings.total ? this.timings.total + " ms" : ""]
		];
		var html = "<table>";
		for (var i = 0; i < rows.length; i++) {
			if (rows[i][1] !== undefined && rows[i][1] !== null && rows[i][1] !== "") {
				html += "<tr><th>" + enyo.string.escapeHtml(rows[i][0]) + "</th><td>" + enyo.string.escapeHtml(String(rows[i][1])) + "</td></tr>";
			}
		}
		html += "</table>";
		this.$.infoDialog.openAtCenter();
		this.$.infoText.setContent(html);
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

	saveCopy: function() {
		this.$.sheet.commitEdit();
		this.doSaveCopy();
	},

	//* The edits and the format a copy is written in (the document's own, xls as xlsx).
	saveRequest: function() {
		var kind = this.doc ? this.doc.kind : "";
		var format = kind === "ods" ? "ods" : kind === "csv" ? "csv" : "xlsx";
		return {edits: this.$.sheet.edits.slice(0), format: format, sheet: this.$.sheet.currentSheet() ? this.$.sheet.currentSheet().index : 0};
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
		return this.closeFind();
	},

	close: function() {
		this.loadToken++;
		this.reset();
		this.hideStatus();
	}
});
