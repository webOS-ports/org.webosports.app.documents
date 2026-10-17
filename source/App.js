/**
	Documents: reads Word, OpenDocument, Excel, PowerPoint, CSV, text and
	Markdown files on LuneOS.

	Two screens in a Pane: the library (every document on the device, the
	recent ones first) and the viewer. Opened from another app with
	{"target": "file:///media/internal/..."} it goes straight to that
	document; the back gesture then returns to the library.
*/
enyo.kind({
	name: "DocumentsApp",
	kind: enyo.VFlexBox,
	className: "documents-app",
	root: "/media/internal",
	rescanAfter: 5 * 60 * 1000,
	components: [
		{kind: "ApplicationEvents", onBack: "goBack", onApplicationRelaunch: "relaunched", onWindowActivated: "activated",
			onKeydown: "keyDown"},
		{kind: "AppMenu", onBeforeOpen: "menuOpening", components: [
			{name: "menuFind", caption: $L("Find"), onclick: "menuFind"},
			{name: "menuNight", caption: $L("Night Reading"), onclick: "menuNight"},
			{name: "menuInfo", caption: $L("Document Info"), onclick: "menuInfo"},
			{name: "menuFreeze", caption: $L("Freeze First Row"), onclick: "menuFreeze"},
			{name: "menuLok", caption: $L("Edit with LibreOffice"), onclick: "menuLok", showing: false},
			{name: "menuRefresh", caption: $L("Look for Documents Again"), onclick: "rescan"},
			{caption: $L("About Documents"), onclick: "showAbout"}
		]},
		{name: "pane", kind: "Pane", flex: 1, transitionKind: "enyo.transitions.Simple", onSelectView: "viewSelected", components: [
			{name: "library", kind: "LibraryView", onOpenDocument: "openDocument", onRefresh: "rescan"},
			{name: "viewer", kind: "ViewerView", lazy: false, onOpenLink: "openLink", onSaveCopy: "askSaveCopy",
				onNight: "nightChanged", onOpened: "documentOpened", onHandOff: "handOffFromViewer"}
		]},
		{name: "saveDialog", kind: "ModalDialog", caption: $L("Save a Copy"), components: [
			{name: "saveNote", className: "save-note", content: ""},
			{name: "saveName", kind: "Input", hint: $L("File name"), autoCapitalize: "lowercase", autocorrect: false,
				spellcheck: false, autoWordComplete: false},
			{name: "saveError", className: "save-error", showing: false, content: ""},
			{kind: enyo.HFlexBox, components: [
				{kind: "Button", flex: 1, caption: $L("Cancel"), onclick: "closeSave"},
				{name: "saveOk", kind: "ActivityButton", flex: 1, className: "enyo-button-affirmative", caption: $L("Save"), onclick: "doSave"}
			]}
		]},
		{name: "about", kind: "ModalDialog", caption: $L("About Documents"), components: [
			{name: "aboutText", className: "about-text", allowHtml: true, content: ""},
			{kind: "Button", caption: $L("OK"), onclick: "closeAbout"}
		]},
		{name: "toaster", className: "documents-toast", showing: false}
	],

	create: function() {
		this.inherited(arguments);
		this.prefs = this.loadPrefs();
		this.$.library.setRecent(this.prefs.recent);
		this.$.viewer.setNight(this.prefs.night);
		this.$.viewer.freezeHeader = this.prefs.freeze;
		this.addRemoveClass("night", !!this.prefs.night);
		this.lastScan = 0;
		this.docs = [];
	},

	rendered: function() {
		this.inherited(arguments);
		var target = this.launchTarget(enyo.windowParams);
		if (target) {
			this.openDocument(null, target);
		}
		// The library is filled in the background either way.
		this.rescan();
		this.findConverter();
	},

	/**
		Whether LibreOffice's service (org.webosports.service.office) is
		installed: asked without a document, it answers "badRequest". Then
		.doc, .ppt and .rtf are listed too, and converted when opened.
	*/
	findConverter: function() {
		var self = this;
		FileService.call("luna://org.webosports.service.office/convert", {}, function(inErr) {
			var present = !!inErr && inErr.code === "badRequest";
			if (present && !FileTypes.converter) {
				FileTypes.converter = true;
				self.rescan();
			}
		});
		// Editing sessions: a service that has them refuses an open without a document the same way
		FileService.call("luna://org.webosports.service.office/open", {}, function(inErr) {
			FileTypes.lokSessions = !!inErr && inErr.code === "badRequest";
		});
	},

	menuLok: function() {
		this.$.viewer.editWithLibreOffice();
	},

	//* The local path a launch asks for, or null.
	launchTarget: function(inParams) {
		var p = inParams;
		if (typeof p === "string") {
			try {
				p = JSON.parse(p);
			} catch (e) {
				p = {};
			}
		}
		return p && p.target ? DocLibrary.targetToPath(p.target) : null;
	},

	//* ApplicationEvents passes the event; its params are the new launch parameters.
	relaunched: function(inSender, inEvent) {
		var target = this.launchTarget((inEvent && inEvent.params) || enyo.windowParams);
		if (target) {
			this.openDocument(null, target);
		}
		return true;
	},

	activated: function() {
		if (Date.now() - this.lastScan > this.rescanAfter && !this.scanner) {
			this.rescan();
		}
	},

	// --- Preferences -----------------------------------------------------

	loadPrefs: function() {
		var prefs = {night: false, recent: [], freeze: true};
		try {
			var stored = JSON.parse(localStorage.getItem("documents.prefs") || "{}");
			if (stored && typeof stored === "object") {
				prefs.night = !!stored.night;
				prefs.freeze = stored.freeze !== false;
				prefs.recent = Array.isArray(stored.recent) ? stored.recent.filter(function(r) {
					return r && typeof r.path === "string" && typeof r.name === "string";
				}).slice(0, 10) : [];
			}
		} catch (e) {}
		return prefs;
	},

	savePrefs: function() {
		try {
			localStorage.setItem("documents.prefs", JSON.stringify(this.prefs));
		} catch (e) {}
	},

	// --- Library ---------------------------------------------------------

	rescan: function() {
		if (this.scanner) {
			this.scanner.cancel();
		}
		var self = this;
		this.$.library.setScanning(true);
		var scanner = this.scanner = new DocLibrary.Scanner({
			list: function(inDir, inCallback) {
				FileService.list(inDir, inCallback);
			},
			exists: function(inPath, inCallback) {
				FileService.exists(inPath, inCallback);
			},
			onProgress: function(inFound) {
				if (scanner === self.scanner && inFound !== self.lastFound) {
					self.lastFound = inFound;
					self.$.library.setDocuments(scanner.docs.slice(0));
				}
			},
			onDone: function(inDocs) {
				if (scanner !== self.scanner) {
					return;
				}
				self.scanner = null;
				self.lastScan = Date.now();
				self.docs = inDocs;
				self.$.library.setScanning(false);
				self.$.library.setDocuments(inDocs);
			}
		});
		scanner.start(this.root);
	},

	// --- Viewer ----------------------------------------------------------

	openDocument: function(inSender, inPath) {
		var mime = FileTypes.handOffMime(FileTypes.kindFromName(inPath));
		if (mime) {
			// A PDF goes to the app that shows PDFs; if none answers, the
			// viewer says why it cannot show it either.
			var self = this;
			FileService.openWith(inPath, mime, function(inErr) {
				if (inErr) {
					enyo.warn("no app opened " + inPath + " as " + mime + ": " + inErr.message);
					self.showViewer(inPath);
				}
			});
			return;
		}
		this.showViewer(inPath);
	},

	showViewer: function(inPath) {
		this.$.library.setHidden(true);
		this.$.pane.selectViewByName("viewer");
		this.$.viewer.openPath(inPath);
	},

	//* The viewer found a file another app shows behind a misleading name (a PDF called ".docx").
	handOffFromViewer: function(inSender, inPath, inMime) {
		var self = this;
		FileService.openWith(inPath, inMime, function(inErr) {
			if (inErr) {
				enyo.warn("no app opened " + inPath + " as " + inMime + ": " + inErr.message);
				self.$.viewer.handOffFailed();
			} else if (self.$.pane.getViewName() === "viewer" && self.$.viewer.path === inPath) {
				self.$.pane.selectViewByName("library");
			}
		});
	},

	documentOpened: function(inSender, inPath, inKind) {
		this.prefs.recent = DocLibrary.addRecent(this.prefs.recent, {path: inPath, name: DocLibrary.baseName(inPath),
			kind: FileTypes.kindFromName(inPath), opened: Date.now()});
		this.savePrefs();
		this.$.library.setRecent(this.prefs.recent);

	},

	viewSelected: function(inSender, inView) {
		this.$.library.setHidden(!inView || inView.name !== "library");
		if (inView && inView.name === "library") {
			this.$.viewer.close();
		}
	},

	goBack: function(inSender, inEvent) {
		if (this.$.saveDialog.isOpen) {
			this.closeSave();
		} else if (this.$.about.isOpen) {
			this.closeAbout();
		} else if (this.$.pane.getViewName() === "viewer") {
			if (!this.$.viewer.back()) {
				this.$.pane.selectViewByName("library");
			}
		} else {
			// On the library, back is the system's (the card is minimised).
			return;
		}
		if (inEvent && inEvent.preventDefault) {
			inEvent.preventDefault();
		}
		return true;
	},

	keyDown: function(inSender, inEvent) {
		if (this.$.pane.getViewName() !== "viewer") {
			return;
		}
		var k = inEvent.keyCode;
		if (inEvent.ctrlKey && k === 70) {
			this.$.viewer.openFind();
			inEvent.preventDefault();
		} else if (inEvent.ctrlKey && (k === 187 || k === 61 || k === 107)) {
			this.$.viewer.zoomIn();
			inEvent.preventDefault();
		} else if (inEvent.ctrlKey && (k === 189 || k === 173 || k === 109)) {
			this.$.viewer.zoomOut();
			inEvent.preventDefault();
		} else if (inEvent.ctrlKey && k === 48) {
			this.$.viewer.zoomFit();
			inEvent.preventDefault();
		}
	},

	openLink: function(inSender, inUrl) {
		FileService.openUrl(inUrl);
	},

	nightChanged: function(inSender, inNight) {
		this.prefs.night = inNight;
		this.addRemoveClass("night", inNight);
		this.savePrefs();
	},

	// --- App menu --------------------------------------------------------

	menuOpening: function() {
		var inViewer = this.$.pane.getViewName() === "viewer" && !!this.$.viewer.doc;
		var sheet = inViewer && this.$.viewer.doc.family === "sheet";
		this.$.menuFind.setDisabled(!inViewer);
		this.$.menuInfo.setDisabled(!inViewer);
		this.$.menuFreeze.setShowing(sheet);
		// LibreOffice edits what it reads, in a session of the office service, when it has them
		this.$.menuLok.setShowing(inViewer && FileTypes.lokSessions && !this.$.viewer.$.lok.isOpen());
		this.$.menuFreeze.setCaption(this.prefs.freeze ? $L("Unfreeze First Row") : $L("Freeze First Row"));
		this.$.menuNight.setCaption(this.prefs.night ? $L("Day Reading") : $L("Night Reading"));
	},

	menuFind: function() {
		this.$.viewer.openFind();
	},

	menuNight: function() {
		this.$.viewer.toggleNight();
	},

	menuInfo: function() {
		this.$.viewer.showInfo();
	},

	menuFreeze: function() {
		this.prefs.freeze = !this.prefs.freeze;
		this.$.viewer.freezeHeader = this.prefs.freeze;
		this.savePrefs();
		this.$.viewer.$.sheet.setFreezeHeader(this.prefs.freeze);
	},

	// --- Save a copy -----------------------------------------------------

	askSaveCopy: function() {
		var v = this.$.viewer;
		var req = v.saveRequest();
		var existing = this.docs.filter(function(d) {
			return d.dir === DocLibrary.dirName(v.path);
		}).map(function(d) {
			return d.name;
		});
		this.$.saveDialog.openAtCenter();
		this.$.saveName.setValue(DocLibrary.copyName(v.fileName, req.format, existing));
		this.$.saveError.hide();
		this.$.saveOk.setActive(false);
		this.$.saveOk.setDisabled(false);
		var note = $L("A new file next to the original; the original is not changed.");
		if (v.doc.kind === "xls") {
			note += " " + $L("It is saved as .xlsx, with the values and formulas but not the formatting, charts or pictures.");
		} else if (v.doc.kind === "txt" || v.doc.kind === "md") {
			note += " " + $L("It is saved as UTF-8.");
		}
		this.$.saveNote.setContent(enyo.string.escapeHtml(note));
	},

	closeSave: function() {
		this.$.saveDialog.close();
	},

	doSave: function() {
		var self = this;
		var v = this.$.viewer;
		var req = v.saveRequest();
		var name = String(this.$.saveName.getValue() || "").replace(/^\s+|\s+$/g, "");
		if (!name || /[\/\\]/.test(name) || name.charAt(0) === ".") {
			this.saveFailed($L("Choose a file name without slashes."));
			return;
		}
		if (FileTypes.extensionOf(name) !== req.format) {
			name += "." + req.format;
		}
		var path = DocLibrary.dirName(v.path) + "/" + name;
		this.$.saveOk.setActive(true);
		this.$.saveOk.setDisabled(true);
		v.makeCopy(req, function(inErr, inBytes) {
			if (inErr) {
				self.saveFailed(inErr.message);
				return;
			}
			FileService.writeNew(path, new Uint8Array(inBytes.buffer || inBytes, inBytes.byteOffset || 0, inBytes.byteLength), function(inErr2) {
				if (inErr2) {
					self.saveFailed(inErr2.code === "EEXIST" || /exists/i.test(inErr2.message) ?
						$L("A file with that name already exists.") : $L("The copy could not be saved: ") + inErr2.message);
					return;
				}
				self.$.saveDialog.close();
				self.toast($L("Saved as #{name}").replace("#{name}", name));
				self.rescan();
			});
		});
	},

	saveFailed: function(inMessage) {
		this.$.saveOk.setActive(false);
		this.$.saveOk.setDisabled(false);
		this.$.saveError.setContent(enyo.string.escapeHtml(inMessage));
		this.$.saveError.show();
	},

	toast: function(inText) {
		var t = this.$.toaster;
		t.setContent(enyo.string.escapeHtml(inText));
		t.show();
		enyo.job("documents-toast", function() {
			t.hide();
		}, 3000);
	},

	// --- About -----------------------------------------------------------

	showAbout: function() {
		this.$.about.openAtCenter();
		this.$.aboutText.setContent(
			"<p>" + enyo.string.escapeHtml($L("Reads Word (.docx), OpenDocument (.odt, .ods, .odp), Excel (.xlsx, .xls), " +
				"PowerPoint (.pptx), CSV, text and Markdown files.")) + "</p>" +
			"<p>" + enyo.string.escapeHtml($L("Built with docx-preview (Apache 2.0), SheetJS Community Edition (Apache 2.0), " +
				"JSZip (MIT), HyperFormula (GPLv3) and SuperDoc (AGPLv3).")) + "</p>" +
			"<p>" + enyo.string.escapeHtml($L("Licensed under the GNU General Public License version 3. WebOS Ports.")) + "</p>");
	},

	closeAbout: function() {
		this.$.about.close();
	}
});
