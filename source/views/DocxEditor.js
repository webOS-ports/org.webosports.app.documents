/**
	Edits a Word document with SuperDoc (lib/superdoc, AGPLv3), which
	reads and writes .docx itself: pages, tables, headers and footers,
	lists and tracked changes stay as Word has them. Its own toolbar is
	not used; the bar here has the common commands, with the app's look.

		editor.open(bytes, function(err) {...});   // loads the library on first use (it is large)
		editor.setEditing(true);                    // or false: read only
		editor.copyBytes(function(err, bytes) {...});
		editor.close();

	Its usage reports ("telemetry") are turned off: nothing is sent.
*/
enyo.kind({
	name: "DocxEditor",
	kind: enyo.VFlexBox,
	className: "docx-editor",
	events: {
		onChanged: "",
		onZoomChanged: ""
	},
	components: [
		{name: "bar", kind: enyo.HFlexBox, className: "docx-editor-bar", align: "center", pack: "center", components: [
			{kind: "ToolButton", className: "viewer-tool", icon: "images/undo.svg", command: "undo", onclick: "run"},
			{kind: "ToolButton", className: "viewer-tool", icon: "images/redo.svg", command: "redo", onclick: "run"},
			{kind: "ToolButton", className: "viewer-tool docx-b", caption: "B", command: "toggleBold", onclick: "run"},
			{kind: "ToolButton", className: "viewer-tool docx-i", caption: "I", command: "toggleItalic", onclick: "run"},
			{kind: "ToolButton", className: "viewer-tool docx-u", caption: "U", command: "toggleUnderline", onclick: "run"},
			{kind: "ToolButton", className: "viewer-tool", icon: "images/list-bullets.svg", command: "toggleBulletList", onclick: "run"},
			{kind: "ToolButton", className: "viewer-tool", icon: "images/list-numbers.svg", command: "toggleOrderedList", onclick: "run"}
		]},
		// The DOM's mousedown, which SuperDoc needs to put the cursor where a tap is: otherwise Enyo's
		// tap handling cancels it and moves the focus off the page (see ViewerView.editorHost).
		{name: "host", flex: 1, className: "docx-editor-host", requiresDomMousedown: true}
	],

	create: function() {
		this.inherited(arguments);
		this.superdoc = null;
		this.editing = false;
		this.openToken = 0;
	},

	/**
		Shows inBytes (an ArrayBuffer of a .docx) in the editor; inCallback(err)
		when it is ready or failed.
	*/
	open: function(inBytes, inCallback) {
		var self = this;
		var token = ++this.openToken;
		this.close();
		this.openToken = token;
		this.original = inBytes;
		DocLoader.loadStyle("lib/superdoc/style.css");
		DocLoader.loadScripts(["lib/jszip/jszip.min.js", "lib/superdoc/superdoc.min.js"], function(inErr) {
			if (token !== self.openToken) {
				return;
			}
			if (inErr || !window.SuperDoc) {
				inCallback(inErr || new Error($L("The editor did not load.")));
				return;
			}
			var host = self.$.host.hasNode();
			host.innerHTML = "";
			var mount = document.createElement("div");
			mount.id = "docx-editor-mount-" + token;
			mount.className = "docx-editor-mount";
			host.appendChild(mount);
			var done = false;
			var finish = function(inError) {
				if (!done && token === self.openToken) {
					done = true;
					inCallback(inError || null);
				}
			};
			var file = new File([inBytes], "document.docx", {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
			try {
				self.superdoc = new window.SuperDoc({
					selector: "#" + mount.id,
					documentMode: "viewing",
					documents: [{id: "document", type: "docx", data: file}],
					telemetry: {enabled: false},
					// Its comments side panel would leave a phone's page a few words wide; comments in the
					// document are still kept in the copy.
					modules: {comments: false},
					zoom: {mode: "fit-width"},
					user: {name: $L("Me")},
					onReady: function() {
						finish(null);
					},
					onEditorUpdate: function() {
						if (self.editing) {
							self.doChanged();
						}
					},
					onContentError: function(inEvent) {
						finish(new Error($L("This Word document could not be opened for editing.") +
							(inEvent && inEvent.error && inEvent.error.message ? " " + inEvent.error.message : "")));
					},
					onException: function(inEvent) {
						enyo.warn("docx editor: " + (inEvent && inEvent.error && inEvent.error.message || inEvent));
					}
				});
			} catch (e) {
				finish(e);
				return;
			}
			if (self.superdoc.on) {
				self.superdoc.on("zoomChange", function() {
					self.doZoomChanged();
				});
			}
		});
	},

	isOpen: function() {
		return !!this.superdoc;
	},

	setEditing: function(inEditing) {
		this.editing = !!inEditing;
		this.$.bar.setShowing(this.editing);
		if (this.superdoc) {
			this.superdoc.setDocumentMode(this.editing ? "editing" : "viewing");
		}
		this.resized();
	},

	//* A toolbar key: its command on the document, where the cursor is.
	run: function(inSender) {
		var editor = this.superdoc && this.superdoc.activeEditor;
		var fn = editor && editor.commands && editor.commands[inSender.command];
		if (fn) {
			fn();
			if (editor.view && editor.view.focus) {
				editor.view.focus();
			}
		}
	},

	//* Zoom in or out (inDirection 1 or -1) to the next of inSteps (fractions, 1 = 100%).
	stepZoom: function(inDirection, inSteps) {
		if (!this.superdoc) {
			return;
		}
		var z = this.superdoc.getZoom() / 100, next = z;
		for (var i = 0; i < inSteps.length; i++) {
			var s = inDirection > 0 ? inSteps[i] : inSteps[inSteps.length - 1 - i];
			if (inDirection > 0 ? s > z + 0.001 : s < z - 0.001) {
				next = s;
				break;
			}
		}
		this.superdoc.setZoom(Math.round(next * 100));
	},

	fitWidth: function() {
		if (this.superdoc) {
			this.superdoc.setZoomMode("fit-width");
		}
	},

	//* The zoom in percent, or 0 when nothing is open.
	zoomPercent: function() {
		return this.superdoc ? Math.round(this.superdoc.getZoom()) : 0;
	},

	//* The edited document as the bytes of a .docx: inCallback(err, Uint8Array).
	copyBytes: function(inCallback) {
		var self = this;
		if (!this.superdoc) {
			inCallback(new Error($L("Nothing to save.")));
			return;
		}
		this.superdoc.export({triggerDownload: false}).then(function(inBlob) {
			return new Response(inBlob).arrayBuffer();
		}).then(function(inBuffer) {
			return DocDocxCopy.clean(window.JSZip, inBuffer, self.original);
		}).then(function(inBytes) {
			inCallback(null, inBytes);
		}, function(e) {
			inCallback(new Error($L("The document could not be written: ") + (e && e.message || e)));
		});
	},

	close: function() {
		this.openToken++;
		if (this.superdoc) {
			try {
				this.superdoc.destroy();
			} catch (e) {
				enyo.warn("docx editor: " + e);
			}
			this.superdoc = null;
		}
		var host = this.$.host.hasNode();
		if (host) {
			host.innerHTML = "";
		}
		this.editing = false;
	}
});
