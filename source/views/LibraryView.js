/**
	The documents on the device: the recently opened ones first, then every
	document under /media/internal, filtered by kind and by a search box.

	The list is one node of escaped HTML (a tap finds its row by index), so
	a device with a thousand documents lists as fast as one with ten.
*/
enyo.kind({
	name: "LibraryView",
	kind: enyo.VFlexBox,
	className: "library",
	events: {
		onOpenDocument: "",
		onRefresh: ""
	},
	filters: {all: null, text: "text", sheet: "sheet", slides: "slides"},
	components: [
		{name: "header", kind: enyo.HFlexBox, className: "library-header", align: "center", components: [
			{kind: enyo.Control, className: "library-appicon"},
			{name: "title", flex: 1, className: "library-title", content: $L("Documents")},
			{name: "search", kind: "SearchInput", className: "library-search", hint: $L("Search"), changeOnInput: true,
				autoCapitalize: "lowercase", autocorrect: false, spellcheck: false, onchange: "searchChanged", onCancel: "searchChanged"}
		]},
		{name: "scroller", kind: "Scroller", flex: 1, className: "library-scroller", components: [
			{name: "list", className: "library-list", onclick: "listClick"},
			{name: "empty", className: "library-empty", showing: false, allowHtml: true, content: ""}
		]},
		{name: "toolbar", kind: "Toolbar", className: "library-toolbar", components: [
			{name: "filterGroup", kind: "RadioToolButtonGroup", className: "library-filters", onChange: "filterChanged", components: [
				{caption: $L("All"), value: "all"},
				{caption: $L("Text"), value: "text"},
				{caption: $L("Sheets"), value: "sheet"},
				{caption: $L("Slides"), value: "slides"}
			]},
			{name: "refreshButton", kind: "ToolButton", className: "library-refresh", icon: "images/refresh.svg", onclick: "doRefresh"}
		]}
	],

	create: function() {
		this.inherited(arguments);
		this.docs = [];
		this.recent = [];
		this.rows = [];
		this.filter = "all";
		this.scanning = false;
		this.$.filterGroup.setValue("all");
	},

	setDocuments: function(inDocs) {
		this.docs = inDocs || [];
		this.refreshList();
	},

	setRecent: function(inRecent) {
		this.recent = inRecent || [];
		this.refreshList();
	},

	setScanning: function(inScanning, inFound) {
		this.scanning = inScanning;
		this.found = inFound || 0;
		this.$.refreshButton.addRemoveClass("library-refresh-busy", !!inScanning);
		this.refreshList();
	},

	filterChanged: function(inSender) {
		this.filter = inSender.getValue();
		this.refreshList();
	},

	searchChanged: function() {
		var self = this;
		enyo.job("documents-library-search", function() {
			self.refreshList();
		}, 150);
	},

	matches: function(inDoc, inQuery) {
		var family = FileTypes.familyOf(inDoc.kind);
		if (this.filters[this.filter] && family !== this.filters[this.filter]) {
			return false;
		}
		return !inQuery || DocFinder.foldString(inDoc.name).indexOf(inQuery) >= 0;
	},

 	//* The list is rebuilt when shown: rebuilding it behind the viewer would make the page lay out for nothing.
	refreshList: function() {
		if (this.hidden) {
			this.stale = true;
			return;
		}
		this.stale = false;
		var q = DocFinder.foldString(String(this.$.search.getValue() || "").replace(/^\s+|\s+$/g, ""));
		var rows = [];
		var html = [];
		var esc = enyo.string.escapeHtml;
		var self = this;
		var add = function(inDoc, inRecent) {
			rows.push(inDoc);
			var family = FileTypes.kinds[inDoc.kind] ? FileTypes.familyOf(inDoc.kind) : "text";
			var folder = String(inDoc.dir || DocLibrary.dirName(inDoc.path)).replace(/^\/media\/internal\/?/, "");
			var sub = inRecent && inDoc.opened ? self.when(inDoc.opened) : (folder || $L("USB storage"));
			html.push("<div class=\"library-item\" data-row=\"" + (rows.length - 1) + "\">" +
				"<div class=\"library-icon library-icon-" + (family || inDoc.kind) + "\"><span>" + esc(FileTypes.extensionOf(inDoc.name).slice(0, 4)) + "</span></div>" +
				"<div class=\"library-text\"><div class=\"library-name\">" + esc(inDoc.name) + "</div>" +
				"<div class=\"library-sub\">" + esc(sub) + "</div></div></div>");
		};
		var recent = this.recent.filter(function(d) {
			return self.matches(d, q);
		});
		if (recent.length) {
			html.push("<div class=\"library-divider\">" + esc($L("Recent")) + "</div>");
			recent.forEach(function(d) {
				add(d, true);
			});
		}
		var all = this.docs.filter(function(d) {
			return self.matches(d, q);
		});
		if (all.length) {
			html.push("<div class=\"library-divider\">" + esc($L("On this device")) +
				" <span class=\"library-count\">" + all.length + "</span></div>");
			all.forEach(function(d) {
				add(d, false);
			});
		}
		this.rows = rows;
		var node = this.$.list.hasNode();
		if (node) {
			node.innerHTML = html.join("");
		} else {
			this.$.list.setContent(html.join(""));
			this.$.list.allowHtml = true;
		}
		var empty = !all.length && !recent.length;
		this.$.empty.setShowing(empty);
		if (empty) {
			this.$.empty.setContent(this.scanning ? "<div class=\"library-scanning\">" + esc($L("Looking for documents…")) + "</div>" :
				q || this.filter !== "all" ? esc($L("No documents match.")) :
				"<p>" + esc($L("No documents found.")) + "</p><p class=\"library-hint\">" +
				esc($L("Copy Word, OpenDocument, Excel, PowerPoint, CSV, text or Markdown files to the device's USB storage, " +
					"or open them from Files or Email.")) + "</p>");
		}
		this.$.title.setContent(this.scanning ? esc($L("Documents")) + " <span class=\"library-busy\">" +
			esc($L("searching…")) + "</span>" : esc($L("Documents")));
		this.$.title.allowHtml = true;
		this.$.title.contentChanged();
	},

	//* Whether the library is on screen (the app tells it); while hidden the list is only marked stale.
	setHidden: function(inHidden) {
		this.hidden = !!inHidden;
		if (!this.hidden && this.stale) {
			this.refreshList();
		}
	},

	when: function(inTime) {
		var d = new Date(inTime);
		var now = new Date();
		var pad = function(n) {
			return (n < 10 ? "0" : "") + n;
		};
		if (d.toDateString() === now.toDateString()) {
			return $L("Today") + " " + pad(d.getHours()) + ":" + pad(d.getMinutes());
		}
		return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
	},

	listClick: function(inSender, inEvent) {
		var t = inEvent.target;
		while (t && t !== this.$.list.hasNode()) {
			if (t.getAttribute && t.getAttribute("data-row") !== null) {
				var doc = this.rows[+t.getAttribute("data-row")];
				if (doc) {
					this.doOpenDocument(doc.path);
				}
				return true;
			}
			t = t.parentNode;
		}
	}
});
