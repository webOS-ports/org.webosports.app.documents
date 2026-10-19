/**
	Documents in the cloud: first the accounts with documents (set up in
	Accounts), then the folders of one account. A tap on a folder opens
	it, a tap on a document asks the app to download and open it. Back
	(the toolbar's arrow or the back gesture) goes up a folder, then to
	the accounts, then to the library.

	Lists folders and the documents the library would list, filtered by
	kind as the library is. The rows are one node of escaped HTML, as in
	the library.
*/
enyo.kind({
	name: "CloudView",
	kind: enyo.VFlexBox,
	className: "cloud",
	events: {
		onOpenCloudDocument: "",
		onLibrary: ""
	},
	filters: {all: null, text: "text", sheet: "sheet", slides: "slides"},
	components: [
		{name: "header", kind: enyo.HFlexBox, className: "library-header cloud-header", align: "center", components: [
			{kind: enyo.Control, className: "cloud-appicon"},
			{name: "titleBox", flex: 1, className: "cloud-titlebox", components: [
				{name: "title", className: "library-title", content: $L("Cloud")},
				{name: "subtitle", className: "cloud-subtitle", content: ""}
			]}
		]},
		{name: "body", flex: 1, className: "cloud-body", components: [
			{name: "scroller", kind: "Scroller", className: "cloud-scroller", components: [
				{name: "list", className: "library-list", onclick: "listClick"},
				{name: "status", className: "library-empty cloud-status", showing: false, components: [
					{name: "statusText", allowHtml: true, content: ""},
					{name: "retryButton", kind: "Button", className: "cloud-retry", caption: $L("Try Again"), showing: false, onclick: "reload"}
				]}
			]},
			// While a document is downloaded: over the list, which stays where it was
			{name: "busy", className: "cloud-busy", showing: false, components: [
				{name: "busySpinner", kind: "SpinnerLarge"},
				{name: "busyText", className: "cloud-busy-text", content: ""}
			]}
		]},
		{name: "toolbar", kind: "Toolbar", className: "library-toolbar cloud-toolbar", components: [
			{name: "backButton", kind: "ToolButton", className: "cloud-back", icon: "images/back.svg", onclick: "backClick"},
			{name: "filterGroup", kind: "RadioToolButtonGroup", className: "library-filters", onChange: "filterChanged", components: [
				{caption: $L("All"), value: "all"},
				{caption: $L("Text"), value: "text"},
				{caption: $L("Sheets"), value: "sheet"},
				{caption: $L("Slides"), value: "slides"}
			]},
			{name: "refreshButton", kind: "ToolButton", className: "library-refresh", icon: "images/refresh.svg", onclick: "reload"}
		]}
	],

	create: function() {
		this.inherited(arguments);
		this.accounts = null;
		this.account = null;
		this.stack = [];
		this.entries = [];
		this.rows = [];
		this.filter = "all";
		this.token = 0;
		this.loading = false;
		this.$.filterGroup.setValue("all");
	},

	//* Shown by the app: the accounts are asked for the first time, after that it is where it was.
	start: function() {
		if (!this.accounts && !this.loading) {
			this.loadAccounts();
		}
	},

	reload: function() {
		if (this.account) {
			this.loadFolder();
		} else {
			this.loadAccounts();
		}
	},

	loadAccounts: function() {
		var self = this;
		var token = ++this.token;
		this.account = null;
		this.stack = [];
		this.showLoading($L("Looking for cloud accounts…"));
		CloudService.accounts(function(inErr, inAccounts) {
			if (token !== self.token) {
				return;
			}
			self.loading = false;
			if (inErr) {
				self.accounts = null;
				self.showError(CloudService.describeError(inErr));
				return;
			}
			self.accounts = inAccounts;
			self.refreshList();
		});
	},

	openAccount: function(inAccount) {
		this.account = inAccount;
		this.stack = [{path: null, name: inAccount.provider}];
		this.loadFolder();
	},

	openFolder: function(inEntry) {
		this.stack.push({path: inEntry.path, name: inEntry.name});
		this.loadFolder();
	},

	loadFolder: function() {
		var self = this;
		var token = ++this.token;
		var account = this.account;
		this.entries = [];
		this.showLoading($L("Opening #{folder}…").replace("#{folder}", this.folderName()));
		CloudService.list(account, this.folderPath(), function(inErr, inEntries) {
			if (token !== self.token) {
				return;
			}
			self.loading = false;
			if (inErr) {
				self.showError(CloudService.describeError(inErr));
				return;
			}
			self.entries = inEntries;
			self.refreshList();
		});
	},

	/**
		One step back: up a folder, or from an account's top folder to the
		accounts. False on the accounts (the app then shows the library).
	*/
	back: function() {
		if (!this.account) {
			return false;
		}
		if (this.stack.length > 1) {
			this.stack.pop();
			this.loadFolder();
		} else {
			// The accounts were listed already: no need to ask again.
			this.token++;
			this.account = null;
			this.stack = [];
			this.entries = [];
			this.loading = false;
			if (this.accounts) {
				this.refreshList();
			} else {
				this.loadAccounts();
			}
		}
		return true;
	},

	backClick: function() {
		if (!this.back()) {
			this.doLibrary();
		}
	},

	//* The connector's locator of the folder on screen; null for an account's top folder.
	folderPath: function() {
		var top = this.stack[this.stack.length - 1];
		return top ? top.path : null;
	},

	folderName: function() {
		var top = this.stack[this.stack.length - 1];
		return top ? top.name : "";
	},

	//* Where a document of this folder is from, said in a line: "Dropbox · Work › Reports".
	place: function() {
		var names = this.stack.slice(1).map(function(s) {
			return s.name;
		});
		return this.account.provider + (names.length ? " · " + names.join(" › ") : "");
	},

	//* While a document is downloaded (inText), or not (null).
	setBusy: function(inText) {
		this.$.busy.setShowing(!!inText);
		this.$.busySpinner.setShowing(!!inText);
		this.$.busyText.setContent(enyo.string.escapeHtml(inText || ""));
	},

	isBusy: function() {
		return this.$.busy.showing;
	},

	filterChanged: function(inSender) {
		this.filter = inSender.getValue();
		this.refreshList();
	},

	showLoading: function(inText) {
		this.loading = true;
		this.rows = [];
		this.setListHtml("");
		this.showStatus("<div class=\"library-scanning\">" + enyo.string.escapeHtml(inText) + "</div>", false);
		this.updateHeader();
	},

	showError: function(inText) {
		this.rows = [];
		this.setListHtml("");
		this.showStatus(enyo.string.escapeHtml(inText), true);
		this.$.status.addClass("cloud-error");
		this.updateHeader();
	},

	showStatus: function(inHtml, inRetry) {
		this.$.status.removeClass("cloud-error");
		this.$.status.setShowing(!!inHtml);
		this.$.statusText.setContent(inHtml || "");
		this.$.retryButton.setShowing(!!inRetry);
	},

	updateHeader: function() {
		var esc = enyo.string.escapeHtml;
		if (!this.account) {
			this.$.title.setContent(esc($L("Cloud")));
			this.$.subtitle.setContent(esc($L("Documents in your online accounts")));
		} else {
			this.$.title.setContent(esc(this.folderName()));
			this.$.subtitle.setContent(esc(this.stack.length > 1 ? this.place() : (this.account.name || this.account.provider)));
		}
	},

	setListHtml: function(inHtml) {
		var node = this.$.list.hasNode();
		if (node) {
			node.innerHTML = inHtml;
		} else {
			this.$.list.setContent(inHtml);
			this.$.list.allowHtml = true;
		}
	},

	refreshList: function() {
		if (this.loading) {
			return;
		}
		var esc = enyo.string.escapeHtml;
		var self = this;
		var rows = [];
		var html = [];
		var add = function(inRow, inIcon, inLabel, inName, inSub, inPicture) {
			rows.push(inRow);
			// A picture's URL sits in style="...": single quotes, and nothing in the path may end it.
			var style = inPicture ? " style=\"background-image: url('file://" + encodeURI(inPicture)
				.replace(/'/g, "%27").replace(/\(/g, "%28").replace(/\)/g, "%29").replace(/"/g, "%22") + "')\"" : "";
			html.push("<div class=\"library-item cloud-item\" data-row=\"" + (rows.length - 1) + "\">" +
				"<div class=\"library-icon " + inIcon + "\"" + style + "><span>" + esc(inLabel) + "</span></div>" +
				"<div class=\"library-text\"><div class=\"library-name\">" + esc(inName) + "</div>" +
				"<div class=\"library-sub\">" + esc(inSub) + "</div></div></div>");
		};
		var empty = "";
		if (!this.account) {
			(this.accounts || []).forEach(function(a) {
				// The provider's own picture, as in Accounts; the cloud where the template has none.
				add({account: a}, a.icon ? "cloud-icon-provider" : "cloud-icon-account", "", a.provider, a.name, a.icon);
			});
			empty = "<p>" + esc($L("No cloud accounts with documents.")) + "</p><p class=\"library-hint\">" +
				esc($L("Add a Dropbox, Box, Google Drive, OneDrive or other online storage account in Accounts, " +
					"and its documents are listed here.")) + "</p>";
		} else {
			var wanted = this.filters[this.filter];
			CloudService.visibleEntries(this.entries).forEach(function(e) {
				if (e.type === "folder") {
					add({folder: e}, "cloud-icon-folder", "", e.name, $L("Folder"));
					return;
				}
				var kind = FileTypes.kindFromName(e.name);
				var family = FileTypes.familyOf(kind);
				if (wanted && family !== wanted) {
					return;
				}
				var sub = [DocLibrary.formatSize(e.size), self.when(e.modified)].filter(function(s) {
					return !!s;
				}).join(" · ");
				add({file: e}, "library-icon-" + (family || kind), FileTypes.extensionOf(e.name).slice(0, 4), e.name, sub);
			});
			empty = esc(this.entries.length ? $L("No documents match.") : $L("This folder is empty."));
		}
		this.rows = rows;
		this.setListHtml(html.join(""));
		this.showStatus(rows.length ? "" : empty, false);
		this.updateHeader();
	},

	//* A modification time ("2026-05-01T10:00:00Z" or ms) as a date.
	when: function(inTime) {
		if (inTime === undefined || inTime === null || inTime === "") {
			return "";
		}
		var d = new Date(typeof inTime === "number" || /^\d+$/.test(inTime) ? +inTime : inTime);
		if (isNaN(d.getTime())) {
			return "";
		}
		var pad = function(n) {
			return (n < 10 ? "0" : "") + n;
		};
		return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
	},

	listClick: function(inSender, inEvent) {
		var t = inEvent.target;
		while (t && t !== this.$.list.hasNode()) {
			if (t.getAttribute && t.getAttribute("data-row") !== null) {
				this.rowTapped(this.rows[+t.getAttribute("data-row")]);
				return true;
			}
			t = t.parentNode;
		}
	},

	rowTapped: function(inRow) {
		if (!inRow || this.isBusy()) {
			return;
		}
		if (inRow.account) {
			this.openAccount(inRow.account);
		} else if (inRow.folder) {
			this.openFolder(inRow.folder);
		} else if (inRow.file) {
			this.doOpenCloudDocument({
				account: this.account,
				entry: inRow.file,
				folderPath: this.folderPath(),
				folderName: this.folderName(),
				place: this.place(),
				// The folder's names, for a copy's name that is not taken
				names: this.entries.map(function(e) {
					return e.name;
				})
			});
		}
	}
});
