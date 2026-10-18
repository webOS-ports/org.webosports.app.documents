/**
	Edits a document in LibreOffice, through the office service's editing
	sessions (org.webosports.service.office: open, paintTile, postKey,
	postMouse, uno, extTextInput, setPart, save, close, and the events
	subscription). LibreOffice keeps the document; this view draws the tiles
	it paints, shows its cursor and selection, and sends taps, typing and the
	key bar's commands. Every format LibreOffice reads can be edited so,
	presentations a slide at a time.

		editor.open("/media/internal/Talk.pptx", function(err) {...});
		editor.copyBytes(function(err, bytes) {...});   // the edited document, as copyFormat()
		editor.close();

	Coordinates are LibreOffice's twips (1/1440 inch); the page is drawn at
	zoom * 96 / 1440 CSS pixels per twip, in tiles of TILE CSS pixels.

	Taps and scrolling are handled here, not by Enyo: the host keeps its
	touch events from the document, where Enyo's touch shim would cancel
	them (no native scrolling) and focus nodes itself. Typing goes through
	a hidden text box: a typed character or a committed composition is
	sent as text, a composition in progress (a keypad's T9, a predicting
	keyboard) as LibreOffice's preedit, and the editing keys as keys.
*/
enyo.kind({
	name: "LokEditor",
	kind: enyo.VFlexBox,
	className: "lok-editor",
	events: {
		onChanged: "",
		onClosed: ""
	},
	//* Tile side, in CSS pixels
	TILE: 256,
	//* Tile requests in flight at once
	PARALLEL: 3,
	SERVICE: "luna://org.webosports.service.office/",
	components: [
		{name: "bar", kind: enyo.HFlexBox, className: "docx-editor-bar lok-bar", align: "center", pack: "center", components: [
			{kind: "ToolButton", className: "viewer-tool", icon: "images/undo.svg", command: ".uno:Undo", onclick: "uno"},
			{kind: "ToolButton", className: "viewer-tool", icon: "images/redo.svg", command: ".uno:Redo", onclick: "uno"},
			{name: "bold", kind: "ToolButton", className: "viewer-tool docx-b", caption: "B", command: ".uno:Bold", onclick: "uno"},
			{name: "italic", kind: "ToolButton", className: "viewer-tool docx-i", caption: "I", command: ".uno:Italic", onclick: "uno"},
			{name: "underline", kind: "ToolButton", className: "viewer-tool docx-u", caption: "U", command: ".uno:Underline", onclick: "uno"},
			{kind: "ToolButton", className: "viewer-tool", icon: "images/list-bullets.svg", command: ".uno:DefaultBullet", onclick: "uno"},
			{name: "prevPart", kind: "ToolButton", className: "viewer-tool", icon: "images/prev.svg", onclick: "prevPart", showing: false},
			{name: "partLabel", className: "lok-part", showing: false},
			{name: "nextPart", kind: "ToolButton", className: "viewer-tool", icon: "images/next.svg", onclick: "nextPart", showing: false}
		]},
		{name: "host", flex: 1, className: "lok-host", requiresDomMousedown: true},
		// A long press: what to do with the selection, by LibreOffice's own commands
		{name: "editMenu", kind: "Menu", className: "lok-edit-menu", components: [
			{caption: $L("Cut"), command: ".uno:Cut", onclick: "menuUno"},
			{caption: $L("Copy"), command: ".uno:Copy", onclick: "menuUno"},
			{caption: $L("Paste"), command: ".uno:Paste", onclick: "menuUno"},
			{caption: $L("Select All"), command: ".uno:SelectAll", onclick: "menuUno"}
		]}
	],

	create: function() {
		this.inherited(arguments);
		this.session = null;
		this.zoom = 1;
		this.fit = true;
		this.tiles = {};
		this.queue = [];
		this.inFlight = 0;
		this.openToken = 0;
	},

	isOpen: function() {
		return !!this.session;
	},

	// --- The bus ------------------------------------------------------------

	//* One call; inCallback(err, result). Answers for an older document are dropped.
	call: function(inMethod, inParams, inCallback) {
		var token = this.openToken;
		var self = this;
		FileService.call(this.SERVICE + inMethod, inParams, function(inErr, inResult) {
			if (token === self.openToken && inCallback) {
				inCallback(inErr, inResult);
			}
		});
	},

	//* The events subscription, kept open until close()
	subscribe: function() {
		var self = this, token = this.openToken;
		if (typeof PalmServiceBridge === "undefined") {
			return;
		}
		var bridge = new PalmServiceBridge();
		bridge.onservicecallback = function(inText) {
			var msg;
			try {
				msg = JSON.parse(inText);
			} catch (e) {
				return;
			}
			if (token === self.openToken && msg.event) {
				self.event(msg);
			}
		};
		bridge.call(this.SERVICE + "events", JSON.stringify({session: this.session.id, subscribe: true}));
		this.eventBridge = bridge;
	},

	// --- Opening and closing -------------------------------------------------

	/**
		Opens inPath (a document under /media/internal) in a session and
		shows its first page or slide; inCallback(err) when shown or failed.
	*/
	open: function(inPath, inCallback) {
		var self = this;
		this.close();
		this.path = inPath;
		this.call("open", {source: inPath}, function(inErr, inResult) {
			if (inErr || !inResult || !inResult.session) {
				inCallback(inErr || new Error($L("LibreOffice could not open the document.")));
				return;
			}
			self.session = {id: inResult.session, kind: inResult.kind, parts: inResult.parts || 1, part: inResult.part || 0,
				width: inResult.width, height: inResult.height, tileMode: inResult.tileMode || "rgba"};
			self.subscribe();
			self.build();
			self.showPart();
			inCallback(null);
		});
	},

	close: function() {
		var s = this.session;
		this.openToken++;
		if (this.eventBridge) {
			try {
				this.eventBridge.cancel();
			} catch (e) {}
			this.eventBridge = null;
		}
		if (s) {
			FileService.call(this.SERVICE + "close", {session: s.id}, function() {});
		}
		this.session = null;
		this.tiles = {};
		this.queue = [];
		this.inFlight = 0;
		var host = this.$.host.hasNode();
		if (host) {
			host.innerHTML = "";
		}
		this.page = null;
	},

	// --- The page ------------------------------------------------------------

	//* The page, its cursor and selection layers, and the hidden text box
	build: function() {
		var host = this.$.host.hasNode();
		var self = this;
		host.innerHTML = "";
		var page = this.page = document.createElement("div");
		page.className = "lok-page";
		host.appendChild(page);
		this.selectionLayer = DocRender.el("div", "lok-selection", page);
		this.cursor = DocRender.el("div", "lok-cursor", page);
		this.cursor.style.display = "none";
		var input = this.input = document.createElement("textarea");
		input.className = "lok-input";
		input.setAttribute("autocapitalize", "off");
		input.setAttribute("autocorrect", "off");
		input.setAttribute("spellcheck", "false");
		page.appendChild(input);
		this.hookInput(input);
		if (!this.hostHooked) {
			this.hostHooked = true;
			this.hookTouch(host);
			host.addEventListener("scroll", function() {
				self.paintVisible();
			});
		}
	},

	//* CSS pixels per twip
	scale: function() {
		return this.zoom * 96 / 1440;
	},

	//* Lays the page out for the current part and zoom, and paints what shows
	showPart: function() {
		var s = this.session, host = this.$.host.hasNode();
		if (!s || !this.page) {
			return;
		}
		if (this.fit && host && host.clientWidth && s.width) {
			// The page's width fills the view, less a margin
			this.zoom = Math.max(0.1, (host.clientWidth - 16) * 1440 / 96 / s.width);
		}
		this.laidOutWidth = host ? host.clientWidth : 0;
		var sc = this.scale();
		this.page.style.width = Math.ceil(s.width * sc) + "px";
		this.page.style.height = Math.ceil(s.height * sc) + "px";
		this.clearTiles();
		var many = s.parts > 1 && (s.kind === "presentation" || s.kind === "drawing");
		this.$.prevPart.setShowing(many);
		this.$.nextPart.setShowing(many);
		this.$.partLabel.setShowing(many);
		this.$.partLabel.setContent($L("#{n} of #{all}").replace("#{n}", s.part + 1).replace("#{all}", s.parts));
		this.$.prevPart.setDisabled(s.part <= 0);
		this.$.nextPart.setDisabled(s.part >= s.parts - 1);
		this.paintVisible();
	},

	clearTiles: function() {
		for (var k in this.tiles) {
			if (this.tiles.hasOwnProperty(k) && this.tiles[k].canvas.parentNode) {
				this.tiles[k].canvas.parentNode.removeChild(this.tiles[k].canvas);
			}
		}
		this.tiles = {};
		this.queue = [];
	},

	//* Every tile in or near the view is drawn, or asked for
	paintVisible: function() {
		var s = this.session, host = this.$.host.hasNode();
		if (!s || !host || !this.page) {
			return;
		}
		var sc = this.scale(), T = this.TILE;
		var cols = Math.ceil(s.width * sc / T), rows = Math.ceil(s.height * sc / T);
		var c0 = Math.max(0, Math.floor(host.scrollLeft / T)), c1 = Math.min(cols - 1, Math.floor((host.scrollLeft + host.clientWidth) / T));
		// A row ahead and behind, for scrolling
		var r0 = Math.max(0, Math.floor(host.scrollTop / T) - 1), r1 = Math.min(rows - 1, Math.floor((host.scrollTop + host.clientHeight) / T) + 1);
		for (var r = r0; r <= r1; r++) {
			for (var c = c0; c <= c1; c++) {
				var key = s.part + ":" + c + ":" + r;
				var tile = this.tiles[key];
				if (!tile) {
					tile = this.tiles[key] = this.makeTile(c, r);
				}
				if (tile.dirty && !tile.queued) {
					tile.queued = true;
					this.queue.push(key);
				}
			}
		}
		this.pump();
	},

	makeTile: function(inCol, inRow) {
		var T = this.TILE;
		var canvas = document.createElement("canvas");
		var ratio = Math.min(4, this.pixelRatio());
		canvas.className = "lok-tile";
		canvas.width = Math.min(1024, Math.round(T * ratio));
		canvas.height = canvas.width;
		canvas.style.left = (inCol * T) + "px";
		canvas.style.top = (inRow * T) + "px";
		canvas.style.width = T + "px";
		canvas.style.height = T + "px";
		this.page.insertBefore(canvas, this.selectionLayer);
		return {col: inCol, row: inRow, canvas: canvas, dirty: true, queued: false, generation: 0};
	},

	/**
		Device pixels per CSS pixel. WAM zooms the whole page on LuneOS (450 CSS
		px wide on the Q25's 720 px screen) and still says devicePixelRatio 1:
		tiles drawn at that would be stretched and blurred.
	*/
	pixelRatio: function() {
		var ratio = window.devicePixelRatio || 1;
		if (window.PalmSystem && window.screen && window.innerWidth && screen.width > window.innerWidth) {
			ratio = Math.max(ratio, screen.width / window.innerWidth);
		}
		return ratio;
	},

	//* Asks for the next tiles, PARALLEL at a time
	pump: function() {
		while (this.inFlight < this.PARALLEL && this.queue.length) {
			var key = this.queue.shift();
			var tile = this.tiles[key];
			if (tile && tile.dirty) {
				this.requestTile(key, tile);
			}
		}
	},

	requestTile: function(inKey, inTile) {
		var s = this.session, self = this;
		var twips = this.TILE / this.scale();
		var generation = ++inTile.generation;
		var part = s.part;
		inTile.dirty = false;
		inTile.queued = false;
		this.inFlight++;
		var t0 = Date.now();
		this.call("paintTile", {session: s.id, part: part, x: Math.round(inTile.col * twips), y: Math.round(inTile.row * twips),
			width: Math.max(1, Math.round(twips)), height: Math.max(1, Math.round(twips)),
			pixelWidth: inTile.canvas.width, pixelHeight: inTile.canvas.height}, function(inErr, inResult) {
			if (inErr || !inResult || !inResult.file) {
				self.inFlight--;
				self.pump();
				return;
			}
			self.readTile(inResult.file, function(inBytes) {
				self.inFlight--;
				self.lastTileMs = Date.now() - t0;
				// A newer request for the same tile wins; one for another part is dropped
				if (inBytes && generation === inTile.generation && self.tiles[inKey] === inTile) {
					self.drawTile(inTile, inBytes, inResult.tileMode || s.tileMode);
				}
				self.pump();
			});
		});
	},

	readTile: function(inPath, inCallback) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", DocLibrary.pathToUrl(inPath), true);
		xhr.responseType = "arraybuffer";
		xhr.onload = function() {
			inCallback(xhr.response && xhr.response.byteLength ? xhr.response : null);
		};
		xhr.onerror = function() {
			inCallback(null);
		};
		xhr.send();
	},

	drawTile: function(inTile, inBytes, inMode) {
		var w = inTile.canvas.width, h = inTile.canvas.height;
		var data = new Uint8ClampedArray(inBytes);
		if (data.length !== w * h * 4) {
			return;
		}
		if (inMode === "bgra") {
			for (var i = 0; i < data.length; i += 4) {
				var b = data[i];
				data[i] = data[i + 2];
				data[i + 2] = b;
			}
		}
		var ctx = inTile.canvas.getContext("2d");
		ctx.putImageData(new ImageData(data, w, h), 0, 0);
	},

	//* Tiles under a twips rectangle of a part are painted again (all of the part when no rectangle)
	invalidate: function(inPart, inRect) {
		var s = this.session;
		var twips = this.TILE / this.scale();
		for (var k in this.tiles) {
			if (!this.tiles.hasOwnProperty(k)) {
				continue;
			}
			var t = this.tiles[k];
			if (+k.split(":")[0] !== inPart) {
				continue;
			}
			var x = t.col * twips, y = t.row * twips;
			if (!inRect || (inRect.x < x + twips && inRect.x + inRect.w > x && inRect.y < y + twips && inRect.y + inRect.h > y)) {
				t.dirty = true;
			}
		}
		if (s && inPart === s.part) {
			var self = this;
			enyo.job("lok-paint", function() {
				self.paintVisible();
			}, 30);
		}
	},

	// --- LibreOffice's callbacks --------------------------------------------

	//* "x, y, w, h" (twips) and more numbers after; null for "EMPTY"
	numbers: function(inText) {
		var parts = String(inText || "").split(",");
		var out = [];
		for (var i = 0; i < parts.length; i++) {
			var n = parseFloat(parts[i]);
			if (isNaN(n)) {
				return null;
			}
			out.push(n);
		}
		return out;
	},

	event: function(inMsg) {
		var s = this.session;
		var p = inMsg.payload || "";
		if (!s) {
			return;
		}
		switch (inMsg.event) {
		case "invalidate":
			if (/^EMPTY/.test(p)) {
				var m = /^EMPTY,\s*(-?\d+)/.exec(p);
				this.invalidate(m ? +m[1] : s.part, null);
			} else {
				var n = this.numbers(p);
				if (n && n.length >= 4) {
					this.invalidate(n.length >= 5 ? n[4] : s.part, {x: n[0], y: n[1], w: n[2], h: n[3]});
				}
			}
			// Not an edit by itself: a tap that selects a text box repaints too
			break;
		case "cursor":
			this.showCursor(p);
			break;
		case "cursorVisible":
			this.cursor.style.display = p === "true" && this.cursorRect ? "" : "none";
			break;
		case "selection":
			this.showSelection(p);
			break;
		case "state":
			this.showState(p);
			break;
		case "documentSize":
			var size = this.numbers(p);
			if (size && size.length >= 2) {
				s.width = size[0];
				s.height = size[1];
				this.showPart();
			}
			break;
		case "part":
			var part = parseInt(p, 10);
			if (!isNaN(part) && part !== s.part && part >= 0 && part < s.parts) {
				s.part = part;
				this.showPart();
			}
			break;
		case "closed":
			this.session = null;
			this.doClosed(p || inMsg.reason);
			break;
		}
	},

	showCursor: function(inPayload) {
		var n = this.numbers(inPayload);
		var sc = this.scale();
		if (!n || n.length < 4) {
			this.cursorRect = null;
			this.cursor.style.display = "none";
			return;
		}
		this.cursorRect = {x: n[0], y: n[1], w: n[2], h: n[3]};
		var st = this.cursor.style;
		st.display = "";
		st.left = Math.round(n[0] * sc) + "px";
		st.top = Math.round(n[1] * sc) + "px";
		st.height = Math.max(8, Math.round(n[3] * sc)) + "px";
		// The text box follows the cursor, so the keyboard's suggestions and a keypad's preedit sit there
		this.input.style.left = st.left;
		this.input.style.top = st.top;
	},

	showSelection: function(inPayload) {
		var layer = this.selectionLayer, sc = this.scale();
		layer.innerHTML = "";
		// In twips, for a long press: on the selection, or elsewhere
		this.selectionRects = [];
		var rects = String(inPayload || "").split(";");
		for (var i = 0; i < rects.length; i++) {
			var n = this.numbers(rects[i]);
			if (!n || n.length < 4) {
				continue;
			}
			this.selectionRects.push(n);
			var d = DocRender.el("div", "lok-selection-rect", layer);
			d.style.left = Math.round(n[0] * sc) + "px";
			d.style.top = Math.round(n[1] * sc) + "px";
			d.style.width = Math.max(1, Math.round(n[2] * sc)) + "px";
			d.style.height = Math.max(1, Math.round(n[3] * sc)) + "px";
		}
	},

	//* ".uno:Bold=true", or LibreOffice's JSON form {commandName, state}
	showState: function(inPayload) {
		var name = "", value = "";
		var text = String(inPayload || "");
		if (text.charAt(0) === "{") {
			try {
				var j = JSON.parse(text);
				name = j.commandName || "";
				value = String(j.state);
			} catch (e) {
				return;
			}
		} else {
			var at = text.indexOf("=");
			name = text.slice(0, at);
			value = text.slice(at + 1);
		}
		var keys = {".uno:Bold": "bold", ".uno:Italic": "italic", ".uno:Underline": "underline"};
		if (keys[name]) {
			this.$[keys[name]].addRemoveClass("viewer-tool-on", value === "true");
		}
	},

	// --- Touch -------------------------------------------------------------------

	//* Holding this long without moving is a long press: the edit menu
	LONG_PRESS_MS: 500,

	//* Taps go to LibreOffice; a long press opens the edit menu; a drag scrolls (natively: Enyo is kept out)
	hookTouch: function(inHost) {
		var self = this, start = null, lastTap = null, timer = null;
		var stopTimer = function() {
			if (timer) {
				clearTimeout(timer);
				timer = null;
			}
		};
		inHost.addEventListener("touchstart", function(e) {
			e.stopPropagation();
			stopTimer();
			var t = e.touches && e.touches[0];
			start = t && e.touches.length === 1 ? {x: t.clientX, y: t.clientY, at: Date.now()} : null;
			if (start) {
				var held = start;
				timer = setTimeout(function() {
					timer = null;
					held.long = true;
					self.longPress(held.x, held.y);
				}, self.LONG_PRESS_MS);
			}
		});
		inHost.addEventListener("touchmove", function(e) {
			e.stopPropagation();
			var t = e.touches && e.touches[0];
			if (start && t && (Math.abs(t.clientX - start.x) > 8 || Math.abs(t.clientY - start.y) > 8)) {
				stopTimer();
				start = null;
			}
		});
		inHost.addEventListener("touchend", function(e) {
			e.stopPropagation();
			stopTimer();
			var t = e.changedTouches && e.changedTouches[0];
			if (!start || !t) {
				return;
			}
			// The browser would make a mouse click of it too: this is the tap, or the end of a long press
			e.preventDefault();
			if (start.long) {
				start = null;
				return;
			}
			var now = Date.now();
			var count = lastTap && now - lastTap.at < 350 && Math.abs(t.clientX - lastTap.x) < 20 &&
				Math.abs(t.clientY - lastTap.y) < 20 ? 2 : 1;
			lastTap = {x: t.clientX, y: t.clientY, at: now};
			start = null;
			self.tap(t.clientX, t.clientY, count);
		});
		// A desktop browser (the tests) and a mouse: a click is a tap, the other button a long press
		inHost.addEventListener("click", function(e) {
			if (e.target === self.input) {
				return;
			}
			self.tap(e.clientX, e.clientY, e.detail >= 2 ? 2 : 1);
		});
		inHost.addEventListener("contextmenu", function(e) {
			e.preventDefault();
			self.longPress(e.clientX, e.clientY);
		});
	},

	/*
	 * A long press at a point of the screen. On the selection it is kept;
	 * anywhere else the word there is selected: a double click, which
	 * LibreOffice takes only as a click and then a second one (a click
	 * counted 2 on its own selects nothing). Then the edit menu opens there.
	 */
	longPress: function(inX, inY) {
		var s = this.session;
		if (!s || !this.page) {
			return;
		}
		var r = this.page.getBoundingClientRect();
		var sc = this.scale();
		var x = Math.round((inX - r.left) / sc), y = Math.round((inY - r.top) / sc);
		if (x < 0 || y < 0 || x > s.width || y > s.height) {
			return;
		}
		var onSelection = (this.selectionRects || []).some(function(n) {
			return x >= n[0] && x <= n[0] + n[2] && y >= n[1] && y <= n[1] + n[3];
		});
		if (!onSelection) {
			this.tap(inX, inY, 1);
			this.tap(inX, inY, 2);
		}
		this.$.editMenu.openAt({left: inX, top: inY}, true);
	},

	menuUno: function(inSender) {
		var s = this.session;
		if (!s) {
			return;
		}
		this.call("uno", {session: s.id, command: inSender.command});
		if (inSender.command !== ".uno:Copy" && inSender.command !== ".uno:SelectAll") {
			this.doChanged();
		}
		this.focusInput();
	},

	//* A tap at a point of the screen: a click in the document there, and the keyboard
	tap: function(inX, inY, inCount) {
		var s = this.session;
		if (!s || !this.page) {
			return;
		}
		var r = this.page.getBoundingClientRect();
		var sc = this.scale();
		var x = Math.round((inX - r.left) / sc), y = Math.round((inY - r.top) / sc);
		if (x < 0 || y < 0 || x > s.width || y > s.height) {
			return;
		}
		this.call("postMouse", {session: s.id, type: "down", x: x, y: y, count: inCount, buttons: 1});
		this.call("postMouse", {session: s.id, type: "up", x: x, y: y, count: inCount, buttons: 1});
		this.focusInput();
	},

	focusInput: function() {
		if (this.input && document.activeElement !== this.input) {
			this.input.focus();
		}
	},

	// --- Typing -------------------------------------------------------------------

	//* LibreOffice's key codes (com.sun.star.awt.Key) for the keys that are not text
	keyCodes: {8: 1283, 9: 1282, 13: 1280, 27: 1281, 33: 1030, 34: 1031, 35: 1029, 36: 1028,
		37: 1026, 38: 1025, 39: 1027, 40: 1024, 46: 1286},

	hookInput: function(inInput) {
		var self = this, composing = false;
		inInput.addEventListener("keydown", function(e) {
			var code = self.keyCodes[e.keyCode];
			if (!code || composing || e.isComposing) {
				return;
			}
			e.preventDefault();
			// Shift (4096) with the arrows selects
			self.key(code + (e.shiftKey ? 4096 : 0));
		});
		inInput.addEventListener("compositionstart", function() {
			composing = true;
		});
		inInput.addEventListener("compositionupdate", function(e) {
			self.text("preedit", e.data || "");
		});
		inInput.addEventListener("compositionend", function(e) {
			composing = false;
			if (e.data) {
				self.text("commit", e.data);
			}
			inInput.value = "";
		});
		inInput.addEventListener("input", function(e) {
			if (composing || e.isComposing) {
				return;
			}
			if (e.inputType === "insertLineBreak" || e.inputType === "insertParagraph") {
				self.key(1280);
			} else if (e.inputType === "deleteContentBackward") {
				self.key(1283);
			} else if (e.data) {
				self.text("commit", e.data);
			}
			inInput.value = "";
		});
	},

	key: function(inKeyCode) {
		var s = this.session;
		if (!s) {
			return;
		}
		this.call("postKey", {session: s.id, type: "input", charCode: 0, keyCode: inKeyCode});
		this.call("postKey", {session: s.id, type: "up", charCode: 0, keyCode: inKeyCode});
		this.doChanged();
	},

	text: function(inType, inText) {
		var s = this.session;
		if (!s) {
			return;
		}
		this.call("extTextInput", {session: s.id, type: inType, text: inText});
		if (inType === "commit") {
			this.doChanged();
		}
	},

	// --- The key bar ------------------------------------------------------------

	uno: function(inSender) {
		var s = this.session;
		if (s) {
			this.call("uno", {session: s.id, command: inSender.command});
			this.doChanged();
			this.focusInput();
		}
	},

	setPart: function(inPart) {
		var s = this.session, self = this;
		if (!s || inPart < 0 || inPart >= s.parts) {
			return;
		}
		this.call("setPart", {session: s.id, part: inPart}, function(inErr, inResult) {
			if (!inErr && inResult && self.session) {
				self.session.part = inResult.part;
				self.session.width = inResult.width || self.session.width;
				self.session.height = inResult.height || self.session.height;
				self.$.host.hasNode().scrollTop = 0;
				self.showPart();
			}
		});
	},

	prevPart: function() {
		this.setPart(this.session ? this.session.part - 1 : -1);
	},

	nextPart: function() {
		this.setPart(this.session ? this.session.part + 1 : -1);
	},

	//* Zoom by inFactor (the viewer's zoom keys); 0 fits the width again
	zoomBy: function(inFactor) {
		if (!inFactor) {
			this.fit = true;
		} else {
			this.fit = false;
			this.zoom = Math.max(0.2, Math.min(4, this.zoom * inFactor));
		}
		this.showPart();
	},

	zoomPercent: function() {
		return Math.round(this.zoom * 100);
	},

	// --- Saving ------------------------------------------------------------------

	//* The format a copy is saved in: the document's own where LibreOffice writes it, else its family's
	copyFormat: function() {
		var ext = FileTypes.extensionOf(this.path || "");
		var same = {odt: 1, docx: 1, ods: 1, xlsx: 1, odp: 1, pptx: 1};
		if (same[ext]) {
			return ext;
		}
		var kind = this.session ? this.session.kind : "";
		return kind === "spreadsheet" ? "xlsx" : kind === "presentation" ? "pptx" : kind === "drawing" ? "odp" : "docx";
	},

	//* The edited document's bytes: inCallback(err, Uint8Array)
	copyBytes: function(inCallback) {
		var s = this.session;
		if (!s) {
			inCallback(new Error($L("Nothing to save.")));
			return;
		}
		this.call("save", {session: s.id, format: this.copyFormat()}, function(inErr, inResult) {
			if (inErr || !inResult || !inResult.target) {
				inCallback(inErr || new Error($L("LibreOffice could not save the document.")));
				return;
			}
			DocLoader.read(inResult.target, function(inErr2, inBuffer) {
				inCallback(inErr2, inBuffer ? new Uint8Array(inBuffer) : null);
			});
		});
	},

	resizeHandler: function() {
		this.inherited(arguments);
		var host = this.$.host.hasNode();
		// The Save button showing up resizes too: only a new width lays the page out again
		if (this.session && this.fit && host && host.clientWidth !== this.laidOutWidth) {
			this.showPart();
		}
	}
});
