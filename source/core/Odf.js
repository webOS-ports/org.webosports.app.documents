/**
	OpenDocument text (.odt) and presentation (.odp) to a plain document
	model that source/views/DocRender.js turns into DOM. The model is plain
	data (no DOM, no functions), so it can be built in a Web Worker and
	posted to the page, and checked in node by the unit tests.

		var doc = DocOdf.text({"content.xml": "...", "styles.xml": "...", "meta.xml": "..."});
		doc.blocks   // [{t: "p", h: 0, s: {css}, r: [inline...]}, {t: "table", ...}, ...]
		doc.page     // {width, height, marginTop, ...} CSS lengths of the first page style
		doc.images   // zip paths of the pictures the blocks use
		doc.meta     // {title, author, ...}

		var track = {};
		var doc = DocOdf.text(files, {track: track});   // editable: paragraphs get an id
		track.paras[id]   // each paragraph's element (content.xml parsed with offsets, see Xml.js)
		track.content, track.styles

		var deck = DocOdf.presentation(files);
		deck.slides  // [{name, width, height (px), bg, items: [{x, y, w, h, kind, ...}]}]

	Blocks:
		{t: "p", h: <heading level or 0>, s, r: [inline], id, pt}
		                                     id and pt (its text as tokens() gives it) when tracked
		{t: "list", o: <ordered>, type: "decimal"|"lower-alpha"|..., start, items: [[block]]}
		{t: "table", s, cols: [css width], rows: [{header, s, cells: [{cs, rs, s, b: [block]}]}]}
		{t: "frame", s, b: [block]}          a text box anchored in the text
		{t: "pagebreak"}
		{t: "notes", items: [{label, b: [block]}]}  footnotes and endnotes, at the end
	Inline:
		{t: "text", v, s}   {t: "br"}   {t: "tab"}   {t: "a", href, r: [inline]}
		{t: "img", src, w, h, alt}   {t: "note", label}   {t: "frame", s, b: [block]}

	Style values are whitelisted CSS (see css()); nothing from the document
	reaches the page as markup or as script.
*/
(function(root) {
	var X = root.DocXml || (typeof require === "function" ? require("./Xml.js") : null);

	// --- CSS from ODF properties -----------------------------------------

	// Values may hold lengths, colours, names and simple lists; anything else
	// (url(), expressions, escapes) is dropped.
	var safeValue = /^[-#%.,()'" a-zA-Z0-9]*$/;

	function put(inCss, inProp, inValue) {
		if (inValue === undefined || inValue === null || inValue === "") {
			return;
		}
		var v = String(inValue);
		if (!safeValue.test(v) || /url\s*\(|expression/i.test(v)) {
			return;
		}
		inCss[inProp] = v;
	}

	var alignMap = {start: "left", end: "right", left: "left", right: "right", center: "center", justify: "justify"};

	function textProps(inAttrs, inCss, inFonts) {
		var a = inAttrs;
		put(inCss, "fontWeight", a["fo:font-weight"]);
		put(inCss, "fontStyle", a["fo:font-style"]);
		put(inCss, "fontSize", a["fo:font-size"]);
		put(inCss, "color", a["fo:color"]);
		if (a["fo:background-color"] && a["fo:background-color"] !== "transparent") {
			put(inCss, "backgroundColor", a["fo:background-color"]);
		}
		var font = a["style:font-name"] || a["fo:font-family"];
		if (font) {
			var family = (inFonts && inFonts[font]) || font;
			put(inCss, "fontFamily", "'" + family.replace(/['"]/g, "") + "', sans-serif");
		}
		var deco = [];
		if (a["style:text-underline-style"] && a["style:text-underline-style"] !== "none") {
			deco.push("underline");
		}
		if (a["style:text-line-through-style"] && a["style:text-line-through-style"] !== "none") {
			deco.push("line-through");
		}
		if (deco.length) {
			inCss.textDecoration = deco.join(" ");
		} else if (a["style:text-underline-style"] === "none" && inCss.textDecoration) {
			delete inCss.textDecoration;
		}
		var pos = a["style:text-position"];
		if (pos) {
			if (/^super|^[1-9]/.test(pos)) {
				inCss.verticalAlign = "super";
				inCss.fontSize = "0.7em";
			} else if (/^sub|^-/.test(pos)) {
				inCss.verticalAlign = "sub";
				inCss.fontSize = "0.7em";
			} else if (/^0/.test(pos)) {
				delete inCss.verticalAlign;
			}
		}
		put(inCss, "textTransform", a["fo:text-transform"]);
		if (a["fo:font-variant"] === "small-caps") {
			inCss.fontVariant = "small-caps";
		}
	}

	function boxProps(inAttrs, inCss) {
		var a = inAttrs;
		var sides = ["top", "right", "bottom", "left"];
		put(inCss, "margin", a["fo:margin"]);
		put(inCss, "padding", a["fo:padding"]);
		put(inCss, "border", a["fo:border"]);
		for (var i = 0; i < sides.length; i++) {
			var s = sides[i], S = s.charAt(0).toUpperCase() + s.slice(1);
			put(inCss, "margin" + S, a["fo:margin-" + s]);
			put(inCss, "padding" + S, a["fo:padding-" + s]);
			put(inCss, "border" + S, a["fo:border-" + s]);
		}
		if (a["fo:background-color"] && a["fo:background-color"] !== "transparent") {
			put(inCss, "backgroundColor", a["fo:background-color"]);
		}
	}

	function paraProps(inAttrs, inCss) {
		var a = inAttrs;
		if (a["fo:text-align"]) {
			put(inCss, "textAlign", alignMap[a["fo:text-align"]] || "left");
		}
		put(inCss, "textIndent", a["fo:text-indent"]);
		var lh = a["fo:line-height"];
		if (lh && lh !== "normal") {
			// A percentage of the font's line: as a plain factor, so it scales.
			put(inCss, "lineHeight", /%$/.test(lh) ? String(parseFloat(lh) / 100 * 1.15) : lh);
		}
		boxProps(a, inCss);
	}

	function cellProps(inAttrs, inCss) {
		boxProps(inAttrs, inCss);
		var va = inAttrs["style:vertical-align"];
		if (va === "middle" || va === "top" || va === "bottom") {
			inCss.verticalAlign = va;
		}
	}

	function merge(inA, inB) {
		var out = {};
		var k;
		for (k in inA) {
			if (inA.hasOwnProperty(k)) {
				out[k] = inA[k];
			}
		}
		for (k in inB) {
			if (inB.hasOwnProperty(k)) {
				out[k] = inB[k];
			}
		}
		return out;
	}

	function isEmpty(inObj) {
		for (var k in inObj) {
			if (inObj.hasOwnProperty(k)) {
				return false;
			}
		}
		return true;
	}

	// --- Lengths ---------------------------------------------------------

	var perUnit = {px: 1, pt: 96 / 72, pc: 16, "in": 96, cm: 96 / 2.54, mm: 96 / 25.4};

	//* An ODF length ("2.5cm", "12pt", "1in") in CSS pixels (96 per inch); NaN when not a length.
	function toPx(inLength) {
		var m = /^\s*(-?[0-9.]+)\s*(px|pt|pc|in|cm|mm)?\s*$/.exec(String(inLength || ""));
		if (!m) {
			return NaN;
		}
		return parseFloat(m[1]) * (perUnit[m[2] || "px"]);
	}

	// --- Styles ----------------------------------------------------------

	/**
		Collects every named style of the documents' style sections:
		{"paragraph:P1": {parent, family, text: {css}, para: {css}, ...}},
		plus the list styles, the font faces and the defaults per family.
	*/
	function Styles() {
		this.byName = {};
		this.defaults = {};
		this.lists = {};
		this.fonts = {};
		this.pageLayouts = {};
		this.masterPages = {};
	}

	Styles.prototype.addFonts = function(inDecls) {
		var faces = X.children(inDecls, "style:font-face");
		for (var i = 0; i < faces.length; i++) {
			var family = faces[i].a["svg:font-family"];
			if (family) {
				this.fonts[faces[i].a["style:name"]] = family.replace(/^['"]|['"]$/g, "");
			}
		}
	};

	Styles.prototype.read = function(inSection) {
		if (!inSection) {
			return;
		}
		for (var i = 0; i < inSection.c.length; i++) {
			var el = inSection.c[i];
			if (typeof el !== "object") {
				continue;
			}
			if (el.n === "style:style" || el.n === "style:default-style") {
				var st = this.parseStyle(el);
				if (el.n === "style:default-style") {
					this.defaults[st.family] = st;
				} else {
					this.byName[st.family + ":" + el.a["style:name"]] = st;
				}
			} else if (el.n === "text:list-style") {
				this.lists[el.a["style:name"]] = this.parseListStyle(el);
			} else if (el.n === "style:page-layout") {
				var pl = X.child(el, "style:page-layout-properties");
				this.pageLayouts[el.a["style:name"]] = pl ? pl.a : {};
			} else if (el.n === "style:master-page") {
				this.masterPages[el.a["style:name"]] = el;
			}
		}
	};

	Styles.prototype.parseStyle = function(inEl) {
		var st = {family: inEl.a["style:family"], parent: inEl.a["style:parent-style-name"],
			list: inEl.a["style:list-style-name"], text: {}, para: {}, cell: {}, col: {}, table: {}, graphic: {},
			breakBefore: false, raw: {}};
		for (var i = 0; i < inEl.c.length; i++) {
			var p = inEl.c[i];
			if (typeof p !== "object") {
				continue;
			}
			if (p.n === "style:text-properties") {
				textProps(p.a, st.text, this.fonts);
			} else if (p.n === "style:paragraph-properties") {
				paraProps(p.a, st.para);
				st.breakBefore = p.a["fo:break-before"] === "page";
			} else if (p.n === "style:table-cell-properties") {
				cellProps(p.a, st.cell);
			} else if (p.n === "style:table-column-properties") {
				st.col.width = p.a["style:column-width"];
			} else if (p.n === "style:table-properties") {
				put(st.table, "width", p.a["style:width"]);
				if (p.a["table:align"] === "center") {
					st.table.marginLeft = st.table.marginRight = "auto";
				}
				if (p.a["fo:background-color"] && p.a["fo:background-color"] !== "transparent") {
					put(st.table, "backgroundColor", p.a["fo:background-color"]);
				}
			} else if (p.n === "style:graphic-properties" || p.n === "style:drawing-page-properties") {
				st.raw = merge(st.raw, p.a);
			}
		}
		return st;
	};

	Styles.prototype.parseListStyle = function(inEl) {
		var levels = {};
		for (var i = 0; i < inEl.c.length; i++) {
			var l = inEl.c[i];
			if (typeof l !== "object") {
				continue;
			}
			var level = +(l.a["text:level"] || 1);
			if (l.n === "text:list-level-style-number") {
				levels[level] = {ordered: true, format: l.a["style:num-format"] || "1", start: +(l.a["text:start-value"] || 1)};
			} else if (l.n === "text:list-level-style-bullet" || l.n === "text:list-level-style-image") {
				levels[level] = {ordered: false};
			}
		}
		return levels;
	};

	/**
		The effective properties of a style, its parents and its family's
		default merged in that order; inKind is "text", "para", "cell", ...
	*/
	Styles.prototype.resolve = function(inFamily, inName, inKind) {
		var chain = [];
		var name = inName, guard = 0;
		while (name && guard++ < 32) {
			var st = this.byName[inFamily + ":" + name];
			if (!st) {
				break;
			}
			chain.unshift(st);
			name = st.parent;
		}
		var out = {};
		if (this.defaults[inFamily]) {
			out = merge(out, this.defaults[inFamily][inKind]);
		}
		for (var i = 0; i < chain.length; i++) {
			out = merge(out, chain[i][inKind]);
		}
		return out;
	};

	Styles.prototype.get = function(inFamily, inName) {
		return this.byName[inFamily + ":" + inName] || null;
	};

	//* The list style a paragraph style points at, following parents.
	Styles.prototype.listOfParagraph = function(inName) {
		var name = inName, guard = 0;
		while (name && guard++ < 32) {
			var st = this.byName["paragraph:" + name];
			if (!st) {
				return null;
			}
			if (st.list) {
				return st.list;
			}
			name = st.parent;
		}
		return null;
	};

	/**
		Properties of a drawing shape: the graphic defaults, then its
		presentation style's chain, then its own graphic style's chain.
	*/
	Styles.prototype.layered = function(inPresentationName, inGraphicName, inKind) {
		var out = this.defaults.graphic ? merge({}, this.defaults.graphic[inKind]) : {};
		var self = this;
		var chain = function(inFamily, inName) {
			var list = [], name = inName, guard = 0;
			while (name && guard++ < 32) {
				var st = self.byName[inFamily + ":" + name];
				if (!st) {
					break;
				}
				list.unshift(st);
				name = st.parent;
			}
			for (var i = 0; i < list.length; i++) {
				out = merge(out, list[i][inKind]);
			}
		};
		chain("presentation", inPresentationName);
		chain("graphic", inGraphicName);
		return out;
	};

	//* A style's graphic properties merged along its parents (draw fill and so on).
	Styles.prototype.graphic = function(inFamily, inName) {
		return this.resolve(inFamily, inName, "raw");
	};

	function loadStyles(inFiles, inOffsets) {
		var styles = new Styles();
		var sx = inFiles["styles.xml"] ? X.parse(inFiles["styles.xml"]) : null;
		var cx = inFiles["content.xml"] ? X.parse(inFiles["content.xml"], {offsets: !!inOffsets}) : null;
		// Fonts first: text properties name them.
		if (sx) {
			styles.addFonts(X.child(sx, "office:font-face-decls"));
		}
		if (cx) {
			styles.addFonts(X.child(cx, "office:font-face-decls"));
		}
		if (sx) {
			styles.read(X.child(sx, "office:styles"));
			styles.read(X.child(sx, "office:automatic-styles"));
			styles.read(X.child(sx, "office:master-styles"));
		}
		if (cx) {
			styles.read(X.child(cx, "office:automatic-styles"));
		}
		return {styles: styles, content: cx, stylesDoc: sx};
	}

	// --- Text content ----------------------------------------------------

	var numFormats = {"1": "decimal", "a": "lower-alpha", "A": "upper-alpha", "i": "lower-roman", "I": "upper-roman"};

	/**
		Walks office:text (or a text box, a table cell, a note body) and
		builds blocks. One Builder per document; it collects the images and
		the notes as it goes.
	*/
	function Builder(inStyles, inTrack) {
		this.styles = inStyles;
		// Paragraphs, by id, when the document is opened for editing.
		this.paras = inTrack || null;
		// Inside repeated table rows or cells, one paragraph is shown more than once: not editable.
		this.repeated = 0;
		this.images = [];
		this.imageSet = {};
		this.notes = [];
		this.listStack = [];
		// Inline content in one cell or paragraph is capped so a hostile
		// file cannot ask for millions of repeated spaces or cells.
		this.budget = 2000000;
	}

	Builder.prototype.useImage = function(inHref) {
		var href = String(inHref || "").replace(/^\.\//, "");
		if (!href || /^[a-z][a-z0-9+.-]*:/i.test(href)) {
			// External or data links are not fetched.
			return null;
		}
		if (!this.imageSet[href]) {
			this.imageSet[href] = true;
			this.images.push(href);
		}
		return href;
	};

	Builder.prototype.blocks = function(inEl) {
		var out = [];
		if (!inEl) {
			return out;
		}
		for (var i = 0; i < inEl.c.length; i++) {
			var el = inEl.c[i];
			if (typeof el === "object") {
				this.block(el, out);
			}
		}
		return out;
	};

	Builder.prototype.block = function(inEl, inOut) {
		switch (inEl.n) {
		case "text:p":
		case "text:h":
			this.paragraph(inEl, inOut);
			break;
		case "text:list":
			inOut.push(this.list(inEl));
			break;
		case "table:table":
			inOut.push(this.table(inEl));
			break;
		case "text:section":
		case "text:index-body":
		case "text:table-of-content":
		case "text:alphabetical-index":
		case "text:illustration-index":
		case "text:bibliography":
		case "text:user-index":
		case "text:object-index":
		case "text:table-index":
		case "office:forms":
		case "text:soft-page-break":
		case "table:table-header-rows":
		case "draw:text-box":
			if (inEl.n !== "office:forms" && inEl.n !== "text:soft-page-break") {
				var inner = this.blocks(inEl);
				for (var i = 0; i < inner.length; i++) {
					inOut.push(inner[i]);
				}
			}
			break;
		case "draw:frame":
		case "draw:custom-shape":
		case "draw:rect":
			// A frame anchored to the page or paragraph rather than a character.
			var r = [];
			this.inline(inEl, r, {});
			if (r.length) {
				inOut.push({t: "p", h: 0, s: {}, r: r});
			}
			break;
		default:
			// Tracked changes, declarations, index titles' templates and the like: skip.
			if (inEl.n === "text:index-title") {
				var t = this.blocks(inEl);
				for (var j = 0; j < t.length; j++) {
					inOut.push(t[j]);
				}
			}
			break;
		}
	};

	Builder.prototype.paragraph = function(inEl, inOut) {
		var styleName = inEl.a["text:style-name"];
		var st = this.styles.get("paragraph", styleName);
		if (st && st.breakBefore && inOut.length) {
			inOut.push({t: "pagebreak"});
		}
		var css = merge(this.styles.resolve("paragraph", styleName, "para"),
			this.styles.resolve("paragraph", styleName, "text"));
		// The id first, so it is in document order (before the paragraphs of its notes and text boxes).
		var id = this.paras && !this.repeated && inEl.p !== undefined ? this.paras.push(inEl) - 1 : -1;
		var runs = [];
		this.inline(inEl, runs, {});
		var level = inEl.n === "text:h" ? Math.max(1, Math.min(6, +(inEl.a["text:outline-level"] || 1))) : 0;
		var block = {t: "p", h: level, s: css, r: runs};
		if (id >= 0 && this.budget > 0) {
			block.id = id;
			block.pt = plainText(tokens(inEl));
		}
		inOut.push(block);
	};

	//* Appends the inline content of inEl to inOut; inCss is the span style in force.
	Builder.prototype.inline = function(inEl, inOut, inCss) {
		for (var i = 0; i < inEl.c.length; i++) {
			var c = inEl.c[i];
			if (typeof c === "string") {
				// ODF collapses white space in text; the explicit text:s are the real spaces.
				var v = c.replace(/[\r\n\t]+/g, " ").replace(/ {2,}/g, " ");
				if (v) {
					this.text(inOut, v, inCss);
				}
				continue;
			}
			switch (c.n) {
			case "text:span":
				var spanCss = merge(inCss, this.styles.resolve("text", c.a["text:style-name"], "text"));
				this.inline(c, inOut, spanCss);
				break;
			case "text:s":
				var n = Math.min(+(c.a["text:c"] || 1) || 1, 200);
				this.text(inOut, new Array(n + 1).join(" "), inCss);
				break;
			case "text:tab":
				inOut.push({t: "tab"});
				break;
			case "text:line-break":
				inOut.push({t: "br"});
				break;
			case "text:a":
				var link = {t: "a", href: c.a["xlink:href"] || "", r: []};
				this.inline(c, link.r, inCss);
				inOut.push(link);
				break;
			case "text:note":
				var label = X.textOf(X.child(c, "text:note-citation")) || String(this.notes.length + 1);
				this.notes.push({label: label, b: this.blocks(X.child(c, "text:note-body"))});
				inOut.push({t: "note", label: label});
				break;
			case "draw:frame":
				this.frame(c, inOut);
				break;
			case "draw:a":
				this.inline(c, inOut, inCss);
				break;
			case "draw:custom-shape":
			case "draw:rect":
			case "draw:ellipse":
				// A drawing with words in it: its words.
				var shapeBlocks = this.blocks(c);
				if (shapeBlocks.length) {
					inOut.push({t: "frame", s: {}, b: shapeBlocks});
				}
				break;
			case "text:soft-page-break":
			case "text:bookmark":
			case "text:bookmark-start":
			case "text:bookmark-end":
			case "text:reference-mark":
			case "text:reference-mark-start":
			case "text:reference-mark-end":
			case "office:annotation":
			case "office:annotation-end":
			case "text:change":
			case "text:change-start":
			case "text:change-end":
			case "text:alphabetical-index-mark":
			case "text:toc-mark":
				break;
			default:
				// Fields (dates, page numbers, references, ...): their current text.
				this.inline(c, inOut, inCss);
				break;
			}
		}
	};

	Builder.prototype.text = function(inOut, inText, inCss) {
		if (this.budget <= 0) {
			return;
		}
		this.budget -= inText.length;
		var last = inOut[inOut.length - 1];
		if (last && last.t === "text" && sameCss(last.s, inCss)) {
			last.v += inText;
		} else {
			var run = {t: "text", v: inText};
			if (!isEmpty(inCss)) {
				run.s = inCss;
			}
			inOut.push(run);
		}
	};

	function sameCss(inA, inB) {
		var a = inA || {}, b = inB || {}, k;
		for (k in a) {
			if (a.hasOwnProperty(k) && a[k] !== b[k]) {
				return false;
			}
		}
		for (k in b) {
			if (b.hasOwnProperty(k) && a[k] !== b[k]) {
				return false;
			}
		}
		return true;
	}

	//* A draw:frame in running text: a picture, a text box, or an object's picture.
	Builder.prototype.frame = function(inEl, inOut) {
		var w = inEl.a["svg:width"], h = inEl.a["svg:height"];
		var image = X.child(inEl, "draw:image");
		var box = X.child(inEl, "draw:text-box");
		var object = X.child(inEl, "draw:object") || X.child(inEl, "draw:object-ole");
		if (box) {
			var css = {};
			put(css, "width", w);
			inOut.push({t: "frame", s: css, b: this.blocks(box)});
			return;
		}
		// An object (chart, formula) is shown by the picture saved with it, if any.
		var images = X.children(inEl, "draw:image");
		if (object && images.length) {
			image = images[images.length - 1];
		}
		if (image) {
			var src = this.useImage(image.a["xlink:href"]);
			var alt = X.textOf(X.child(inEl, "svg:title")) || X.textOf(X.child(inEl, "svg:desc")) || "";
			if (src) {
				inOut.push({t: "img", src: src, w: w || "", h: h || "", alt: alt});
			} else if (alt) {
				this.text(inOut, "[" + alt + "]", {});
			}
		} else if (object) {
			this.text(inOut, "[object]", {fontStyle: "italic"});
		}
	};

	Builder.prototype.list = function(inEl) {
		var styleName = inEl.a["text:style-name"];
		var parent = this.listStack[this.listStack.length - 1];
		var listStyleName = styleName || (parent && parent.style);
		// A list without a style takes the one its first paragraph's style names.
		if (!listStyleName) {
			var firstP = X.findAll(inEl, "text:p")[0] || X.findAll(inEl, "text:h")[0];
			listStyleName = firstP ? this.styles.listOfParagraph(firstP.a["text:style-name"]) : null;
		}
		var depth = this.listStack.length + 1;
		var levels = (listStyleName && this.styles.lists[listStyleName]) || {};
		var level = levels[depth] || levels[1] || {ordered: false};
		this.listStack.push({style: listStyleName});
		var items = [];
		var kids = X.children(inEl);
		for (var i = 0; i < kids.length; i++) {
			if (kids[i].n === "text:list-item" || kids[i].n === "text:list-header") {
				items.push(this.blocks(kids[i]));
			}
		}
		this.listStack.pop();
		var out = {t: "list", o: !!level.ordered, items: items};
		if (level.ordered) {
			out.type = level.format === "" ? "none" : (numFormats[level.format] || "decimal");
			out.start = level.start || 1;
		}
		return out;
	};

	//* Repeat counts are capped: trailing "1024 empty columns" are common and never content.
	var maxRepeat = 256;

	Builder.prototype.table = function(inEl) {
		var styles = this.styles;
		var name = inEl.a["table:style-name"];
		var cols = [];
		var rows = [];
		var self = this;
		var addRow = function(inRow, inHeader) {
			var repeat = Math.min(+(inRow.a["table:number-rows-repeated"] || 1) || 1, maxRepeat);
			if (repeat > 1) {
				self.repeated++;
			}
			var rowCss = styles.resolve("table-row", inRow.a["table:style-name"], "cell");
			var cells = [];
			var kids = X.children(inRow);
			for (var c = 0; c < kids.length; c++) {
				var k = kids[c];
				if (k.n !== "table:table-cell" && k.n !== "table:covered-table-cell") {
					continue;
				}
				var reps = Math.min(+(k.a["table:number-columns-repeated"] || 1) || 1, maxRepeat);
				if (k.n === "table:covered-table-cell") {
					continue;
				}
				if (reps > 1) {
					self.repeated++;
				}
				for (var r = 0; r < reps; r++) {
					cells.push({
						cs: +(k.a["table:number-columns-spanned"] || 1) || 1,
						rs: +(k.a["table:number-rows-spanned"] || 1) || 1,
						s: merge(styles.resolve("table-cell", k.a["table:style-name"], "cell"),
							styles.resolve("table-cell", k.a["table:style-name"], "text")),
						b: self.blocks(k)
					});
				}
				if (reps > 1) {
					self.repeated--;
				}
			}
			if (repeat > 1) {
				self.repeated--;
			}
			for (var i = 0; i < repeat; i++) {
				rows.push({header: inHeader, s: rowCss, cells: cells});
			}
		};
		var walk = function(inParent, inHeader) {
			var kids = X.children(inParent);
			for (var i = 0; i < kids.length; i++) {
				var k = kids[i];
				if (k.n === "table:table-column") {
					var n = Math.min(+(k.a["table:number-columns-repeated"] || 1) || 1, maxRepeat);
					var col = styles.get("table-column", k.a["table:style-name"]);
					for (var j = 0; j < n; j++) {
						cols.push(col && col.col.width ? col.col.width : "");
					}
				} else if (k.n === "table:table-columns" || k.n === "table:table-column-group" || k.n === "table:table-header-columns") {
					walk(k, inHeader);
				} else if (k.n === "table:table-row") {
					addRow(k, inHeader);
				} else if (k.n === "table:table-header-rows") {
					walk(k, true);
				} else if (k.n === "table:table-rows" || k.n === "table:table-row-group") {
					walk(k, inHeader);
				}
			}
		};
		walk(inEl, false);
		// Drop trailing empty columns and rows (repeated filler from the editor).
		trimTable(rows, cols);
		return {t: "table", s: styles.resolve("table", name, "table"), cols: cols, rows: rows};
	};

	function cellIsEmpty(inCell) {
		for (var i = 0; i < inCell.b.length; i++) {
			var b = inCell.b[i];
			if (b.t !== "p" || b.r.length) {
				return false;
			}
		}
		return true;
	}

	function trimTable(inRows, inCols) {
		while (inRows.length > 1) {
			var last = inRows[inRows.length - 1];
			var empty = true;
			for (var i = 0; i < last.cells.length && empty; i++) {
				empty = cellIsEmpty(last.cells[i]);
			}
			if (!empty) {
				break;
			}
			inRows.pop();
		}
		var used = 0;
		for (var r = 0; r < inRows.length; r++) {
			var cells = inRows[r].cells;
			var width = 0;
			for (var c = 0; c < cells.length; c++) {
				if (!cellIsEmpty(cells[c])) {
					width = c + 1;
				}
			}
			used = Math.max(used, width);
		}
		used = Math.max(used, 1);
		for (var r2 = 0; r2 < inRows.length; r2++) {
			if (inRows[r2].cells.length > used) {
				inRows[r2].cells.length = used;
			}
		}
		if (inCols.length > used) {
			inCols.length = used;
		}
	}

	// --- Metadata --------------------------------------------------------

	function meta(inMetaXml) {
		var out = {};
		if (!inMetaXml) {
			return out;
		}
		var m;
		try {
			m = X.path(X.parse(inMetaXml), "office:meta");
		} catch (e) {
			return out;
		}
		var fields = {"dc:title": "title", "dc:subject": "subject", "dc:creator": "author", "meta:initial-creator": "createdBy",
			"dc:date": "modified", "meta:creation-date": "created", "meta:generator": "generator", "dc:description": "description"};
		var kids = X.children(m);
		for (var i = 0; i < kids.length; i++) {
			if (fields[kids[i].n]) {
				out[fields[kids[i].n]] = X.textOf(kids[i]).trim();
			}
			if (kids[i].n === "meta:document-statistic") {
				var a = kids[i].a;
				if (a["meta:page-count"]) {
					out.pages = +a["meta:page-count"];
				}
				if (a["meta:word-count"]) {
					out.words = +a["meta:word-count"];
				}
			}
		}
		return out;
	}

	// --- Documents -------------------------------------------------------

	function pageOf(inStyles, inMasterName) {
		var masters = inStyles.masterPages;
		var master = masters[inMasterName] || masters.Standard || masters.Default;
		if (!master) {
			for (var k in masters) {
				if (masters.hasOwnProperty(k)) {
					master = masters[k];
					break;
				}
			}
		}
		var layout = master ? inStyles.pageLayouts[master.a["style:page-layout-name"]] || {} : {};
		return {
			width: layout["fo:page-width"] || "21cm",
			height: layout["fo:page-height"] || "29.7cm",
			marginTop: layout["fo:margin-top"] || "2cm",
			marginBottom: layout["fo:margin-bottom"] || "2cm",
			marginLeft: layout["fo:margin-left"] || "2cm",
			marginRight: layout["fo:margin-right"] || "2cm"
		};
	}

	/**
		An OpenDocument text. inFiles holds the package's XML parts as
		strings: "content.xml" (required), "styles.xml", "meta.xml".
	*/
	// --- Paragraph text, for editing ---------------------------------------

	/**
		A paragraph's content as tokens, in the order the reader shows it:
			{k: "t", v, s, e}    text; s, e: where it is in the XML
			{k: "w", v, el}      a space run (text:s), a tab or a line break
			{k: "o", v, el}      an object: a picture, a note, a frame ("\ufffc"), a field (its text)
		Elements the reader skips (bookmarks, marks) give no token and are kept.
		Needs the paragraph parsed with offsets.
	*/
	function tokens(inEl, inOut) {
		var out = inOut || [];
		for (var i = 0; i < inEl.c.length; i++) {
			var c = inEl.c[i];
			if (typeof c === "string") {
				// As Builder.inline collapses it.
				var v = c.replace(/[\r\n\t]+/g, " ").replace(/ {2,}/g, " ");
				if (v && inEl.co && inEl.co[i]) {
					out.push({k: "t", v: v, s: inEl.co[i][0], e: inEl.co[i][1]});
				}
				continue;
			}
			switch (c.n) {
			case "text:span":
			case "text:a":
				tokens(c, out);
				break;
			case "text:s":
				out.push({k: "w", v: new Array(Math.min(+(c.a["text:c"] || 1) || 1, 200) + 1).join(" "), el: c});
				break;
			case "text:tab":
				out.push({k: "w", v: "\t", el: c});
				break;
			case "text:line-break":
				out.push({k: "w", v: "\n", el: c});
				break;
			case "text:note":
			case "draw:frame":
			case "draw:a":
			case "draw:custom-shape":
			case "draw:rect":
			case "draw:ellipse":
				out.push({k: "o", v: "\ufffc", el: c});
				break;
			case "text:soft-page-break":
			case "text:bookmark":
			case "text:bookmark-start":
			case "text:bookmark-end":
			case "text:reference-mark":
			case "text:reference-mark-start":
			case "text:reference-mark-end":
			case "office:annotation":
			case "office:annotation-end":
			case "text:change":
			case "text:change-start":
			case "text:change-end":
			case "text:alphabetical-index-mark":
			case "text:toc-mark":
				break;
			default:
				// A field (date, page number, reference): one piece, its current text.
				var f = plainText(tokens(c, []));
				if (f) {
					out.push({k: "o", v: f, el: c});
				}
				break;
			}
		}
		return out;
	}

	function plainText(inTokens) {
		var s = "";
		for (var i = 0; i < inTokens.length; i++) {
			s += inTokens[i].v;
		}
		return s;
	}

	/**
		One paragraph's block again, after it was edited (inEl parsed with
		offsets). inIdOf(el) gives the id of it and of the paragraphs in it
		(in its text boxes), -1 for one that has none.
	*/
	function paragraphBlock(inStyles, inEl, inIdOf) {
		var local = [];
		var b = new Builder(inStyles, local);
		var out = [];
		b.paragraph(inEl, out);
		var block = out[out.length - 1];
		var fix = function(inBlocks) {
			(inBlocks || []).forEach(function(x) {
				if (x.id !== undefined) {
					var id = inIdOf(local[x.id]);
					if (id >= 0) {
						x.id = id;
					} else {
						delete x.id;
						delete x.pt;
					}
				}
				(x.r || []).forEach(function(run) {
					if (run.t === "frame") {
						fix(run.b);
					}
				});
				if (x.b) {
					fix(x.b);
				}
			});
		};
		fix([block]);
		return {block: block, images: b.images};
	}

	function text(inFiles, inOptions) {
		var track = inOptions && inOptions.track;
		var loaded = loadStyles(inFiles, !!track);
		var body = X.path(loaded.content, "office:body", "office:text");
		if (!body) {
			throw new Error("ODF: no office:text in content.xml");
		}
		var b = new Builder(loaded.styles, track ? [] : null);
		if (track) {
			track.paras = b.paras;
			track.content = loaded.content;
			track.styles = loaded.styles;
		}
		var blocks = b.blocks(body);
		if (b.notes.length) {
			blocks.push({t: "notes", items: b.notes});
		}
		var defaults = merge(loaded.styles.resolve("paragraph", "Standard", "text"), {});
		return {
			kind: "odt",
			blocks: blocks,
			base: defaults,
			page: pageOf(loaded.styles, null),
			images: b.images,
			meta: meta(inFiles["meta.xml"])
		};
	}

	// --- Presentations ---------------------------------------------------

	function fillOf(inGraphic) {
		if (!inGraphic) {
			return "";
		}
		if (inGraphic["draw:fill"] === "solid" || (!inGraphic["draw:fill"] && inGraphic["draw:fill-color"])) {
			var v = inGraphic["draw:fill-color"] || "";
			return /^#[0-9a-fA-F]{3,8}$/.test(v) ? v : "";
		}
		if (inGraphic["draw:fill"] === "gradient" && inGraphic["draw:fill-color"]) {
			return /^#[0-9a-fA-F]{3,8}$/.test(inGraphic["draw:fill-color"]) ? inGraphic["draw:fill-color"] : "";
		}
		return "";
	}

	function strokeOf(inGraphic) {
		if (!inGraphic || inGraphic["draw:stroke"] === "none" || !inGraphic["draw:stroke"]) {
			return "";
		}
		var c = inGraphic["svg:stroke-color"] || "#000000";
		return /^#[0-9a-fA-F]{3,8}$/.test(c) ? c : "";
	}

	/**
		An OpenDocument presentation. Every draw:page is a slide of absolutely
		placed items in CSS pixels: text boxes (blocks), pictures and plain
		filled shapes. Master page decoration is not drawn.
	*/
	function presentation(inFiles) {
		var loaded = loadStyles(inFiles);
		var styles = loaded.styles;
		var body = X.path(loaded.content, "office:body", "office:presentation");
		if (!body) {
			throw new Error("ODF: no office:presentation in content.xml");
		}
		var b = new Builder(styles);
		var slides = [];
		var pages = X.children(body, "draw:page");
		for (var i = 0; i < pages.length; i++) {
			var page = pages[i];
			var size = pageOf(styles, page.a["draw:master-page-name"]);
			var w = toPx(size.width) || 1024, h = toPx(size.height) || 768;
			var master = styles.masterPages[page.a["draw:master-page-name"]];
			var bg = fillOf(styles.graphic("drawing-page", page.a["draw:style-name"])) ||
				(master ? fillOf(styles.graphic("drawing-page", master.a["draw:style-name"])) : "") || "#ffffff";
			var items = [];
			collectShapes(page, items, b, styles, 0);
			var notes = X.child(page, "presentation:notes");
			var noteText = "";
			if (notes) {
				var nf = X.findAll(notes, "draw:text-box");
				for (var n = 0; n < nf.length; n++) {
					noteText += X.findAll(nf[n], "text:p").map(X.textOf).join("\n");
				}
			}
			slides.push({name: page.a["draw:name"] || "", width: Math.round(w), height: Math.round(h), bg: bg,
				items: items, notes: noteText});
		}
		return {kind: "odp", slides: slides, images: b.images, meta: meta(inFiles["meta.xml"])};
	}

	function collectShapes(inParent, inItems, inBuilder, inStyles, inDepth) {
		if (inDepth > 16) {
			return;
		}
		var kids = X.children(inParent);
		for (var i = 0; i < kids.length; i++) {
			var el = kids[i];
			var a = el.a;
			if (el.n === "draw:g") {
				collectShapes(el, inItems, inBuilder, inStyles, inDepth + 1);
				continue;
			}
			if (["draw:frame", "draw:custom-shape", "draw:rect", "draw:ellipse"].indexOf(el.n) < 0) {
				continue;
			}
			var item = {
				x: Math.round(toPx(a["svg:x"]) || 0), y: Math.round(toPx(a["svg:y"]) || 0),
				w: Math.round(toPx(a["svg:width"]) || 0), h: Math.round(toPx(a["svg:height"]) || 0)
			};
			// The family defaults first, then the presentation style (title, outline), then the shape's own.
			var g = inStyles.layered(a["presentation:style-name"], a["draw:style-name"], "raw");
			var textCss = inStyles.layered(a["presentation:style-name"], a["draw:style-name"], "text");
			var paraCss = inStyles.layered(a["presentation:style-name"], a["draw:style-name"], "para");
			item.fill = fillOf(g);
			item.stroke = strokeOf(g);
			item.round = el.n === "draw:ellipse" || (el.n === "draw:custom-shape" &&
				/ellipse|circle/.test((X.child(el, "draw:enhanced-geometry") || {a: {}}).a["draw:type"] || ""));
			if (g["draw:textarea-vertical-align"]) {
				item.valign = g["draw:textarea-vertical-align"];
			}
			var image = X.child(el, "draw:image");
			var box = X.child(el, "draw:text-box");
			if (el.n === "draw:frame" && image && !box) {
				var images = X.children(el, "draw:image");
				var src = inBuilder.useImage(images[images.length - 1].a["xlink:href"]);
				if (!src) {
					continue;
				}
				item.kind = "image";
				item.src = src;
			} else if (el.n === "draw:frame" && X.child(el, "draw:object")) {
				item.kind = "shape";
				item.b = [{t: "p", h: 0, s: {}, r: [{t: "text", v: "[object]"}]}];
			} else {
				item.kind = box || el.n !== "draw:frame" ? "text" : "shape";
				var blocks = inBuilder.blocks(box || el);
				// An empty placeholder ("Click to add text") has no text of its own: skip it.
				if (!blocks.length && !item.fill && !item.stroke) {
					continue;
				}
				item.b = blocks;
				item.s = merge(paraCss, textCss);
			}
			inItems.push(item);
		}
	}

	var DocOdf = {
		text: text,
		tokens: tokens,
		plainText: plainText,
		paragraphBlock: paragraphBlock,
		presentation: presentation,
		meta: meta,
		toPx: toPx,
		css: put,
		Styles: Styles
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocOdf;
	} else {
		root.DocOdf = DocOdf;
	}
})(this);
