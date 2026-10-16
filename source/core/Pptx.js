/**
	PowerPoint (.pptx) slides to the same slide model as DocOdf.presentation:

		var deck = DocPptx.parse(files);   // files: {"ppt/presentation.xml": "...", ...} every XML part
		deck.slides  // [{width, height, bg, bgImage, items: [{x, y, w, h, kind, b, s, fill, stroke, src}]}]
		deck.images  // zip paths of the pictures used

	What is drawn, best effort: text boxes and placeholders (with the
	sizes, colours, alignment and bullets of the slide master's text styles
	and the layout's placeholders), pictures, plainly filled rectangles and
	ellipses, tables, groups, and the decoration of the layout and master.
	Not drawn: charts, SmartArt, gradients (their first colour is used),
	effects, animations, custom geometry outlines.

	Pure logic on the parsed XML; tested in node (tools/tests/unit-tests.js).
*/
(function(root) {
	var X = root.DocXml || (typeof require === "function" ? require("./Xml.js") : null);

	var EMU_PER_PX = 9525;

	function px(inEmu) {
		return Math.round((+inEmu || 0) / EMU_PER_PX);
	}

	// --- Paths and relationships -----------------------------------------

	function dirOf(inPath) {
		var i = inPath.lastIndexOf("/");
		return i < 0 ? "" : inPath.slice(0, i + 1);
	}

	//* Resolves a relationship target against the part it belongs to.
	function resolvePath(inBase, inTarget) {
		if (!inTarget) {
			return "";
		}
		if (inTarget.charAt(0) === "/") {
			return inTarget.slice(1);
		}
		var parts = (dirOf(inBase) + inTarget).split("/");
		var out = [];
		for (var i = 0; i < parts.length; i++) {
			if (parts[i] === "..") {
				out.pop();
			} else if (parts[i] !== "." && parts[i] !== "") {
				out.push(parts[i]);
			}
		}
		return out.join("/");
	}

	function relsPath(inPart) {
		return dirOf(inPart) + "_rels/" + inPart.slice(dirOf(inPart).length) + ".rels";
	}

	//* {rId: {target (resolved path), type, external}} of a part.
	function rels(inFiles, inPart) {
		var out = {};
		var text = inFiles[relsPath(inPart)];
		if (!text) {
			return out;
		}
		var r = X.parse(text);
		var list = X.children(r);
		for (var i = 0; i < list.length; i++) {
			var a = list[i].a;
			var external = a.TargetMode === "External";
			out[a.Id] = {target: external ? a.Target : resolvePath(inPart, a.Target), type: (a.Type || "").replace(/^.*\//, ""),
				external: external};
		}
		return out;
	}

	function relOfType(inRels, inType) {
		for (var k in inRels) {
			if (inRels.hasOwnProperty(k) && inRels[k].type === inType) {
				return inRels[k];
			}
		}
		return null;
	}

	// --- Colours ---------------------------------------------------------

	var presetColors = {black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00",
		gray: "808080", grey: "808080", darkGray: "A9A9A9", lightGray: "D3D3D3", orange: "FFA500", navy: "000080"};

	function hexToRgb(inHex) {
		var n = parseInt(inHex, 16);
		return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
	}

	function rgbToHex(inRgb) {
		var s = "#";
		for (var i = 0; i < 3; i++) {
			var v = Math.max(0, Math.min(255, Math.round(inRgb[i])));
			s += (v < 16 ? "0" : "") + v.toString(16);
		}
		return s;
	}

	function rgbToHsl(inRgb) {
		var r = inRgb[0] / 255, g = inRgb[1] / 255, b = inRgb[2] / 255;
		var max = Math.max(r, g, b), min = Math.min(r, g, b);
		var h = 0, s = 0, l = (max + min) / 2;
		if (max !== min) {
			var d = max - min;
			s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
			h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
			h /= 6;
		}
		return [h, s, l];
	}

	function hslToRgb(inHsl) {
		var h = inHsl[0], s = inHsl[1], l = inHsl[2];
		if (!s) {
			return [l * 255, l * 255, l * 255];
		}
		var hue = function(p, q, t) {
			if (t < 0) {
				t += 1;
			}
			if (t > 1) {
				t -= 1;
			}
			if (t < 1 / 6) {
				return p + (q - p) * 6 * t;
			}
			if (t < 1 / 2) {
				return q;
			}
			if (t < 2 / 3) {
				return p + (q - p) * (2 / 3 - t) * 6;
			}
			return p;
		};
		var q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
		return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
	}

	/**
		The colour of a fill-like element (a:solidFill, a:fgClr, ...), with
		lumMod/lumOff/tint/shade applied; "" when there is none.
	*/
	function colorOf(inEl, inTheme) {
		if (!inEl) {
			return "";
		}
		var c = X.child(inEl);
		if (!c) {
			return "";
		}
		var hex = "";
		if (c.n === "a:srgbClr") {
			hex = c.a.val;
		} else if (c.n === "a:schemeClr") {
			hex = inTheme.scheme(c.a.val);
		} else if (c.n === "a:sysClr") {
			hex = c.a.lastClr || (c.a.val === "window" ? "FFFFFF" : "000000");
		} else if (c.n === "a:prstClr") {
			hex = presetColors[c.a.val] || "";
		} else if (c.n === "a:scrgbClr") {
			hex = rgbToHex([+c.a.r / 1000 * 2.55, +c.a.g / 1000 * 2.55, +c.a.b / 1000 * 2.55]).slice(1);
		}
		if (!/^[0-9a-fA-F]{6}$/.test(hex || "")) {
			return "";
		}
		var rgb = hexToRgb(hex);
		var mods = X.children(c);
		for (var i = 0; i < mods.length; i++) {
			var v = +mods[i].a.val / 100000;
			var hsl;
			switch (mods[i].n) {
			case "a:lumMod":
				hsl = rgbToHsl(rgb);
				hsl[2] *= v;
				rgb = hslToRgb(hsl);
				break;
			case "a:lumOff":
				hsl = rgbToHsl(rgb);
				hsl[2] = Math.min(1, hsl[2] + v);
				rgb = hslToRgb(hsl);
				break;
			case "a:tint":
				rgb = [rgb[0] + (255 - rgb[0]) * (1 - v), rgb[1] + (255 - rgb[1]) * (1 - v), rgb[2] + (255 - rgb[2]) * (1 - v)];
				break;
			case "a:shade":
				rgb = [rgb[0] * v, rgb[1] * v, rgb[2] * v];
				break;
			}
		}
		return rgbToHex(rgb);
	}

	//* A fill's colour from a:solidFill, or a gradient's first stop.
	function fillOf(inSpPr, inTheme) {
		if (!inSpPr) {
			return null;
		}
		if (X.child(inSpPr, "a:noFill")) {
			return "none";
		}
		var solid = X.child(inSpPr, "a:solidFill");
		if (solid) {
			return colorOf(solid, inTheme) || null;
		}
		var grad = X.child(inSpPr, "a:gradFill");
		if (grad) {
			var gs = X.findAll(grad, "a:gs")[0];
			return gs ? colorOf(gs, inTheme) || null : null;
		}
		return null;
	}

	function Theme(inThemeXml, inClrMap) {
		this.colors = {};
		this.map = inClrMap || {};
		if (inThemeXml) {
			var scheme = X.findAll(X.parse(inThemeXml), "a:clrScheme")[0];
			var kids = X.children(scheme);
			for (var i = 0; i < kids.length; i++) {
				var c = X.child(kids[i]);
				if (c) {
					var name = kids[i].n.replace(/^a:/, "");
					this.colors[name] = c.n === "a:sysClr" ? (c.a.lastClr || "000000") : (c.a.val || "");
				}
			}
		}
	}

	Theme.prototype.scheme = function(inName) {
		var name = inName;
		var aliases = {bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2"};
		if (this.map[name]) {
			name = this.map[name];
		} else if (aliases[name]) {
			name = aliases[name];
		}
		return this.colors[name] || (name === "phClr" ? "" : "");
	};

	// --- Text ------------------------------------------------------------

	var aligns = {l: "left", ctr: "center", r: "right", just: "justify", dist: "justify"};
	var autoNums = {arabicPeriod: [".", "decimal"], arabicParenR: [")", "decimal"], alphaLcPeriod: [".", "lower-alpha"],
		alphaUcPeriod: [".", "upper-alpha"], romanLcPeriod: [".", "lower-roman"], romanUcPeriod: [".", "upper-roman"],
		alphaLcParenR: [")", "lower-alpha"], arabicPlain: ["", "decimal"]};

	function toAlpha(inN, inUpper) {
		var s = "";
		var n = inN;
		while (n > 0) {
			n--;
			s = String.fromCharCode(97 + (n % 26)) + s;
			n = Math.floor(n / 26);
		}
		return inUpper ? s.toUpperCase() : s;
	}

	function toRoman(inN, inUpper) {
		var vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
		var syms = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"];
		var s = "", n = inN;
		for (var i = 0; i < vals.length; i++) {
			while (n >= vals[i]) {
				s += syms[i];
				n -= vals[i];
			}
		}
		return inUpper ? s.toUpperCase() : s;
	}

	function numberLabel(inN, inScheme) {
		var spec = autoNums[inScheme] || autoNums.arabicPeriod;
		var t = spec[1];
		var v = t === "lower-alpha" ? toAlpha(inN) : t === "upper-alpha" ? toAlpha(inN, true) :
			t === "lower-roman" ? toRoman(inN) : t === "upper-roman" ? toRoman(inN, true) : String(inN);
		return v + spec[0];
	}

	/**
		Paragraph level properties from one a:lvlNpPr / a:pPr (and its
		a:defRPr), merged over inBase: {align, size (pt), color, bold, italic,
		bullet: "char"|"auto"|"none", buChar, buScheme, marL, indent, font}.
	*/
	function levelProps(inPPr, inBase, inTheme) {
		var out = {};
		var k;
		for (k in inBase) {
			if (inBase.hasOwnProperty(k)) {
				out[k] = inBase[k];
			}
		}
		if (!inPPr) {
			return out;
		}
		var a = inPPr.a;
		if (a.algn) {
			out.align = aligns[a.algn] || "left";
		}
		if (a.marL !== undefined) {
			out.marL = px(a.marL);
		}
		if (a.indent !== undefined) {
			out.indent = px(a.indent);
		}
		if (X.child(inPPr, "a:buNone")) {
			out.bullet = "none";
		} else if (X.child(inPPr, "a:buChar")) {
			var buFont = X.child(inPPr, "a:buFont");
			out.buFont = buFont ? buFont.a.typeface || "" : "";
			out.bullet = "char";
			out.buChar = X.child(inPPr, "a:buChar").a["char"] || "\u2022";
		} else if (X.child(inPPr, "a:buAutoNum")) {
			out.bullet = "auto";
			out.buScheme = X.child(inPPr, "a:buAutoNum").a.type;
			out.buStart = +(X.child(inPPr, "a:buAutoNum").a.startAt || 1);
		}
		var buClr = X.child(inPPr, "a:buClr");
		if (buClr) {
			out.buColor = colorOf(buClr, inTheme);
		}
		var lnSpc = X.child(X.child(inPPr, "a:lnSpc"), "a:spcPct");
		if (lnSpc) {
			out.lineHeight = +lnSpc.a.val / 100000;
		}
		runProps(X.child(inPPr, "a:defRPr"), out, inTheme);
		return out;
	}

	//* Run properties (a:rPr, a:defRPr, a:endParaRPr) over inOut.
	function runProps(inRPr, inOut, inTheme) {
		if (!inRPr) {
			return inOut;
		}
		var a = inRPr.a;
		if (a.sz) {
			inOut.size = +a.sz / 100;
		}
		if (a.b !== undefined) {
			inOut.bold = a.b === "1" || a.b === "true";
		}
		if (a.i !== undefined) {
			inOut.italic = a.i === "1" || a.i === "true";
		}
		if (a.u !== undefined) {
			inOut.underline = a.u !== "none";
		}
		if (a.strike !== undefined) {
			inOut.strike = a.strike !== "noStrike";
		}
		if (a.baseline) {
			inOut.baseline = +a.baseline;
		}
		if (a.cap === "all") {
			inOut.caps = true;
		}
		var fill = X.child(inRPr, "a:solidFill");
		if (fill) {
			inOut.color = colorOf(fill, inTheme) || inOut.color;
		}
		var latin = X.child(inRPr, "a:latin");
		if (latin && latin.a.typeface && latin.a.typeface.charAt(0) !== "+") {
			inOut.font = latin.a.typeface;
		}
		return inOut;
	}

	function runCss(inProps, inScale) {
		var css = {};
		if (inProps.size) {
			css.fontSize = (Math.round(inProps.size * (inScale || 1) * 96 / 72 * 10) / 10) + "px";
		}
		if (inProps.bold) {
			css.fontWeight = "bold";
		}
		if (inProps.italic) {
			css.fontStyle = "italic";
		}
		var deco = [];
		if (inProps.underline) {
			deco.push("underline");
		}
		if (inProps.strike) {
			deco.push("line-through");
		}
		if (deco.length) {
			css.textDecoration = deco.join(" ");
		}
		if (inProps.color) {
			css.color = inProps.color;
		}
		if (inProps.baseline) {
			css.verticalAlign = inProps.baseline > 0 ? "super" : "sub";
			css.fontSize = "0.7em";
		}
		if (inProps.caps) {
			css.textTransform = "uppercase";
		}
		if (inProps.font && /^[-a-zA-Z0-9 ]+$/.test(inProps.font)) {
			css.fontFamily = "'" + inProps.font + "', sans-serif";
		}
		return css;
	}

	// --- Slides ----------------------------------------------------------

	function Deck(inFiles) {
		this.files = inFiles;
		this.images = [];
		this.imageSet = {};
		this.cache = {};
	}

	Deck.prototype.xml = function(inPath) {
		if (!this.cache.hasOwnProperty(inPath)) {
			var t = this.files[inPath];
			this.cache[inPath] = t ? X.parse(t) : null;
		}
		return this.cache[inPath];
	};

	Deck.prototype.useImage = function(inPath) {
		if (!inPath) {
			return null;
		}
		if (!this.imageSet[inPath]) {
			this.imageSet[inPath] = true;
			this.images.push(inPath);
		}
		return inPath;
	};

	/**
		The placeholder key of a shape ({type, idx}) or null: matched against
		the layout and master to inherit position and text style.
	*/
	function placeholderOf(inSp) {
		var nv = X.child(inSp, "p:nvSpPr") || X.child(inSp, "p:nvPicPr") || X.child(inSp, "p:nvGraphicFramePr");
		var ph = X.path(nv, "p:nvPr", "p:ph");
		if (!ph) {
			return null;
		}
		return {type: ph.a.type || "body", idx: ph.a.idx !== undefined ? ph.a.idx : null};
	}

	function findPlaceholder(inTree, inPh) {
		if (!inTree || !inPh) {
			return null;
		}
		var sps = X.findAll(inTree, "p:sp");
		var byType = null, byIdx = null;
		for (var i = 0; i < sps.length; i++) {
			var p = placeholderOf(sps[i]);
			if (!p) {
				continue;
			}
			if (inPh.idx !== null && p.idx === inPh.idx) {
				byIdx = byIdx || sps[i];
			}
			var t = p.type === "ctrTitle" ? "title" : p.type === "subTitle" ? "body" : p.type;
			var want = inPh.type === "ctrTitle" ? "title" : inPh.type === "subTitle" ? "body" : inPh.type;
			if (t === want) {
				byType = byType || sps[i];
			}
		}
		return byIdx || byType;
	}

	function xfrmOf(inSp) {
		var spPr = X.child(inSp, "p:spPr") || X.child(inSp, "p:grpSpPr");
		var xfrm = X.child(spPr, "a:xfrm") || X.child(inSp, "p:xfrm");
		if (!xfrm) {
			return null;
		}
		var off = X.child(xfrm, "a:off"), ext = X.child(xfrm, "a:ext");
		if (!off || !ext) {
			return null;
		}
		var out = {x: +off.a.x || 0, y: +off.a.y || 0, cx: +ext.a.cx || 0, cy: +ext.a.cy || 0};
		var chOff = X.child(xfrm, "a:chOff"), chExt = X.child(xfrm, "a:chExt");
		if (chOff && chExt) {
			out.chx = +chOff.a.x || 0;
			out.chy = +chOff.a.y || 0;
			out.chcx = +chExt.a.cx || 0;
			out.chcy = +chExt.a.cy || 0;
		}
		return out;
	}

	/**
		Builds one slide (or the decoration of a layout or master) into
		inItems. inCtx: {part, rels, layoutTree, masterTree, theme, txStyles,
		transform}.
	*/
	Deck.prototype.shapes = function(inTree, inCtx, inItems, inOnlyDecoration, inDepth) {
		if (!inTree || (inDepth || 0) > 16) {
			return;
		}
		var kids = X.children(inTree);
		for (var i = 0; i < kids.length; i++) {
			var el = kids[i];
			if (el.n === "p:sp" || el.n === "p:pic" || el.n === "p:graphicFrame" || el.n === "p:cxnSp") {
				var ph = placeholderOf(el);
				if (inOnlyDecoration && ph) {
					continue;
				}
				var item = this.shape(el, ph, inCtx);
				if (item) {
					inItems.push(item);
				}
			} else if (el.n === "p:grpSp") {
				var g = xfrmOf(el);
				var ctx = inCtx;
				if (g && g.chcx && g.chcy) {
					ctx = Object.create(inCtx);
					ctx.transform = {parent: inCtx.transform, g: g};
				}
				this.shapes(el, ctx, inItems, inOnlyDecoration, (inDepth || 0) + 1);
			} else if (el.n === "mc:AlternateContent") {
				// Newer content with a fallback: take the fallback, which older readers understand.
				var fb = X.child(el, "mc:Fallback");
				if (fb) {
					this.shapes(fb, inCtx, inItems, inOnlyDecoration, (inDepth || 0) + 1);
				}
			}
		}
	};

	function applyTransform(inT, inBox) {
		var b = inBox;
		var t = inT;
		while (t) {
			var g = t.g;
			var sx = g.cx / g.chcx, sy = g.cy / g.chcy;
			b = {x: g.x + (b.x - g.chx) * sx, y: g.y + (b.y - g.chy) * sy, cx: b.cx * sx, cy: b.cy * sy};
			t = t.parent;
		}
		return b;
	}

	Deck.prototype.shape = function(inEl, inPh, inCtx) {
		var theme = inCtx.theme;
		var layoutSp = inPh ? findPlaceholder(inCtx.layoutTree, inPh) : null;
		var masterSp = inPh ? findPlaceholder(inCtx.masterTree, inPh) : null;
		var box = xfrmOf(inEl) || (layoutSp && xfrmOf(layoutSp)) || (masterSp && xfrmOf(masterSp));
		if (!box) {
			return null;
		}
		if (inCtx.transform) {
			box = applyTransform(inCtx.transform, box);
		}
		var item = {x: px(box.x), y: px(box.y), w: px(box.cx), h: px(box.cy)};
		var spPr = X.child(inEl, "p:spPr");
		var fill = fillOf(spPr, theme);
		if (fill === null && layoutSp) {
			fill = fillOf(X.child(layoutSp, "p:spPr"), theme);
		}
		if (fill === null) {
			// The shape's style reference (a:fillRef with a colour) when nothing overrides it.
			var fillRef = X.path(inEl, "p:style", "a:fillRef");
			if (fillRef && +fillRef.a.idx > 0) {
				fill = colorOf(fillRef, theme) || null;
			}
		}
		item.fill = fill && fill !== "none" ? fill : "";
		var ln = X.child(spPr, "a:ln");
		if (ln && !X.child(ln, "a:noFill")) {
			item.stroke = colorOf(X.child(ln, "a:solidFill"), theme) || "";
		} else if (!ln) {
			var lnRef = X.path(inEl, "p:style", "a:lnRef");
			item.stroke = lnRef && +lnRef.a.idx > 0 ? colorOf(lnRef, theme) : "";
		} else {
			item.stroke = "";
		}
		var geom = X.child(spPr, "a:prstGeom");
		var prst = geom ? geom.a.prst : "rect";
		item.round = prst === "ellipse" || prst === "circle";
		item.radius = prst === "roundRect";
		if (inEl.n === "p:cxnSp" || prst === "line" || prst === "straightConnector1") {
			item.kind = "line";
			item.flipH = !!(X.path(spPr, "a:xfrm") && X.path(spPr, "a:xfrm").a.flipH === "1");
			item.flipV = !!(X.path(spPr, "a:xfrm") && X.path(spPr, "a:xfrm").a.flipV === "1");
			item.stroke = item.stroke || "#000000";
			return item;
		}

		if (inEl.n === "p:pic") {
			var blip = X.findAll(X.child(inEl, "p:blipFill"), "a:blip")[0];
			var rel = blip && inCtx.rels[blip.a["r:embed"]];
			if (!rel || rel.external) {
				return null;
			}
			item.kind = "image";
			item.src = this.useImage(rel.target);
			return item;
		}

		if (inEl.n === "p:graphicFrame") {
			var data = X.findAll(inEl, "a:graphicData")[0];
			var tbl = data && X.child(data, "a:tbl");
			if (tbl) {
				item.kind = "text";
				item.b = [this.table(tbl, theme, inCtx)];
				item.s = {};
				return item;
			}
			var uri = data ? data.a.uri || "" : "";
			item.kind = "shape";
			item.stroke = item.stroke || "#999999";
			item.b = [{t: "p", h: 0, s: {textAlign: "center", color: "#666666"}, r: [{t: "text",
				v: /chart/.test(uri) ? "[chart]" : /diagram/.test(uri) ? "[diagram]" : "[object]"}]}];
			item.valign = "middle";
			return item;
		}

		// p:sp
		var txBody = X.child(inEl, "p:txBody");
		var blocks = txBody ? this.text(txBody, inPh, layoutSp, masterSp, inCtx) : [];
		var hasText = false;
		for (var i = 0; i < blocks.length; i++) {
			if (blocks[i].r && blocks[i].r.length) {
				hasText = true;
			}
		}
		if (!hasText && !item.fill && !item.stroke) {
			return null;
		}
		item.kind = hasText ? "text" : "shape";
		item.b = hasText ? blocks : [];
		item.s = {};
		var bodyPr = txBody ? X.child(txBody, "a:bodyPr") : null;
		var layoutBodyPr = layoutSp ? X.path(layoutSp, "p:txBody", "a:bodyPr") : null;
		var masterBodyPr = masterSp ? X.path(masterSp, "p:txBody", "a:bodyPr") : null;
		var anchor = (bodyPr && bodyPr.a.anchor) || (layoutBodyPr && layoutBodyPr.a.anchor) ||
			(masterBodyPr && masterBodyPr.a.anchor) || (inPh && /title/i.test(inPh.type) ? "ctr" : "t");
		item.valign = anchor === "ctr" ? "middle" : anchor === "b" ? "bottom" : "top";
		var ins = function(inName, inDefault) {
			var v = bodyPr && bodyPr.a[inName];
			return px(v !== undefined ? v : inDefault);
		};
		item.pad = [ins("tIns", 45720), ins("rIns", 91440), ins("bIns", 45720), ins("lIns", 91440)];
		return item;
	};

	/**
		Paragraphs of a text body, with the inherited level properties:
		master text style -> master placeholder -> layout placeholder ->
		the shape's own list style -> the paragraph.
	*/
	Deck.prototype.text = function(inTxBody, inPh, inLayoutSp, inMasterSp, inCtx) {
		var theme = inCtx.theme;
		var styleKey = !inPh ? "p:otherStyle" : /title/i.test(inPh.type) ? "p:titleStyle" :
			/^(body|obj|subTitle)$/.test(inPh.type) || inPh.idx !== null && !/^(dt|ftr|sldNum)$/.test(inPh.type) ? "p:bodyStyle" :
			"p:otherStyle";
		var masterStyle = inCtx.txStyles ? X.child(inCtx.txStyles, styleKey) : null;
		var lists = [
			masterStyle,
			inMasterSp ? X.path(inMasterSp, "p:txBody", "a:lstStyle") : null,
			inLayoutSp ? X.path(inLayoutSp, "p:txBody", "a:lstStyle") : null,
			X.child(inTxBody, "a:lstStyle")
		];
		var bodyPr = X.child(inTxBody, "a:bodyPr");
		var autofit = bodyPr ? X.child(bodyPr, "a:normAutofit") : null;
		var scale = autofit && autofit.a.fontScale ? +autofit.a.fontScale / 100000 : 1;
		var out = [];
		var counters = {};
		var paras = X.children(inTxBody, "a:p");
		for (var i = 0; i < paras.length; i++) {
			var p = paras[i];
			var pPr = X.child(p, "a:pPr");
			var lvl = pPr && pPr.a.lvl ? +pPr.a.lvl + 1 : 1;
			var props = {size: 18};
			if (styleKey === "p:otherStyle" && !inPh) {
				props.bullet = "none";
			}
			for (var l = 0; l < lists.length; l++) {
				props = levelProps(lists[l] ? X.child(lists[l], "a:lvl" + lvl + "pPr") : null, props, theme);
			}
			props = levelProps(pPr, props, theme);
			var css = {};
			if (props.align) {
				css.textAlign = props.align;
			}
			if (props.lineHeight) {
				css.lineHeight = String(Math.round(props.lineHeight * 1.2 * 100) / 100);
			}
			var base = runCss(props, scale);
			var runs = [];
			var kids = X.children(p);
			for (var k = 0; k < kids.length; k++) {
				var r = kids[k];
				if (r.n === "a:r" || r.n === "a:fld") {
					var rp = runProps(X.child(r, "a:rPr"), {}, theme);
					var t = X.textOf(X.child(r, "a:t"));
					if (!t) {
						continue;
					}
					var rcss = runCss(rp, scale);
					var link = X.child(X.child(r, "a:rPr"), "a:hlinkClick");
					var run = {t: "text", v: t};
					if (!isEmptyObj(rcss)) {
						run.s = rcss;
					}
					if (link && inCtx.rels[link.a["r:id"]] && inCtx.rels[link.a["r:id"]].external) {
						runs.push({t: "a", href: inCtx.rels[link.a["r:id"]].target, r: [run]});
					} else {
						runs.push(run);
					}
				} else if (r.n === "a:br") {
					runs.push({t: "br"});
				}
			}
			var block = {t: "p", h: 0, s: merge(css, base), r: runs};
			// Bullets only on paragraphs with words.
			if (runs.length && props.bullet && props.bullet !== "none") {
				if (props.bullet === "auto") {
					var key = lvl + ":" + (props.buScheme || "");
					counters[key] = counters[key] ? counters[key] + 1 : props.buStart || 1;
					block.bu = numberLabel(counters[key], props.buScheme);
				} else {
					var ch = props.buChar || "";
					// Symbol-font bullets are private glyphs (U+E000-U+F8FF, or "§", "l", "Ø" in Wingdings): draw a plain one.
					block.bu = ch.length === 1 && !/[\ue000-\uf8ff]/.test(ch) && !/^[l\u00a7\u00d8\u00fc\u00a8q]$/.test(ch) &&
						!/symbol|wingdings|starbats|webdings/i.test(props.buFont || "") ? ch : "\u2022";
				}
				if (props.buColor) {
					block.buColor = props.buColor;
				}
			} else if (!runs.length) {
				counters = {};
			}
			if (props.marL) {
				block.s.paddingLeft = props.marL + "px";
				if (props.indent) {
					block.s.textIndent = props.indent + "px";
				}
			}
			out.push(block);
		}
		return out;
	};

	Deck.prototype.table = function(inTbl, inTheme, inCtx) {
		var cols = X.children(X.child(inTbl, "a:tblGrid"), "a:gridCol").map(function(c) {
			return px(c.a.w) + "px";
		});
		var rows = [];
		var trs = X.children(inTbl, "a:tr");
		for (var r = 0; r < trs.length; r++) {
			var cells = [];
			var tcs = X.children(trs[r], "a:tc");
			for (var c = 0; c < tcs.length; c++) {
				var tc = tcs[c];
				if (tc.a.hMerge === "1" || tc.a.vMerge === "1") {
					continue;
				}
				var tcPr = X.child(tc, "a:tcPr");
				var css = {border: "1px solid #888888", padding: "4px 6px", verticalAlign: "top"};
				var f = fillOf(tcPr, inTheme);
				if (f && f !== "none") {
					css.backgroundColor = f;
				}
				var txBody = X.child(tc, "a:txBody");
				cells.push({cs: +(tc.a.gridSpan || 1), rs: +(tc.a.rowSpan || 1), s: css,
					b: txBody ? this.text(txBody, null, null, null, inCtx) : []});
			}
			rows.push({header: r === 0, s: {}, cells: cells});
		}
		return {t: "table", s: {borderCollapse: "collapse", width: "100%"}, cols: cols, rows: rows};
	};

	function merge(inA, inB) {
		var out = {}, k;
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

	function isEmptyObj(inObj) {
		for (var k in inObj) {
			if (inObj.hasOwnProperty(k)) {
				return false;
			}
		}
		return true;
	}

	function background(inTree, inCtx, inDeck) {
		var bg = X.path(inTree, "p:cSld", "p:bg");
		if (!bg) {
			return null;
		}
		var bgPr = X.child(bg, "p:bgPr");
		if (bgPr) {
			var blip = X.findAll(X.child(bgPr, "a:blipFill"), "a:blip")[0];
			if (blip && inCtx.rels[blip.a["r:embed"]] && !inCtx.rels[blip.a["r:embed"]].external) {
				return {image: inDeck.useImage(inCtx.rels[blip.a["r:embed"]].target)};
			}
			var f = fillOf(bgPr, inCtx.theme);
			return f && f !== "none" ? {color: f} : null;
		}
		var ref = X.child(bg, "p:bgRef");
		if (ref) {
			var c = colorOf(ref, inCtx.theme);
			return c ? {color: c} : null;
		}
		return null;
	}

	/**
		Parses a whole deck. inFiles maps zip paths to the text of every XML
		and .rels part (pictures are not needed here).
	*/
	function parse(inFiles) {
		var deck = new Deck(inFiles);
		var presPath = "ppt/presentation.xml";
		var pres = deck.xml(presPath);
		if (!pres) {
			throw new Error("PPTX: no ppt/presentation.xml");
		}
		var size = X.child(pres, "p:sldSz");
		var width = size ? px(size.a.cx) : 960, height = size ? px(size.a.cy) : 540;
		var presRels = rels(inFiles, presPath);
		var ids = X.children(X.child(pres, "p:sldIdLst"), "p:sldId");
		var slides = [];
		for (var i = 0; i < ids.length; i++) {
			var rel = presRels[ids[i].a["r:id"]];
			if (!rel || !inFiles[rel.target]) {
				continue;
			}
			var slidePath = rel.target;
			var slide = deck.xml(slidePath);
			var slideRels = rels(inFiles, slidePath);
			var layoutRel = relOfType(slideRels, "slideLayout");
			var layoutPath = layoutRel ? layoutRel.target : null;
			var layout = layoutPath ? deck.xml(layoutPath) : null;
			var layoutRels = layoutPath ? rels(inFiles, layoutPath) : {};
			var masterRel = relOfType(layoutRels, "slideMaster");
			var masterPath = masterRel ? masterRel.target : null;
			var master = masterPath ? deck.xml(masterPath) : null;
			var masterRels = masterPath ? rels(inFiles, masterPath) : {};
			var themeRel = relOfType(masterRels, "theme");
			var clrMap = master ? X.child(master, "p:clrMap") : null;
			var theme = new Theme(themeRel ? inFiles[themeRel.target] : null, clrMap ? clrMap.a : null);
			var txStyles = master ? X.child(master, "p:txStyles") : null;
			var layoutTree = layout ? X.path(layout, "p:cSld", "p:spTree") : null;
			var masterTree = master ? X.path(master, "p:cSld", "p:spTree") : null;

			var items = [];
			var base = {theme: theme, txStyles: txStyles, layoutTree: layoutTree, masterTree: masterTree};
			var showMaster = !(slide.a.showMasterSp === "0" || slide.a.showMasterSp === "false");
			if (showMaster && master && !(layout && (layout.a.showMasterSp === "0" || layout.a.showMasterSp === "false"))) {
				deck.shapes(masterTree, merge(base, {rels: masterRels, layoutTree: null, masterTree: null}), items, true);
			}
			if (showMaster && layout) {
				deck.shapes(layoutTree, merge(base, {rels: layoutRels, layoutTree: null}), items, true);
			}
			var ctx = merge(base, {rels: slideRels});
			deck.shapes(X.path(slide, "p:cSld", "p:spTree"), ctx, items, false);

			var bg = background(slide, ctx, deck) ||
				(layout ? background(layout, merge(base, {rels: layoutRels}), deck) : null) ||
				(master ? background(master, merge(base, {rels: masterRels}), deck) : null) || {};
			var hidden = slide.a.show === "0";
			slides.push({width: width, height: height, bg: bg.color || "#ffffff", bgImage: bg.image || "", items: items,
				hidden: hidden, name: ""});
		}
		return {kind: "pptx", slides: slides, images: deck.images, width: width, height: height};
	}

	//* Title, author and dates from docProps/core.xml (any OOXML package).
	function coreProps(inXml) {
		var out = {};
		if (!inXml) {
			return out;
		}
		var core;
		try {
			core = X.parse(inXml);
		} catch (e) {
			return out;
		}
		var fields = {"dc:title": "title", "dc:subject": "subject", "dc:creator": "author", "cp:lastModifiedBy": "modifiedBy",
			"dcterms:created": "created", "dcterms:modified": "modified", "dc:description": "description"};
		var kids = X.children(core);
		for (var i = 0; i < kids.length; i++) {
			if (fields[kids[i].n]) {
				var v = X.textOf(kids[i]).trim();
				if (v) {
					out[fields[kids[i].n]] = v;
				}
			}
		}
		return out;
	}

	var DocPptx = {
		parse: parse,
		coreProps: coreProps,
		resolvePath: resolvePath,
		colorOf: colorOf,
		Theme: Theme,
		numberLabel: numberLabel
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocPptx;
	} else {
		root.DocPptx = DocPptx;
	}
})(this);
