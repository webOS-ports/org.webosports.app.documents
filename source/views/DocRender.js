/**
	Turns the engine's document model (see source/core/Odf.js) into DOM.
	Everything is built with createElement and text nodes; styles come only
	from the model's whitelisted CSS. Links get no href: a tap on one is
	handled by the viewer (data-href), so the page never navigates.

		DocRender.blocks(parentNode, blocks, {images: {zipPath: blobUrl}});
		DocRender.slides(parentNode, deck.slides, ctx);   // one .doc-slide per slide
		DocRender.neutraliseLinks(node);                  // for docx-preview's output
*/
var DocRender = {
	apply: function(inNode, inCss) {
		if (!inCss) {
			return;
		}
		for (var k in inCss) {
			if (inCss.hasOwnProperty(k)) {
				try {
					inNode.style[k] = inCss[k];
				} catch (e) {}
			}
		}
	},

	el: function(inTag, inClass, inParent) {
		var n = document.createElement(inTag);
		if (inClass) {
			n.className = inClass;
		}
		if (inParent) {
			inParent.appendChild(n);
		}
		return n;
	},

	/**
		Renders blocks a slice at a time so a long document never blocks the
		page: the first slice at once, then about inBudgetMs of work per turn
		of the event loop. Each slice is its own container with
		content-visibility: auto, so the engine only lays out the slices
		near the screen. Returns {cancel()}; calls inDone() at the end.
	*/
	progressive: function(inParent, inBlocks, inCtx, inDone, inBudgetMs) {
		var self = this;
		var i = 0, cancelled = false;
		var budget = inBudgetMs || 30;
		var slice = function(inFirst) {
			if (cancelled) {
				return;
			}
			var chunk = self.el("div", "doc-chunk", inParent);
			var start = Date.now();
			var count = 0;
			while (i < inBlocks.length && (count < 8 || Date.now() - start < (inFirst ? budget * 2 : budget))) {
				self.block(chunk, inBlocks[i], inCtx);
				i++;
				count++;
			}
			if (i < inBlocks.length) {
				setTimeout(slice, 0);
			} else if (inDone) {
				inDone();
			}
		};
		slice(true);
		return {
			cancel: function() {
				cancelled = true;
			},
			done: function() {
				return i >= inBlocks.length;
			},
			progress: function() {
				return inBlocks.length ? i / inBlocks.length : 1;
			}
		};
	},

	blocks: function(inParent, inBlocks, inCtx) {
		for (var i = 0; i < (inBlocks || []).length; i++) {
			this.block(inParent, inBlocks[i], inCtx);
		}
	},

	block: function(inParent, inBlock, inCtx) {
		var b = inBlock, n;
		switch (b.t) {
		case "p":
			n = this.el(b.h ? "h" + b.h : "p", "doc-p", inParent);
			this.apply(n, b.s);
			if (b.bu) {
				var bullet = this.el("span", "doc-bullet", n);
				bullet.appendChild(document.createTextNode(b.bu));
				// In the hanging indent when the paragraph has one, else just before the text.
				var hang = b.s && /^-\d/.test(b.s.textIndent || "") ? b.s.textIndent.slice(1) : "";
				bullet.style.width = hang || "auto";
				bullet.style.paddingRight = hang ? "0" : "0.4em";
				if (b.buColor) {
					bullet.style.color = b.buColor;
				}
			}
			this.inlines(n, b.r, inCtx);
			if (!b.r || !b.r.length) {
				// An empty paragraph keeps its line, as in the editor.
				n.appendChild(document.createElement("br"));
			}
			break;
		case "list":
			n = this.el(b.o ? "ol" : "ul", "doc-list", inParent);
			if (b.o) {
				if (b.type) {
					n.style.listStyleType = b.type;
				}
				if (b.start && b.start !== 1) {
					n.setAttribute("start", String(b.start));
				}
			}
			for (var i = 0; i < b.items.length; i++) {
				var li = this.el("li", "", n);
				this.blocks(li, b.items[i], inCtx);
			}
			break;
		case "table":
			this.table(inParent, b, inCtx);
			break;
		case "frame":
			n = this.el("div", "doc-frame", inParent);
			this.apply(n, b.s);
			this.blocks(n, b.b, inCtx);
			break;
		case "pagebreak":
			this.el("div", "doc-pagebreak", inParent);
			break;
		case "notes":
			n = this.el("div", "doc-notes", inParent);
			for (var k = 0; k < b.items.length; k++) {
				var note = this.el("div", "doc-note", n);
				var label = this.el("span", "doc-note-label", note);
				label.appendChild(document.createTextNode(b.items[k].label));
				this.blocks(note, b.items[k].b, inCtx);
			}
			break;
		}
	},

	table: function(inParent, inTable, inCtx) {
		var wrap = this.el("div", "doc-table-wrap", inParent);
		var t = this.el("table", "doc-table", wrap);
		this.apply(t, inTable.s);
		if (inTable.cols && inTable.cols.length) {
			var cg = this.el("colgroup", "", t);
			for (var c = 0; c < inTable.cols.length; c++) {
				var col = this.el("col", "", cg);
				if (inTable.cols[c]) {
					col.style.width = inTable.cols[c];
				}
			}
		}
		var body = this.el("tbody", "", t);
		for (var r = 0; r < inTable.rows.length; r++) {
			var row = inTable.rows[r];
			var tr = this.el("tr", row.header ? "doc-th-row" : "", body);
			this.apply(tr, row.s);
			for (var k = 0; k < row.cells.length; k++) {
				var cell = row.cells[k];
				var td = this.el(row.header ? "th" : "td", "", tr);
				if (cell.cs > 1) {
					td.colSpan = cell.cs;
				}
				if (cell.rs > 1) {
					td.rowSpan = cell.rs;
				}
				this.apply(td, cell.s);
				this.blocks(td, cell.b, inCtx);
			}
		}
	},

	inlines: function(inParent, inRuns, inCtx) {
		for (var i = 0; i < (inRuns || []).length; i++) {
			var r = inRuns[i], n;
			switch (r.t) {
			case "text":
				if (r.s) {
					n = this.el("span", "", inParent);
					this.apply(n, r.s);
					n.appendChild(document.createTextNode(r.v));
				} else {
					inParent.appendChild(document.createTextNode(r.v));
				}
				break;
			case "br":
				inParent.appendChild(document.createElement("br"));
				break;
			case "tab":
				n = this.el("span", "doc-tab", inParent);
				n.appendChild(document.createTextNode("\t"));
				break;
			case "a":
				n = this.el("a", "doc-link", inParent);
				n.setAttribute("data-href", r.href || "");
				this.inlines(n, r.r, inCtx);
				break;
			case "img":
				this.image(inParent, r, inCtx);
				break;
			case "anchor":
				// A bookmark: the target of in-document links ("#name").
				n = this.el("span", "doc-anchor", inParent);
				n.setAttribute("data-anchor", r.id);
				break;
			case "note":
				n = this.el("sup", "doc-note-ref", inParent);
				n.appendChild(document.createTextNode(r.label));
				break;
			case "frame":
				n = this.el("span", "doc-frame doc-inline-frame", inParent);
				this.apply(n, r.s);
				this.blocks(n, r.b, inCtx);
				break;
			}
		}
	},

	image: function(inParent, inImg, inCtx) {
		var url = inCtx && inCtx.images && inCtx.images[inImg.src];
		var n;
		if (url) {
			n = this.el("img", "doc-img", inParent);
			n.src = url;
			n.alt = inImg.alt || "";
		} else {
			// No picture (missing, or a format the engine cannot show, such as WMF): keep its place.
			n = this.el("span", "doc-img doc-img-missing", inParent);
			n.appendChild(document.createTextNode(inImg.alt || "picture"));
		}
		if (inImg.w) {
			n.style.width = inImg.w;
		}
		if (inImg.h) {
			n.style.height = inImg.h;
		}
	},

	/**
		Slides: each a box of the slide's size in CSS pixels, its items
		placed absolutely. The viewer scales the boxes to the screen.
	*/
	slides: function(inParent, inSlides, inCtx) {
		var out = [];
		for (var i = 0; i < inSlides.length; i++) {
			var s = inSlides[i];
			var frame = this.el("div", "doc-slide-frame" + (s.hidden ? " doc-slide-hidden" : ""), inParent);
			var slide = this.el("div", "doc-slide", frame);
			slide.style.width = s.width + "px";
			slide.style.height = s.height + "px";
			slide.style.backgroundColor = s.bg || "#ffffff";
			if (s.bgImage && inCtx.images[s.bgImage]) {
				var bg = this.el("img", "doc-slide-bg", slide);
				bg.src = inCtx.images[s.bgImage];
			}
			for (var k = 0; k < s.items.length; k++) {
				this.slideItem(slide, s.items[k], inCtx);
			}
			var num = this.el("div", "doc-slide-number", frame);
			num.appendChild(document.createTextNode(String(i + 1) + (s.hidden ? " (hidden)" : "")));
			out.push({frame: frame, slide: slide, width: s.width, height: s.height});
		}
		return out;
	},

	slideItem: function(inSlide, inItem, inCtx) {
		var it = inItem;
		var n;
		if (it.kind === "line") {
			n = this.el("div", "doc-slide-line", inSlide);
			var w = Math.max(it.w, 1), h = Math.max(it.h, 1);
			var len = Math.sqrt(w * w + h * h);
			var angle = Math.atan2(it.h, it.w) * 180 / Math.PI;
			if (it.flipH !== it.flipV) {
				angle = -angle;
			}
			n.style.left = (it.x + it.w / 2 - len / 2) + "px";
			n.style.top = (it.y + it.h / 2) + "px";
			n.style.width = len + "px";
			n.style.borderTopColor = it.stroke || "#000000";
			n.style.transform = "rotate(" + angle + "deg)";
			return;
		}
		n = this.el("div", "doc-slide-item doc-slide-" + it.kind, inSlide);
		n.style.left = it.x + "px";
		n.style.top = it.y + "px";
		n.style.width = it.w + "px";
		n.style.height = it.h + "px";
		if (it.fill) {
			n.style.backgroundColor = it.fill;
		}
		if (it.stroke) {
			n.style.border = "1px solid " + it.stroke;
		}
		if (it.round) {
			n.style.borderRadius = "50%";
		} else if (it.radius) {
			n.style.borderRadius = "8px";
		}
		if (it.kind === "image") {
			var url = inCtx.images[it.src];
			if (url) {
				var img = this.el("img", "", n);
				img.src = url;
			} else {
				n.className += " doc-img-missing";
			}
			return;
		}
		var text = this.el("div", "doc-slide-text", n);
		this.apply(text, it.s);
		if (it.pad) {
			text.style.padding = it.pad.join("px ") + "px";
		}
		if (it.valign === "middle" || it.valign === "bottom") {
			n.style.display = "flex";
			n.style.flexDirection = "column";
			n.style.justifyContent = it.valign === "middle" ? "center" : "flex-end";
		}
		this.blocks(text, it.b, inCtx);
	},

	//* Moves every href to data-href, so no tap leaves the app; returns how many.
	neutraliseLinks: function(inNode) {
		var links = inNode.querySelectorAll("a[href]");
		for (var i = 0; i < links.length; i++) {
			var a = links[i];
			a.setAttribute("data-href", a.getAttribute("href"));
			a.removeAttribute("href");
			a.className += " doc-link";
		}
		// Forms, scripts and frames have no place in a document.
		var bad = inNode.querySelectorAll("script, iframe, object, embed, form");
		for (var j = 0; j < bad.length; j++) {
			bad[j].parentNode.removeChild(bad[j]);
		}
		return links.length;
	}
};
