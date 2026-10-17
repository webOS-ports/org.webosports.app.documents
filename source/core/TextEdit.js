/**
	Edits the text of an OpenDocument paragraph in content.xml, keeping
	what was not changed exactly as it was: the formatting of the words
	around the change, pictures, notes, fields and bookmarks.

		var doc = DocTextEdit.session(contentXml, track);    // track from DocOdf.text(files, {track})
		var r = doc.edit(id, "The paragraph's new text");     // r.block: the paragraph again, for the page
		doc.xml                                               // content.xml as it is now

	The text is the paragraph's as DocOdf.tokens() gives it: tabs as "\t",
	line breaks as "\n", objects (pictures, notes, frames) as U+FFFC.

	Only the changed parts are written: the old and the new text are
	compared character by character, and each run of text keeps what it
	had that was not changed. Text that replaces text takes its
	formatting; added text goes into the run of text before it (it takes
	that run's formatting), at the start of a paragraph into the first
	run. An object or a field that loses any of its text is
	removed.
*/
(function(root) {
	var X = root.DocXml || (typeof require === "function" ? require("./Xml.js") : null);
	var Odf = root.DocOdf || (typeof require === "function" ? require("./Odf.js") : null);

	//* The prefix (with ":") the document gave the text namespace.
	function textPrefix(inScope) {
		for (var k in inScope) {
			// The scope's prototype chain holds the outer declarations: for..in walks them too.
			if (inScope[k] === "text") {
				return k ? k + ":" : "";
			}
		}
		return "text:";
	}

	/**
		inText as ODF paragraph content: spaces that would collapse become
		text:s, tabs text:tab, line breaks text:line-break. Characters XML
		cannot hold, and U+FFFC (an object), are dropped.
	*/
	function encode(inText, inPrefix) {
		var t = inPrefix;
		var s = String(inText).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￼￾￿\r]/g, "");
		var out = "";
		var i = 0;
		while (i < s.length) {
			var ch = s.charAt(i);
			if (ch === "\t") {
				out += "<" + t + "tab/>";
				i++;
			} else if (ch === "\n") {
				out += "<" + t + "line-break/>";
				i++;
			} else if (ch === " ") {
				var j = i;
				while (j < s.length && s.charAt(j) === " ") {
					j++;
				}
				var n = j - i;
				// One space after a letter is kept as it is; ODF collapses any other.
				var literal = i > 0 && !/[ \t\n]/.test(s.charAt(i - 1));
				if (literal) {
					out += " ";
					n--;
				}
				if (n > 0) {
					out += "<" + t + "s" + (n > 1 ? " " + t + "c=\"" + n + "\"" : "") + "/>";
				}
				i = j;
			} else {
				var k = i;
				while (k < s.length && !/[ \t\n]/.test(s.charAt(k))) {
					k++;
				}
				out += X.escape(s.slice(i, k));
				i = k;
			}
		}
		return out;
	}

	/**
		Where two texts differ: [{s, e, ins}], old[s, e) replaced by ins, in
		order (Myers' diff, by character). Long texts with many changes fall
		back to one change between the shared start and end.
	*/
	function diff(inOld, inNew) {
		var a = inOld, b = inNew;
		var pre = 0, max0 = Math.min(a.length, b.length);
		while (pre < max0 && a.charCodeAt(pre) === b.charCodeAt(pre)) {
			pre++;
		}
		var suf = 0;
		while (suf < max0 - pre && a.charCodeAt(a.length - 1 - suf) === b.charCodeAt(b.length - 1 - suf)) {
			suf++;
		}
		var one = [{s: pre, e: a.length - suf, ins: b.slice(pre, b.length - suf)}];
		a = a.slice(pre, a.length - suf);
		b = b.slice(pre, b.length - suf);
		var n = a.length, m = b.length;
		if (!n || !m || n + m > 20000) {
			return one;
		}
		var limit = Math.min(n + m, 500);
		var off = limit + 1;
		var v = [];
		v[off + 1] = 0;
		var trace = [];
		var found = -1;
		for (var d = 0; d <= limit && found < 0; d++) {
			trace.push(v.slice(off - d - 1, off + d + 2));
			for (var k = -d; k <= d; k += 2) {
				var x = k === -d || k !== d && v[off + k - 1] < v[off + k + 1] ? v[off + k + 1] : v[off + k - 1] + 1;
				var y = x - k;
				while (x < n && y < m && a.charCodeAt(x) === b.charCodeAt(y)) {
					x++;
					y++;
				}
				v[off + k] = x;
				if (x >= n && y >= m) {
					found = d;
					break;
				}
			}
		}
		if (found < 0) {
			return one;
		}
		// Back from the end: each step is one character removed or added.
		var steps = [];
		var cx = n, cy = m;
		for (var dd = found; dd > 0; dd--) {
			var tv = trace[dd];
			var get = function(inK) {
				return tv[inK + dd + 1];
			};
			var kk = cx - cy;
			var prevK = kk === -dd || kk !== dd && get(kk - 1) < get(kk + 1) ? kk + 1 : kk - 1;
			var px = get(prevK), py = px - prevK;
			while (cx > px && cy > py) {
				cx--;
				cy--;
			}
			if (cx === px) {
				steps.push({t: "i", x: cx, ch: b.charAt(py)});
			} else {
				steps.push({t: "d", x: px});
			}
			cx = px;
			cy = py;
		}
		steps.reverse();
		var hunks = [];
		steps.forEach(function(st) {
			var last = hunks[hunks.length - 1];
			var x = st.x + pre;
			if (last && (st.t === "d" ? last.e === x : last.e === x)) {
				if (st.t === "d") {
					last.e++;
				} else {
					last.ins += st.ch;
				}
			} else {
				hunks.push(st.t === "d" ? {s: x, e: x + 1, ins: ""} : {s: x, e: x, ins: st.ch});
			}
		});
		return hunks;
	}

	/**
		How paragraphs are written, per format. A dialect gives:
			tokens(el), plainText(tokens)   the paragraph as text (see DocOdf.tokens)
			text(tk, text, el)              XML for a text token's range (tk.s..tk.e) with new text
			keep(tk, text, xml)             XML for a space run or break (its element) that kept part of its text
			insert(xml, el, toks, at, text) {s, e, v}: new text where no run of text takes it
	*/
	var odfDialect = {
		tokens: function(inEl) {
			return Odf.tokens(inEl);
		},
		plainText: function(inToks) {
			return Odf.plainText(inToks);
		},
		text: function(inTk, inText, inEl) {
			return encode(inText, textPrefix(inEl.ns || {}));
		},
		keep: function(inTk, inText, inXml, inEl) {
			return encode(inText, textPrefix(inEl.ns || {}));
		},
		// Next to objects only, or an empty paragraph: the text goes in as it is, without a run of its own.
		insert: function(inXml, inEl, inToks, inAt, inText) {
			var raw = encode(inText, textPrefix(inEl.ns || {}));
			var pos = inAt;
			if (pos === null && inEl.ci === inEl.e) {
				// <text:p .../> becomes <text:p ...>text</text:p>.
				var tag = inXml.slice(inEl.p, inEl.ci).replace(/\s*\/>$/, ">");
				return {s: inEl.p, e: inEl.e, v: tag + raw + "</" + inEl.raw + ">"};
			}
			if (pos === null) {
				var last = inToks[inToks.length - 1];
				pos = last ? (last.k === "t" ? last.e : last.el.e) : inEl.ce;
			}
			return {s: pos, e: pos, v: raw};
		}
	};

	/**
		DrawingML (.pptx slides): text is in runs, a:r/a:t, kept as typed (no
		collapsing); a line break is an a:br between runs, so a new line in a
		run ends it and starts another with the same properties.
	*/
	function drawingPrefix(inEl) {
		var i = inEl.raw.indexOf(":");
		return i < 0 ? "" : inEl.raw.slice(0, i + 1);
	}

	function drawingRuns(inText, inPrefix, inRPr) {
		var a = inPrefix;
		var clean = String(inText).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffc\ufffe\uffff\r]/g, "");
		return clean.split("\n").map(function(inPart, i) {
			return (i ? "<" + a + "br/>" : "") + (inPart ? "<" + a + "r>" + inRPr + "<" + a + "t>" + X.escape(inPart) + "</" + a + "t></" + a + "r>" : "");
		}).join("");
	}

	var drawingDialect = {
		tokens: function(inEl) {
			return Pptx().tokens(inEl);
		},
		plainText: function(inToks) {
			return Pptx().plainText(inToks);
		},
		// Within a:t: the first line stays in this run; each line after a break gets a run like it.
		text: function(inTk, inText, inEl, inXml) {
			var a = drawingPrefix(inTk.run);
			var rPr = inTk.rPr ? inXml.slice(inTk.rPr.p, inTk.rPr.e) : "";
			var lines = String(inText).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffc\ufffe\uffff\r]/g, "").split("\n");
			var out = X.escape(lines[0]);
			for (var i = 1; i < lines.length; i++) {
				out += "</" + a + "t></" + a + "r><" + a + "br/><" + a + "r>" + rPr + "<" + a + "t>" + X.escape(lines[i]);
			}
			return out;
		},
		keep: function(inTk, inText, inXml) {
			return inText === inTk.v ? inXml.slice(inTk.el.p, inTk.el.e) : "";
		},
		// New runs before the token at inAt (or the paragraph's end properties), formatted as the paragraph's end.
		insert: function(inXml, inEl, inToks, inAt, inText) {
			var a = drawingPrefix(inEl);
			var end = X.child(inEl, "a:endParaRPr");
			var rPr = end ? inXml.slice(end.p, end.e).replace(/^<([\w]+:)?endParaRPr\b/, "<$1rPr").replace(/<\/([\w]+:)?endParaRPr>$/, "</$1rPr>") : "";
			var runs = drawingRuns(inText, a, rPr);
			if (inEl.ci === inEl.e) {
				var tag = inXml.slice(inEl.p, inEl.ci).replace(/\s*\/>$/, ">");
				return {s: inEl.p, e: inEl.e, v: tag + runs + "</" + inEl.raw + ">"};
			}
			var pos = null;
			for (var i = 0; i < inToks.length && pos === null; i++) {
				if (inToks[i].a >= inAt) {
					pos = inToks[i].k === "t" ? inToks[i].run.p : inToks[i].el.p;
				}
			}
			if (pos === null) {
				pos = end ? end.p : inEl.ce;
			}
			return {s: pos, e: pos, v: runs};
		}
	};

	//* Pptx.js, when it is needed (a text document's session never loads it).
	function Pptx() {
		return root.DocPptx || (typeof require === "function" ? require("./Pptx.js") : null);
	}

	/**
		The splices that turn paragraph inEl (parsed with offsets, in inXml)
		into one with inText. Returns null when the text is the same.
		inDialect: odfDialect (the default) or drawingDialect.
	*/
	function splices(inXml, inEl, inText, inDialect) {
		var dialect = inDialect || odfDialect;
		var toks = dialect.tokens(inEl);
		var old = dialect.plainText(toks);
		var neu = String(inText);
		if (old === neu) {
			return null;
		}
		var hunks = diff(old, neu);
		// Positions of the tokens in the text.
		var at = 0;
		toks.forEach(function(tk) {
			tk.a = at;
			at += tk.v.length;
			tk.b = at;
		});
		// Where each hunk's new text goes: the text run it replaces text in, else the one it continues, else
		// the one it starts.
		hunks.forEach(function(h) {
			h.target = null;
			if (!h.ins) {
				return;
			}
			for (var r = 0; r < toks.length && !h.target && h.s < h.e; r++) {
				if (toks[r].k === "t" && toks[r].a <= h.s && h.s < toks[r].b) {
					h.target = toks[r];
				}
			}
			for (var i = 0; i < toks.length && !h.target; i++) {
				if (toks[i].k === "t" && toks[i].a < h.s && h.s <= toks[i].b) {
					h.target = toks[i];
				}
			}
			for (var j = 0; j < toks.length && !h.target; j++) {
				if (toks[j].k === "t" && toks[j].a === h.s) {
					h.target = toks[j];
				}
			}
		});
		var out = [];
		toks.forEach(function(tk) {
			var mine = hunks.filter(function(h) {
				return h.target === tk || h.s < h.e && h.s < tk.b && h.e > tk.a && tk.b > tk.a;
			});
			if (!mine.length) {
				return;
			}
			if (tk.k === "o") {
				// An object or a field that lost any of its text goes.
				out.push({s: tk.el.p, e: tk.el.e, v: ""});
				return;
			}
			var nv = "";
			var pos = tk.a;
			mine.forEach(function(h) {
				var s0 = Math.max(tk.a, Math.min(tk.b, h.s));
				nv += tk.v.slice(pos - tk.a, Math.max(pos, s0) - tk.a);
				if (h.target === tk) {
					nv += h.ins;
				}
				pos = Math.max(pos, Math.min(tk.b, h.e));
			});
			nv += tk.v.slice(pos - tk.a);
			out.push(tk.k === "t" ? {s: tk.s, e: tk.e, v: dialect.text(tk, nv, inEl, inXml)} :
				{s: tk.el.p, e: tk.el.e, v: dialect.keep(tk, nv, inXml, inEl)});
		});
		hunks.forEach(function(h) {
			if (!h.ins || h.target) {
				return;
			}
			if (dialect === odfDialect) {
				var pos = null;
				for (var i = 0; i < toks.length && pos === null; i++) {
					if (toks[i].a >= h.s && toks[i].b > toks[i].a) {
						pos = toks[i].k === "t" ? toks[i].s : toks[i].el.p;
					}
				}
				out.push(dialect.insert(inXml, inEl, toks, pos, h.ins));
			} else {
				out.push(dialect.insert(inXml, inEl, toks, h.s, h.ins));
			}
		});
		return out;
	}

	//* Moves every offset at or after inFrom by inDelta (the text after an edit moved).
	function shift(inEl, inFrom, inDelta) {
		if (inEl.e < inFrom) {
			// Nothing in it moved.
			return;
		}
		if (inEl.p >= inFrom) {
			inEl.p += inDelta;
		}
		if (inEl.ci >= inFrom) {
			inEl.ci += inDelta;
		}
		if (inEl.ce >= inFrom) {
			inEl.ce += inDelta;
		}
		inEl.e += inDelta;
		for (var i = 0; i < inEl.c.length; i++) {
			var c = inEl.c[i];
			if (typeof c === "object") {
				shift(c, inFrom, inDelta);
			} else if (inEl.co && inEl.co[i]) {
				var r = inEl.co[i];
				if (r[0] >= inFrom) {
					r[0] += inDelta;
				}
				if (r[1] >= inFrom) {
					r[1] += inDelta;
				}
			}
		}
	}

	//* The paragraphs below inEl (not inEl itself), in document order.
	function paragraphsIn(inEl) {
		var out = [];
		var walk = function(inNode) {
			for (var i = 0; i < inNode.c.length; i++) {
				var c = inNode.c[i];
				if (typeof c === "object") {
					if (c.n === "text:p" || c.n === "text:h") {
						out.push(c);
					}
					walk(c);
				}
			}
		};
		walk(inEl);
		return out;
	}

	//* The element and the index in its parent of inTarget, below inEl.
	function parentOf(inEl, inTarget) {
		for (var i = 0; i < inEl.c.length; i++) {
			var c = inEl.c[i];
			if (c === inTarget) {
				return {parent: inEl, index: i};
			}
			if (typeof c === "object" && c.p <= inTarget.p && c.e >= inTarget.e) {
				var r = parentOf(c, inTarget);
				if (r) {
					return r;
				}
			}
		}
		return null;
	}

	function Session(inXml, inTrack) {
		this.xml = inXml;
		this.root = inTrack.content;
		this.paras = inTrack.paras;
		this.styles = inTrack.styles;
		this.edited = false;
	}

	/**
		Sets paragraph inId's text. Returns {block, images} (the paragraph
		as DocOdf builds it, to show again), or null when nothing changed.
	*/
	Session.prototype.edit = function(inId, inText) {
		var el = this.paras[inId];
		if (!el) {
			throw new Error("This paragraph can no longer be edited.");
		}
		var list = splices(this.xml, el, inText);
		if (!list) {
			return null;
		}
		var start = el.p, oldEnd = el.e;
		var piece = X.splice(this.xml.slice(start, oldEnd), list.map(function(x) {
			return {s: x.s - start, e: x.e - start, v: x.v};
		}));
		var found = parentOf(this.root, el);
		if (!found) {
			throw new Error("This paragraph can no longer be edited.");
		}
		var oldXml = this.xml;
		var oldInside = paragraphsIn(el);
		this.xml = this.xml.slice(0, start) + piece + this.xml.slice(oldEnd);
		// What follows moved; the new paragraph has its own offsets already, so it goes in after.
		shift(this.root, oldEnd, piece.length - (oldEnd - start));
		var neu = X.parse(piece, {offsets: true, scope: el.ns, base: start});
		found.parent.c[found.index] = neu;
		this.paras[inId] = neu;
		// Paragraphs inside this one (in notes and text boxes) were parsed again with it. An edit only
		// removes objects, so the new ones are the old ones less those that went: matched by their text.
		var newInside = paragraphsIn(neu);
		var j = 0;
		for (var i = 0; i < oldInside.length; i++) {
			var o = oldInside[i];
			var id = this.paras.indexOf(o);
			var n = newInside[j];
			if (n && oldXml.slice(o.p, o.e) === this.xml.slice(n.p, n.e)) {
				j++;
				if (id >= 0) {
					this.paras[id] = n;
				}
			} else if (id >= 0) {
				this.paras[id] = null;
			}
		}
		this.edited = true;
		var paras = this.paras;
		return Odf.paragraphBlock(this.styles, neu, function(inEl) {
			return paras.indexOf(inEl);
		});
	};

	var DocTextEdit = {
		session: function(inXml, inTrack) {
			return new Session(inXml, inTrack);
		},
		encode: encode,
		diff: diff,
		splices: splices,
		odfDialect: odfDialect,
		drawingDialect: drawingDialect
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocTextEdit;
	} else {
		root.DocTextEdit = DocTextEdit;
	}
})(this);
