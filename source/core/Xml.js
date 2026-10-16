/**
	A small, forgiving XML parser for the parts of office documents (ODF and
	Office Open XML) that the viewer reads. No DOM is needed, so the same code
	runs in the app, in a Web Worker and in node for the unit tests.

		var root = DocXml.parse('<office:text xmlns:office="...">...</office:text>');
		root.n        // element name, its prefix made canonical ("office:text")
		root.a        // attributes, by canonical name ({"text:style-name": "P1"})
		root.c        // children: elements and strings (text)

	Prefixes are made canonical from the namespace URI, so a document that
	binds, say, "t:" to the ODF text namespace reads the same as one that uses
	"text:". Unknown namespaces keep the prefix the document gave them.

	Not supported, because office documents do not use them: DTDs with
	internal entities, and namespace prefixes rebound to another URI in a
	nested scope with a different meaning.
*/
(function(root) {
	//* Namespace URI -> canonical prefix.
	var namespaces = {
		"urn:oasis:names:tc:opendocument:xmlns:office:1.0": "office",
		"urn:oasis:names:tc:opendocument:xmlns:style:1.0": "style",
		"urn:oasis:names:tc:opendocument:xmlns:text:1.0": "text",
		"urn:oasis:names:tc:opendocument:xmlns:table:1.0": "table",
		"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0": "draw",
		"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0": "fo",
		"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0": "svg",
		"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0": "presentation",
		"urn:oasis:names:tc:opendocument:xmlns:meta:1.0": "meta",
		"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0": "number",
		"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0": "manifest",
		"http://www.w3.org/1999/xlink": "xlink",
		"http://purl.org/dc/elements/1.1/": "dc",
		"http://purl.org/dc/terms/": "dcterms",
		"http://schemas.openxmlformats.org/drawingml/2006/main": "a",
		"http://schemas.openxmlformats.org/presentationml/2006/main": "p",
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships": "r",
		"http://schemas.openxmlformats.org/package/2006/relationships": "rel",
		"http://schemas.openxmlformats.org/package/2006/metadata/core-properties": "cp",
		"http://schemas.openxmlformats.org/package/2006/content-types": "ct",
		"http://schemas.openxmlformats.org/wordprocessingml/2006/main": "w",
		"http://schemas.openxmlformats.org/spreadsheetml/2006/main": "x"
	};

	var entities = {lt: "<", gt: ">", amp: "&", quot: "\"", apos: "'"};

	function decodeEntities(inText) {
		if (inText.indexOf("&") < 0) {
			return inText;
		}
		return inText.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, function(m, e) {
			if (e.charAt(0) === "#") {
				var code = e.charAt(1) === "x" || e.charAt(1) === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
				if (isNaN(code) || code < 0 || code > 0x10FFFF) {
					return m;
				}
				if (code > 0xFFFF) {
					code -= 0x10000;
					return String.fromCharCode(0xD800 + (code >> 10), 0xDC00 + (code & 0x3FF));
				}
				return String.fromCharCode(code);
			}
			return entities.hasOwnProperty(e) ? entities[e] : m;
		});
	}

	var attrRe = /([^\s=\/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

	//* Splits "prefix:local" and maps the prefix through the scope.
	function canonical(inName, inScope) {
		var i = inName.indexOf(":");
		if (i < 0) {
			var def = inScope[""];
			return def ? def + ":" + inName : inName;
		}
		var prefix = inName.slice(0, i);
		var mapped = inScope[prefix];
		return (mapped || prefix) + inName.slice(i);
	}

	/**
		Parses inText; returns the root element. Throws on input that has no
		root element at all; tolerates stray end tags and unclosed elements
		(they are closed at the end), as a viewer should.
	*/
	function parse(inText) {
		var text = String(inText);
		// A byte order mark is not content.
		if (text.charCodeAt(0) === 0xFEFF) {
			text = text.slice(1);
		}
		var doc = {n: "#document", a: {}, c: []};
		var stack = [doc];
		var scopes = [{}];
		var pos = 0, len = text.length;
		while (pos < len) {
			var lt = text.indexOf("<", pos);
			if (lt < 0) {
				lt = len;
			}
			if (lt > pos) {
				var chunk = text.slice(pos, lt);
				if (stack.length > 1) {
					stack[stack.length - 1].c.push(decodeEntities(chunk));
				}
			}
			if (lt >= len) {
				break;
			}
			var next = text.charAt(lt + 1);
			var end;
			if (next === "?") {
				end = text.indexOf("?>", lt + 2);
				pos = end < 0 ? len : end + 2;
			} else if (next === "!") {
				if (text.substr(lt, 4) === "<!--") {
					end = text.indexOf("-->", lt + 4);
					pos = end < 0 ? len : end + 3;
				} else if (text.substr(lt, 9) === "<![CDATA[") {
					end = text.indexOf("]]>", lt + 9);
					if (stack.length > 1) {
						stack[stack.length - 1].c.push(text.slice(lt + 9, end < 0 ? len : end));
					}
					pos = end < 0 ? len : end + 3;
				} else {
					// <!DOCTYPE ...>, possibly with an internal subset in [...].
					var bracket = text.indexOf("[", lt);
					var gt0 = text.indexOf(">", lt);
					if (bracket >= 0 && bracket < gt0) {
						end = text.indexOf("]>", bracket);
						pos = end < 0 ? len : end + 2;
					} else {
						pos = gt0 < 0 ? len : gt0 + 1;
					}
				}
			} else if (next === "/") {
				end = text.indexOf(">", lt + 2);
				var closeName = text.slice(lt + 2, end < 0 ? len : end).replace(/\s+$/, "");
				pos = end < 0 ? len : end + 1;
				// Pop to the matching element; ignore an end tag nothing opened.
				for (var s = stack.length - 1; s > 0; s--) {
					if (stack[s].raw === closeName) {
						stack.length = s;
						scopes.length = s;
						break;
					}
				}
			} else {
				// Start tag; find its end outside quoted attribute values.
				var q = null, j = lt + 1;
				for (; j < len; j++) {
					var ch = text.charAt(j);
					if (q) {
						if (ch === q) {
							q = null;
						}
					} else if (ch === "\"" || ch === "'") {
						q = ch;
					} else if (ch === ">") {
						break;
					}
				}
				var body = text.slice(lt + 1, j);
				pos = j + 1;
				var selfClosing = body.charAt(body.length - 1) === "/";
				if (selfClosing) {
					body = body.slice(0, -1);
				}
				var m = /^[^\s\/>]+/.exec(body);
				if (!m) {
					continue;
				}
				var raw = m[0];
				var rawAttrs = {};
				var scope = null;
				attrRe.lastIndex = raw.length;
				var am;
				while ((am = attrRe.exec(body))) {
					var an = am[1];
					var av = decodeEntities(am[3] !== undefined ? am[3] : am[4]);
					if (an === "xmlns" || an.indexOf("xmlns:") === 0) {
						if (!scope) {
							scope = Object.create(scopes[scopes.length - 1]);
						}
						var canon = namespaces[av] || (an === "xmlns" ? "" : an.slice(6));
						scope[an === "xmlns" ? "" : an.slice(6)] = canon;
					} else {
						rawAttrs[an] = av;
					}
				}
				scope = scope || scopes[scopes.length - 1];
				var attrs = {};
				for (var k in rawAttrs) {
					if (rawAttrs.hasOwnProperty(k)) {
						// Unprefixed attributes have no namespace (not the default one).
						attrs[k.indexOf(":") < 0 ? k : canonical(k, scope)] = rawAttrs[k];
					}
				}
				var el = {n: canonical(raw, scope), a: attrs, c: [], raw: raw};
				stack[stack.length - 1].c.push(el);
				if (!selfClosing) {
					stack.push(el);
					scopes.push(scope);
				}
			}
		}
		for (var r = 0; r < doc.c.length; r++) {
			if (typeof doc.c[r] === "object") {
				return doc.c[r];
			}
		}
		throw new Error("XML: no root element");
	}

	// --- Helpers for walking the tree ------------------------------------

	//* The first child element named inName (or any element when omitted).
	function child(inEl, inName) {
		if (!inEl) {
			return null;
		}
		for (var i = 0; i < inEl.c.length; i++) {
			var c = inEl.c[i];
			if (typeof c === "object" && (!inName || c.n === inName)) {
				return c;
			}
		}
		return null;
	}

	//* All child elements named inName (all elements when omitted).
	function children(inEl, inName) {
		var out = [];
		if (!inEl) {
			return out;
		}
		for (var i = 0; i < inEl.c.length; i++) {
			var c = inEl.c[i];
			if (typeof c === "object" && (!inName || c.n === inName)) {
				out.push(c);
			}
		}
		return out;
	}

	//* Follows a path of child names: path(root, "office:body", "office:text").
	function path(inEl) {
		var el = inEl;
		for (var i = 1; i < arguments.length && el; i++) {
			el = child(el, arguments[i]);
		}
		return el || null;
	}

	//* Depth-first search for every element named inName below inEl.
	function findAll(inEl, inName, inOut) {
		var out = inOut || [];
		if (!inEl) {
			return out;
		}
		for (var i = 0; i < inEl.c.length; i++) {
			var c = inEl.c[i];
			if (typeof c === "object") {
				if (c.n === inName) {
					out.push(c);
				}
				findAll(c, inName, out);
			}
		}
		return out;
	}

	//* The concatenated text below inEl.
	function textOf(inEl) {
		if (!inEl) {
			return "";
		}
		if (typeof inEl === "string") {
			return inEl;
		}
		var s = "";
		for (var i = 0; i < inEl.c.length; i++) {
			s += textOf(inEl.c[i]);
		}
		return s;
	}

	var DocXml = {
		parse: parse,
		decodeEntities: decodeEntities,
		child: child,
		children: children,
		path: path,
		findAll: findAll,
		textOf: textOf,
		namespaces: namespaces
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocXml;
	} else {
		root.DocXml = DocXml;
	}
})(this);
