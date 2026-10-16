/**
	Plain text, Markdown and CSV: decoding the bytes, turning Markdown into
	safe HTML, and splitting CSV into rows.

		DocText.decode(uint8array)        // {text, encoding}
		DocText.markdown("# Title\n...")  // HTML string; every bit of the input is escaped
		DocText.csv("a;b\n1;2")           // {rows: [["a","b"],["1","2"]], delimiter: ";"}

	Pure logic, tested in node (tools/tests/unit-tests.js).
*/
(function(root) {
	// --- Decoding --------------------------------------------------------

	//* Bytes to text: UTF-8 (or UTF-16 with a byte order mark), else Windows-1252.
	function decode(inBytes) {
		var b = inBytes;
		var Decoder = root.TextDecoder || (typeof TextDecoder !== "undefined" ? TextDecoder : null);
		if (b.length >= 2 && b[0] === 0xFF && b[1] === 0xFE) {
			return {text: new Decoder("utf-16le").decode(b.subarray(2)), encoding: "UTF-16LE"};
		}
		if (b.length >= 2 && b[0] === 0xFE && b[1] === 0xFF) {
			return {text: new Decoder("utf-16be").decode(b.subarray(2)), encoding: "UTF-16BE"};
		}
		var start = b.length >= 3 && b[0] === 0xEF && b[1] === 0xBB && b[2] === 0xBF ? 3 : 0;
		try {
			return {text: new Decoder("utf-8", {fatal: true}).decode(b.subarray(start)), encoding: "UTF-8"};
		} catch (e) {
			return {text: new Decoder("windows-1252").decode(b), encoding: "Windows-1252"};
		}
	}

	// --- Markdown --------------------------------------------------------

	function escapeHtml(inText) {
		return String(inText).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;").replace(/'/g, "&#39;");
	}

	//* Only web and mail links become links; anything else stays text.
	function safeHref(inUrl) {
		var u = String(inUrl || "").trim();
		return /^(https?:\/\/|mailto:)/i.test(u) ? u : null;
	}

	/**
		Inline Markdown on one already-unescaped line: code spans, links,
		bold, italic, strike. The text is escaped first and the markup
		applied to the escaped text, so no tag from the input survives.
	*/
	function inline(inText) {
		var codes = [];
		// Code spans first, so their content is not touched by the rest.
		// \x60 is the backtick (written so, as no template strings are used here).
		var s = String(inText).replace(/\x60([^\x60]+)\x60/g, function(m, c) {
			codes.push("<code>" + escapeHtml(c) + "</code>");
			return "\u0000" + (codes.length - 1) + "\u0000";
		});
		var links = [];
		s = s.replace(/!?\[([^\]]*)\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g, function(m, label, url) {
			var href = safeHref(url);
			var html = href && m.charAt(0) !== "!" ? "<a href=\"" + escapeHtml(href) + "\">" + escapeHtml(label) + "</a>" :
				escapeHtml(label || url);
			links.push(html);
			return "\u0001" + (links.length - 1) + "\u0001";
		});
		s = s.replace(/<(https?:\/\/[^>\s]+)>/g, function(m, url) {
			links.push("<a href=\"" + escapeHtml(url) + "\">" + escapeHtml(url) + "</a>");
			return "\u0001" + (links.length - 1) + "\u0001";
		});
		s = escapeHtml(s);
		s = s.replace(/(\*\*|__)(?=\S)([\s\S]*?\S)\1/g, "<strong>$2</strong>");
		s = s.replace(/(^|[^*\w])\*(?=\S)([^*]*?\S)\*/g, "$1<em>$2</em>");
		s = s.replace(/(^|[^_\w])_(?=\S)([^_]*?\S)_(?!\w)/g, "$1<em>$2</em>");
		s = s.replace(/~~(?=\S)([\s\S]*?\S)~~/g, "<del>$1</del>");
		s = s.replace(/\u0001(\d+)\u0001/g, function(m, i) {
			return links[+i];
		});
		s = s.replace(/\u0000(\d+)\u0000/g, function(m, i) {
			return codes[+i];
		});
		return s;
	}

	function isTableSeparator(inLine) {
		return /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$/.test(inLine);
	}

	function tableCells(inLine) {
		var l = inLine.trim().replace(/^\|/, "").replace(/\|$/, "");
		return l.split("|").map(function(c) {
			return c.trim();
		});
	}

	/**
		Markdown to HTML: headings (# and underlined), paragraphs, emphasis,
		code (spans, fenced and indented blocks), block quotes, ordered and
		unordered lists (nested by indent), horizontal rules, tables and
		links. Raw HTML in the input is shown as text.
	*/
	function markdown(inText) {
		var lines = String(inText).replace(/\r\n?/g, "\n").split("\n");
		var out = [];
		var i = 0;
		var para = [];

		function flushPara() {
			if (para.length) {
				out.push("<p>" + para.map(inline).join("<br>") + "</p>");
				para = [];
			}
		}

		while (i < lines.length) {
			var line = lines[i];
			var m;
			if (/^\s*$/.test(line)) {
				flushPara();
				i++;
				continue;
			}
			// Fenced code.
			if ((m = /^\s*(\x60\x60\x60|~~~)/.exec(line))) {
				flushPara();
				var fence = m[1];
				var code = [];
				i++;
				while (i < lines.length && lines[i].trim().indexOf(fence) !== 0) {
					code.push(lines[i]);
					i++;
				}
				i++;
				out.push("<pre><code>" + escapeHtml(code.join("\n")) + "</code></pre>");
				continue;
			}
			// ATX heading.
			if ((m = /^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$/.exec(line))) {
				flushPara();
				out.push("<h" + m[1].length + ">" + inline(m[2]) + "</h" + m[1].length + ">");
				i++;
				continue;
			}
			// Setext heading: a paragraph line underlined with === or ---.
			if (para.length === 0 && i + 1 < lines.length && /^\s*(=+|-+)\s*$/.test(lines[i + 1]) && !/^\s*[-*+]\s/.test(line) &&
					!/^\s*\|/.test(line)) {
				var level = lines[i + 1].trim().charAt(0) === "=" ? 1 : 2;
				out.push("<h" + level + ">" + inline(line.trim()) + "</h" + level + ">");
				i += 2;
				continue;
			}
			// Horizontal rule.
			if (/^\s{0,3}([-*_])(\s*\1){2,}\s*$/.test(line)) {
				flushPara();
				out.push("<hr>");
				i++;
				continue;
			}
			// Table: a header row, then a separator row.
			if (/\|/.test(line) && i + 1 < lines.length && isTableSeparator(lines[i + 1])) {
				flushPara();
				var head = tableCells(line);
				var aligns = tableCells(lines[i + 1]).map(function(c) {
					return /^:-+:$/.test(c) ? "center" : /-+:$/.test(c) ? "right" : "";
				});
				var cell = function(inTag, inText, inCol) {
					var a = aligns[inCol] ? " style=\"text-align:" + aligns[inCol] + "\"" : "";
					return "<" + inTag + a + ">" + inline(inText) + "</" + inTag + ">";
				};
				var html = "<table><thead><tr>" + head.map(function(c, k) {
					return cell("th", c, k);
				}).join("") + "</tr></thead><tbody>";
				i += 2;
				while (i < lines.length && /\|/.test(lines[i]) && !/^\s*$/.test(lines[i])) {
					html += "<tr>" + tableCells(lines[i]).map(function(c, k) {
						return cell("td", c, k);
					}).join("") + "</tr>";
					i++;
				}
				out.push(html + "</tbody></table>");
				continue;
			}
			// Block quote: the quoted lines, rendered as Markdown themselves.
			if (/^\s{0,3}>/.test(line)) {
				flushPara();
				var quoted = [];
				while (i < lines.length && /^\s{0,3}>/.test(lines[i])) {
					quoted.push(lines[i].replace(/^\s{0,3}>\s?/, ""));
					i++;
				}
				out.push("<blockquote>" + markdown(quoted.join("\n")) + "</blockquote>");
				continue;
			}
			// Lists.
			if (/^\s*([-*+]|\d{1,9}[.)])\s+/.test(line)) {
				flushPara();
				var block = [];
				while (i < lines.length && (/^\s*([-*+]|\d{1,9}[.)])\s+/.test(lines[i]) ||
						(/^\s+\S/.test(lines[i]) && block.length))) {
					block.push(lines[i]);
					i++;
				}
				out.push(list(block));
				continue;
			}
			// Indented code (only outside a paragraph).
			if (para.length === 0 && /^( {4}|\t)/.test(line)) {
				var ind = [];
				while (i < lines.length && (/^( {4}|\t)/.test(lines[i]) || /^\s*$/.test(lines[i]))) {
					ind.push(lines[i].replace(/^( {4}|\t)/, ""));
					i++;
				}
				while (ind.length && /^\s*$/.test(ind[ind.length - 1])) {
					ind.pop();
				}
				out.push("<pre><code>" + escapeHtml(ind.join("\n")) + "</code></pre>");
				continue;
			}
			para.push(line.trim());
			i++;
		}
		flushPara();
		return out.join("\n");
	}

	//* A list block (lines starting with markers, nested by indentation).
	function list(inLines) {
		var indentOf = function(l) {
			return /^\s*/.exec(l.replace(/\t/g, "    "))[0].length;
		};
		var base = indentOf(inLines[0]);
		var ordered = /^\s*\d/.test(inLines[0]);
		var startMatch = /^\s*(\d+)/.exec(inLines[0]);
		var items = [];
		for (var i = 0; i < inLines.length; i++) {
			var l = inLines[i];
			var ind = indentOf(l);
			var m = /^\s*([-*+]|\d{1,9}[.)])\s+(.*)$/.exec(l);
			if (m && ind <= base + 1) {
				items.push({text: m[2], sub: []});
			} else if (items.length) {
				items[items.length - 1].sub.push(l);
			}
		}
		var tag = ordered ? "ol" : "ul";
		var start = ordered && startMatch && +startMatch[1] !== 1 ? " start=\"" + (+startMatch[1]) + "\"" : "";
		return "<" + tag + start + ">" + items.map(function(it) {
			var task = /^\[([ xX])\]\s+(.*)$/.exec(it.text);
			var body = task ? (task[1] === " " ? "☐ " : "☑ ") + inline(task[2]) : inline(it.text);
			var subs = it.sub.filter(function(s) {
				return /^\s*([-*+]|\d{1,9}[.)])\s+/.test(s);
			});
			var cont = it.sub.filter(function(s) {
				return !/^\s*([-*+]|\d{1,9}[.)])\s+/.test(s);
			});
			if (cont.length && !subs.length) {
				body += "<br>" + cont.map(function(s) {
					return inline(s.trim());
				}).join("<br>");
			}
			return "<li>" + body + (subs.length ? list(it.sub.filter(function(s) {
				return !/^\s*$/.test(s);
			})) : "") + "</li>";
		}).join("") + "</" + tag + ">";
	}

	// --- CSV -------------------------------------------------------------

	//* The delimiter that splits the first lines most consistently.
	function sniffDelimiter(inText) {
		var sample = inText.slice(0, 20000).split(/\r\n|\n|\r/).slice(0, 20);
		var best = ",", bestScore = -1;
		var candidates = [",", ";", "\t", "|"];
		for (var c = 0; c < candidates.length; c++) {
			var d = candidates[c];
			var counts = [];
			for (var i = 0; i < sample.length; i++) {
				if (sample[i]) {
					counts.push(splitLine(sample[i], d).length);
				}
			}
			if (!counts.length) {
				continue;
			}
			var first = counts[0];
			var same = counts.filter(function(n) {
				return n === first;
			}).length;
			var score = first > 1 ? same * 100 + first : 0;
			if (score > bestScore) {
				best = d;
				bestScore = score;
			}
		}
		return best;
	}

	function splitLine(inLine, inDelim) {
		return parseRows(inLine, inDelim, 1)[0] || [];
	}

	function escapeRe(inChar) {
		return inChar.replace(/[\\^$.*+?()[\]{}|]/g, "\\$&");
	}

	/**
		RFC 4180 rows: quoted fields with doubled quotes, delimiters and
		line breaks inside quotes. inMax stops after that many rows,
		inMaxCells after that many cells (the result then has .truncated).
		Fields are matched with a sticky regular expression rather than a
		character at a time: a 6 MB file is read in well under a second on
		a desktop.
	*/
	function parseRows(inText, inDelim, inMax, inMaxCells) {
		var text = String(inText);
		var d = escapeRe(inDelim);
		var plain = new RegExp("([^\"" + d + "\\r\\n]*)(" + d + "|\\r\\n|\\n|\\r|$)", "y");
		var quoted = new RegExp("\"((?:[^\"]|\"\")*)\"(" + d + "|\\r\\n|\\n|\\r|$)", "y");
		var rows = [];
		var row = [];
		var pos = 0, n = text.length;
		var max = inMax || Infinity;
		var maxCells = inMaxCells || Infinity;
		var cells = 0;
		while (pos < n) {
			var m = null;
			if (text.charCodeAt(pos) === 34) {
				quoted.lastIndex = pos;
				m = quoted.exec(text);
				if (m) {
					row.push(m[1].indexOf("\"\"") >= 0 ? m[1].replace(/""/g, "\"") : m[1]);
				}
			}
			if (!m) {
				plain.lastIndex = pos;
				m = plain.exec(text);
				if (!m) {
					// A stray quote inside an unquoted field: take the field up to the delimiter or line end.
					var rest = new RegExp("([^" + d + "\\r\\n]*)(" + d + "|\\r\\n|\\n|\\r|$)", "y");
					rest.lastIndex = pos;
					m = rest.exec(text);
				}
				row.push(m[1]);
			}
			cells++;
			pos += m[0].length;
			var sep = m[2];
			if (sep !== inDelim) {
				rows.push(row);
				row = [];
				if (rows.length >= max) {
					return rows;
				}
				if (cells >= maxCells) {
					rows.truncated = pos < n;
					return rows;
				}
				if (!sep) {
					break;
				}
			} else if (pos >= n) {
				// A delimiter at the very end: one more, empty, field.
				row.push("");
				rows.push(row);
				row = [];
			}
			if (!m[0].length) {
				break;
			}
		}
		if (row.length) {
			rows.push(row);
		}
		return rows;
	}

	function csv(inText, inDelimiter, inMaxCells) {
		var text = String(inText);
		var d = inDelimiter || sniffDelimiter(text);
		var rows = parseRows(text, d, 0, inMaxCells);
		return {rows: rows, delimiter: d, truncated: !!rows.truncated};
	}

	//* Rows back to CSV text, quoting where needed.
	function toCsv(inRows, inDelimiter) {
		var d = inDelimiter || ",";
		return inRows.map(function(r) {
			return r.map(function(v) {
				var s = v === null || v === undefined ? "" : String(v);
				return s.indexOf(d) >= 0 || /["\r\n]/.test(s) || /^\s|\s$/.test(s) ? "\"" + s.replace(/"/g, "\"\"") + "\"" : s;
			}).join(d);
		}).join("\r\n") + "\r\n";
	}

	var DocText = {
		decode: decode,
		escapeHtml: escapeHtml,
		safeHref: safeHref,
		markdown: markdown,
		inline: inline,
		csv: csv,
		parseRows: parseRows,
		sniffDelimiter: sniffDelimiter,
		toCsv: toCsv
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocText;
	} else {
		root.DocText = DocText;
	}
})(this);
