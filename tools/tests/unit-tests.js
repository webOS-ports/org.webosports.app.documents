/*
	Unit tests of the readers and helpers in source/core/ (and the
	service's file writer), in node against the fixture documents in
	fixtures/ (built by make-fixtures.py with LibreOffice).

		node tools/tests/unit-tests.js
		TEST_FILTER=odt node tools/tests/unit-tests.js
*/
var fs = require("fs");
var path = require("path");
var t = require("./lib/harness.js");
var assert = t.assert;

var ROOT = path.join(__dirname, "..", "..");
var FIX = path.join(__dirname, "fixtures");
var JSZip = require(path.join(ROOT, "lib/jszip/jszip.min.js"));
var XLSX = require(path.join(ROOT, "lib/sheetjs/xlsx.full.min.js"));
var Xml = require(path.join(ROOT, "source/core/Xml.js"));
var FileTypes = require(path.join(ROOT, "source/core/FileTypes.js"));
var Odf = require(path.join(ROOT, "source/core/Odf.js"));
var Pptx = require(path.join(ROOT, "source/core/Pptx.js"));
var Docx = require(path.join(ROOT, "source/core/Docx.js"));
var Text = require(path.join(ROOT, "source/core/Text.js"));
var Sheets = require(path.join(ROOT, "source/core/Sheets.js"));
var Stream = require(path.join(ROOT, "source/core/SheetStream.js"));
var Library = require(path.join(ROOT, "source/core/Library.js"));
var Engine = require(path.join(ROOT, "source/core/Engine.js"));
var writer = require(path.join(ROOT, "services/org.webosports.app.documents.service/writer.js"));

function bytes(inName) {
	var b = fs.readFileSync(path.join(FIX, inName));
	return new Uint8Array(b.buffer.slice(b.byteOffset, b.byteOffset + b.length));
}

function buffer(inName) {
	return bytes(inName).buffer;
}

async function parts(inName, inList) {
	var zip = await JSZip.loadAsync(fs.readFileSync(path.join(FIX, inName)));
	var out = {};
	var names = inList || Object.keys(zip.files).filter(function(n) {
		return /\.(xml|rels)$/.test(n);
	});
	for (var i = 0; i < names.length; i++) {
		var f = zip.file(names[i]);
		if (f) {
			out[names[i]] = await f.async("string");
		}
	}
	return out;
}

//* All the text runs below a block list, joined.
function textOf(inBlocks) {
	var s = "";
	(inBlocks || []).forEach(function(b) {
		if (b.r) {
			s += runsText(b.r) + "\n";
		}
		if (b.items) {
			b.items.forEach(function(it) {
				s += it.b ? textOf(it.b) : textOf(it);
			});
		}
		if (b.rows) {
			b.rows.forEach(function(r) {
				r.cells.forEach(function(c) {
					s += textOf(c.b);
				});
			});
		}
		if (b.b) {
			s += textOf(b.b);
		}
	});
	return s;
}

function runsText(inRuns) {
	return inRuns.map(function(r) {
		return r.t === "text" ? r.v : r.r ? runsText(r.r) : r.t === "tab" ? "\t" : r.t === "br" ? "\n" : "";
	}).join("");
}

function findRun(inBlocks, inText) {
	var found = null;
	var walk = function(inRuns) {
		inRuns.forEach(function(r) {
			if (r.t === "text" && r.v.indexOf(inText) >= 0 && !found) {
				found = r;
			}
			if (r.r) {
				walk(r.r);
			}
		});
	};
	var blocks = function(inList) {
		(inList || []).forEach(function(b) {
			if (b.r) {
				walk(b.r);
			}
			(b.items || []).forEach(function(it) {
				blocks(it.b || it);
			});
			(b.rows || []).forEach(function(r) {
				r.cells.forEach(function(c) {
					blocks(c.b);
				});
			});
			if (b.b) {
				blocks(b.b);
			}
		});
	};
	blocks(inBlocks);
	return found;
}

// --- Xml -----------------------------------------------------------------

t.test("xml: elements, attributes, text, entities, CDATA and comments", function() {
	var r = Xml.parse("<?xml version=\"1.0\"?><!-- c --><a x=\"1 &amp; 2\" y='q'><b/>t&lt;&#65;&#x42;<![CDATA[<raw>]]><c>z</c></a>");
	assert.strictEqual(r.n, "a");
	assert.strictEqual(r.a.x, "1 & 2");
	assert.strictEqual(r.a.y, "q");
	assert.strictEqual(Xml.textOf(r), "t<AB<raw>z");
	assert.strictEqual(Xml.children(r).length, 2);
});

t.test("xml: prefixes are made canonical from the namespace URI", function() {
	var r = Xml.parse("<t:p xmlns:t=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" t:style-name=\"P1\"><t:span>x</t:span></t:p>");
	assert.strictEqual(r.n, "text:p");
	assert.strictEqual(r.a["text:style-name"], "P1");
	assert.strictEqual(Xml.child(r, "text:span").n, "text:span");
});

t.test("xml: tolerates a stray end tag and an unclosed element", function() {
	var r = Xml.parse("<a><b>x</c></b><d>y</a>");
	assert.strictEqual(Xml.textOf(r), "xy");
	assert.throws(function() {
		Xml.parse("just text");
	});
});

t.test("xml: characters above the BMP and a byte order mark", function() {
	var r = Xml.parse("﻿<a>&#x1F600;</a>");
	assert.strictEqual(Xml.textOf(r), "😀");
});

// --- FileTypes -----------------------------------------------------------

t.test("filetypes: extension and kind from the name", function() {
	assert.strictEqual(FileTypes.extensionOf("/a/b/Report.Final.DOCX"), "docx");
	assert.strictEqual(FileTypes.extensionOf("noext"), "");
	assert.strictEqual(FileTypes.extensionOf(".hidden"), "");
	assert.strictEqual(FileTypes.kindFromName("x.ods"), "ods");
	assert.strictEqual(FileTypes.kindFromName("x.tsv"), "csv");
	assert.strictEqual(FileTypes.kindFromName("x.PDF"), "pdf");
	assert.ok(FileTypes.isDocumentName("a.md"));
	assert.ok(FileTypes.isDocumentName("a.pdf"), "PDFs are listed, to be handed on");
	assert.ok(!FileTypes.isSupported("pdf"));
	assert.ok(!FileTypes.isDocumentName("a.doc"));
	assert.ok(!FileTypes.isDocumentName("a.mp3"));
});

t.test("filetypes: sniffing the first bytes", function() {
	assert.strictEqual(FileTypes.sniff(bytes("sample.odt")), "zip");
	assert.strictEqual(FileTypes.sniff(bytes("sample.xls")), "ole");
	assert.strictEqual(FileTypes.sniff(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2D, 0x31])), "pdf");
	assert.strictEqual(FileTypes.sniff(bytes("sample.txt")), "text");
	assert.strictEqual(FileTypes.sniff(new Uint8Array([1, 2, 0, 4, 5])), "binary");
	assert.strictEqual(FileTypes.sniff(new Uint8Array([0xFF, 0xFE, 0x41, 0x00])), "text");
});

t.test("filetypes: PDFs are handed to the application/pdf handler, nothing else is", function() {
	assert.strictEqual(FileTypes.handOffMime("pdf"), "application/pdf");
	["docx", "xlsx", "txt", "doc", "rtf", "unknown", "", undefined].forEach(function(k) {
		assert.strictEqual(FileTypes.handOffMime(k), null, String(k));
	});
	var pdf = new TextEncoder().encode("%PDF-1.7\n");
	assert.strictEqual(FileTypes.resolve("report.docx", pdf).kind, "pdf");
	assert.strictEqual(FileTypes.resolve("report.pdf", pdf).label, "PDF document");
});

t.test("filetypes: the content wins over a wrong name", function() {
	var odt = bytes("sample.odt");
	var r = FileTypes.resolve("letter.docx", odt, {names: ["mimetype", "content.xml"], mimetype: "application/vnd.oasis.opendocument.text"});
	assert.strictEqual(r.kind, "odt");
	assert.ok(r.mismatch);
	assert.strictEqual(FileTypes.resolve("notes.docx", bytes("sample.txt")).kind, "txt");
	assert.strictEqual(FileTypes.resolve("old.doc", bytes("sample.xls")).kind, "doc");
	assert.strictEqual(FileTypes.resolve("table.bin", bytes("sample.xls")).kind, "xls");
	assert.strictEqual(FileTypes.resolve("thing.zip", odt).kind, "zip");
	assert.strictEqual(FileTypes.resolve("thing.zip", odt, {names: ["a.txt"], mimetype: ""}).kind, "unknown");
	assert.strictEqual(FileTypes.resolve("flat.fodt", bytes("sample.txt")).kind, "odt");
	assert.strictEqual(FileTypes.kindFromZip(["word/document.xml"], ""), "docx");
	assert.strictEqual(FileTypes.kindFromZip(["xl/workbook.xml"], ""), "xlsx");
	assert.strictEqual(FileTypes.kindFromZip(["ppt/presentation.xml"], ""), "pptx");
});

t.test("filetypes: appinfo.json registers every kind it opens, with its extension", function() {
	var appinfo = JSON.parse(fs.readFileSync(path.join(ROOT, "appinfo.json"), "utf8"));
	var registered = appinfo.mimeTypes.map(function(m) {
		return m.extension + " " + m.mime;
	}).sort();
	var expected = FileTypes.appinfoMimeTypes().map(function(m) {
		return m.extension + " " + m.mime;
	}).sort();
	assert.deepStrictEqual(registered, expected);
});

// --- ODF text --------------------------------------------------------------

var odt = null;

async function loadOdt() {
	if (!odt) {
		odt = Odf.text(await parts("sample.odt", ["content.xml", "styles.xml", "meta.xml"]));
	}
	return odt;
}

t.test("odt: headings, styled runs, superscript", async function() {
	var d = await loadOdt();
	assert.strictEqual(d.blocks[0].h, 1);
	assert.strictEqual(runsText(d.blocks[0].r), "Fixture heading");
	assert.strictEqual(findRun(d.blocks, "bold words").s.fontWeight, "bold");
	var red = findRun(d.blocks, "red underlined").s;
	assert.strictEqual(red.color, "#c00000");
	assert.strictEqual(red.textDecoration, "underline");
	assert.strictEqual(findRun(d.blocks, "2").s.verticalAlign, "super");
});

t.test("odt: spaces, tabs, links and footnotes", async function() {
	var d = await loadOdt();
	var p = d.blocks.filter(function(b) {
		return b.r && runsText(b.r).indexOf("Two") === 0;
	})[0];
	// Three spaces kept (ODF collapses plain spaces; LibreOffice writes the extra ones as text:s).
	assert.ok(/^Two[ \u00a0]{3}spaces/.test(runsText(p.r)), JSON.stringify(runsText(p.r)));
	assert.ok(p.r.some(function(r) {
		return r.t === "tab";
	}));
	var link = p.r.filter(function(r) {
		return r.t === "a";
	})[0];
	assert.strictEqual(link.href, "https://www.webos-ports.org/");
	assert.strictEqual(runsText(link.r), "link");
	assert.ok(p.r.some(function(r) {
		return r.t === "note" && r.label === "1";
	}));
	var notes = d.blocks[d.blocks.length - 1];
	assert.strictEqual(notes.t, "notes");
	assert.strictEqual(textOf(notes.items[0].b).trim(), "The footnote text.");
});

t.test("odt: numbered, nested and bulleted lists", async function() {
	var d = await loadOdt();
	var lists = d.blocks.filter(function(b) {
		return b.t === "list";
	});
	assert.strictEqual(lists.length, 2);
	assert.ok(lists[0].o);
	assert.strictEqual(lists[0].type, "decimal");
	assert.strictEqual(lists[0].items.length, 2);
	var nested = lists[0].items[1].filter(function(b) {
		return b.t === "list";
	})[0];
	assert.ok(nested.o);
	assert.strictEqual(nested.type, "lower-alpha");
	assert.ok(!lists[1].o);
	assert.strictEqual(textOf(lists[1].items[0]).trim(), "Bullet one");
});

t.test("odt: table with header row, a span and cell shading", async function() {
	var d = await loadOdt();
	var table = d.blocks.filter(function(b) {
		return b.t === "table";
	})[0];
	assert.strictEqual(table.rows.length, 3);
	assert.ok(table.rows[0].header);
	assert.strictEqual(table.rows[0].cells[0].s.backgroundColor, "#ddeeff");
	assert.strictEqual(table.rows[2].cells[0].cs, 2);
	assert.strictEqual(table.rows[2].cells.length, 2);
	assert.strictEqual(textOf(table.rows[1].cells[2].b).trim(), "3");
});

t.test("odt: pictures, page breaks, page size and metadata", async function() {
	var d = await loadOdt();
	assert.strictEqual(d.images.length, 1);
	assert.ok(/^Pictures\/.*\.png$/.test(d.images[0]), d.images[0]);
	var pic = findRunOfType(d.blocks, "img");
	assert.strictEqual(Math.round(Odf.toPx(pic.w)), Math.round(Odf.toPx("2cm")));
	assert.strictEqual(pic.alt, "Red box");
	var i = d.blocks.map(function(b) {
		return b.t;
	}).indexOf("pagebreak");
	assert.ok(i > 0);
	assert.ok(runsText(d.blocks[i + 1].r).indexOf("Second page") === 0);
	assert.ok(/cm$|in$/.test(d.page.width));
	assert.strictEqual(d.meta.title, "Fixture Text");
	assert.strictEqual(d.meta.author, "Documents tests");
});

function findRunOfType(inBlocks, inType) {
	var found = null;
	inBlocks.forEach(function(b) {
		(b.r || []).forEach(function(r) {
			if (r.t === inType && !found) {
				found = r;
			}
		});
	});
	return found;
}

t.test("odt: hostile values never reach the CSS", function() {
	var css = {};
	Odf.css(css, "color", "red; background: url(http://evil/)");
	Odf.css(css, "width", "expression(alert(1))");
	Odf.css(css, "fontFamily", "'Liberation Sans', sans-serif");
	assert.deepStrictEqual(css, {fontFamily: "'Liberation Sans', sans-serif"});
});

t.test("odt: lengths in pixels", function() {
	assert.strictEqual(Math.round(Odf.toPx("2.54cm")), 96);
	assert.strictEqual(Odf.toPx("72pt"), 96);
	assert.strictEqual(Odf.toPx("1in"), 96);
	assert.ok(isNaN(Odf.toPx("auto")));
});

t.test("odt: a flat .fodt reads the same", function() {
	var xml = fs.readFileSync(path.join(FIX, "src", "sample.fodt"), "utf8");
	var d = Odf.text({"content.xml": xml, "styles.xml": xml, "meta.xml": xml});
	assert.strictEqual(runsText(d.blocks[0].r), "Fixture heading");
	assert.strictEqual(d.meta.title, "Fixture Text");
});

// --- ODF presentation ------------------------------------------------------

t.test("odp: slides, page size, background, title size, list, picture and shape", async function() {
	var deck = Odf.presentation(await parts("sample.odp", ["content.xml", "styles.xml", "meta.xml"]));
	assert.strictEqual(deck.slides.length, 3);
	var s0 = deck.slides[0];
	assert.strictEqual(s0.width, Math.round(Odf.toPx("28cm")));
	assert.strictEqual(s0.bg, "#eeeeff");
	assert.strictEqual(s0.items[0].s.fontSize, "40pt");
	assert.strictEqual(s0.items[0].fill, "");
	assert.strictEqual(s0.items[1].b[0].t, "list");
	var s1 = deck.slides[1];
	assert.strictEqual(s1.items[0].kind, "image");
	assert.ok(deck.images.indexOf(s1.items[0].src) >= 0);
	assert.strictEqual(s1.items[1].fill, "#ffcc00");
	assert.strictEqual(textOf(s1.items[1].b).trim(), "Yellow box");
});

// --- PPTX ------------------------------------------------------------------

t.test("pptx: slides with placeholders, bullets, pictures and backgrounds", async function() {
	var deck = Pptx.parse(await parts("sample.pptx"));
	assert.strictEqual(deck.slides.length, 3);
	var s0 = deck.slides[0];
	assert.strictEqual(s0.bg, "#eeeeff");
	var title = s0.items.filter(function(i) {
		return i.kind === "text" && textOf(i.b).indexOf("Fixture slides") >= 0;
	})[0];
	assert.ok(title, "title item");
	assert.strictEqual(title.b[0].s.fontSize, "53.3px");
	var body = s0.items.filter(function(i) {
		return i.kind === "text" && textOf(i.b).indexOf("First point") >= 0;
	})[0];
	assert.strictEqual(body.b[0].bu, "•", "the symbol-font bullet is drawn as a plain one");
	var pic = deck.slides[1].items.filter(function(i) {
		return i.kind === "image";
	})[0];
	assert.strictEqual(pic.src, "ppt/media/image1.png");
	assert.deepStrictEqual(deck.images, ["ppt/media/image1.png"]);
});

t.test("pptx: colours with luminance changes, number labels, paths", function() {
	var theme = new Pptx.Theme(null, {});
	var el = Xml.parse("<a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:srgbClr val=\"FF0000\"><a:lumMod val=\"50000\"/></a:srgbClr></a:solidFill>");
	assert.strictEqual(Pptx.colorOf(el, theme), "#800000");
	assert.strictEqual(Pptx.numberLabel(3, "arabicPeriod"), "3.");
	assert.strictEqual(Pptx.numberLabel(4, "romanUcPeriod"), "IV.");
	assert.strictEqual(Pptx.numberLabel(28, "alphaLcParenR"), "ab)");
	assert.strictEqual(Pptx.resolvePath("ppt/slides/slide1.xml", "../media/a.png"), "ppt/media/a.png");
	assert.strictEqual(Pptx.resolvePath("ppt/slides/slide1.xml", "/ppt/x.xml"), "ppt/x.xml");
});

// --- DOCX (the reader for large documents) ----------------------------------

t.test("docx: headings, runs, list labels, table, footnote, link and picture", async function() {
	var d = Docx.parse(await parts("sample.docx"));
	assert.strictEqual(d.blocks[0].h, 1);
	assert.strictEqual(runsText(d.blocks[0].r), "Fixture heading");
	assert.strictEqual(findRun(d.blocks, "bold words").s.fontWeight, "bold");
	assert.strictEqual(findRun(d.blocks, "red underlined").s.textDecoration, "underline");
	var labels = d.blocks.filter(function(b) {
		return b.bu;
	}).map(function(b) {
		return b.bu + " " + runsText(b.r);
	});
	assert.deepStrictEqual(labels.slice(0, 3), ["1. First item", "2. Second item", "a) Nested item"]);
	var table = d.blocks.filter(function(b) {
		return b.t === "table";
	})[0];
	assert.strictEqual(table.rows.length, 3);
	assert.strictEqual(table.rows[2].cells[0].cs, 2);
	var notes = d.blocks[d.blocks.length - 1];
	assert.strictEqual(notes.t, "notes");
	assert.strictEqual(textOf(notes.items[0].b).trim(), "The footnote text.");
	var p = d.blocks.filter(function(b) {
		return b.r && b.r.some(function(r) {
			return r.t === "a";
		});
	})[0];
	assert.strictEqual(p.r.filter(function(r) {
		return r.t === "a";
	})[0].href, "https://www.webos-ports.org/");
	assert.deepStrictEqual(d.images, ["word/media/image1.png"]);
	assert.ok(/pt$/.test(d.page.width));
});

t.test("docx: runs carry only what differs from their paragraph", async function() {
	var d = Docx.parse(await parts("sample.docx"));
	var p = d.blocks[1];
	assert.ok(p.s.fontSize, "the paragraph has the size");
	var plain = p.r[0];
	assert.ok(!plain.s || !plain.s.fontSize, "the run does not repeat it");
});

// --- Text, Markdown, CSV ----------------------------------------------------

t.test("text: UTF-8, UTF-16 with BOM, and Windows-1252 fallback", function() {
	assert.strictEqual(Text.decode(new Uint8Array([0xC3, 0xA9])).text, "é");
	assert.strictEqual(Text.decode(new Uint8Array([0xFF, 0xFE, 0x41, 0x00])).text, "A");
	var r = Text.decode(new Uint8Array([0x63, 0x61, 0x66, 0xE9]));
	assert.strictEqual(r.text, "café");
	assert.strictEqual(r.encoding, "Windows-1252");
});

t.test("markdown: markup is escaped, only web links stay links", function() {
	var html = Text.markdown(fs.readFileSync(path.join(FIX, "sample.md"), "utf8"));
	assert.ok(html.indexOf("<script") < 0);
	assert.ok(html.indexOf("&lt;script&gt;") >= 0);
	assert.ok(html.indexOf("<h1>Markdown fixture</h1>") >= 0);
	assert.ok(html.indexOf("<a href=\"https://www.webos-ports.org/\">link</a>") >= 0);
	assert.ok(/<table>.*<th>Name<\/th>.*text-align:right/.test(html));
	assert.ok(html.indexOf("<pre><code>code block &lt;b&gt;kept&lt;/b&gt;</code></pre>") >= 0);
	assert.ok(/<ul><li>one<\/li><li>two<ul><li>nested<\/li><\/ul><\/li><\/ul>/.test(html), html);
	assert.ok(html.indexOf("<blockquote>") >= 0);
	assert.ok(!/<a /.test(Text.markdown("[x](javascript:alert(1))")));
	assert.strictEqual(Text.markdown("<img src=x onerror=alert(1)>"), "<p>&lt;img src=x onerror=alert(1)&gt;</p>");
	assert.strictEqual(Text.markdown("[a](\"onmouseover=\")"), "<p>a</p>");
});

t.test("csv: delimiter sniffing, quotes, doubled quotes, line breaks in fields", function() {
	var r = Text.csv(fs.readFileSync(path.join(FIX, "semicolon.csv"), "utf8"));
	assert.strictEqual(r.delimiter, ";");
	assert.deepStrictEqual(r.rows[1], ["Smith; J.", "12,5", "said \"hi\""]);
	assert.deepStrictEqual(Text.parseRows("a,\"b\nc\",d\n", ","), [["a", "b\nc", "d"]]);
	assert.deepStrictEqual(Text.parseRows("a\tb\n", "\t"), [["a", "b"]]);
	assert.strictEqual(Text.sniffDelimiter("a|b|c\n1|2|3\n"), "|");
});

t.test("csv: a cap on cells cuts at a row and says so", function() {
	var lines = [];
	for (var i = 0; i < 100; i++) {
		lines.push(i + ",x,y");
	}
	var r = Text.csv(lines.join("\n"), ",", 30);
	assert.strictEqual(r.rows.length, 10);
	assert.ok(r.truncated);
});

t.test("csv: writing back quotes what needs quoting", function() {
	assert.strictEqual(Text.toCsv([["a", "b,c", "d\"e"], [" x", 1, null]], ","), "a,\"b,c\",\"d\"\"e\"\r\n\" x\",1,\r\n");
});

// --- Sheets (SheetJS) ---------------------------------------------------------

function workbook(inName) {
	return XLSX.read(bytes(inName), {type: "array", dense: true, cellFormula: true});
}

["sample.xlsx", "sample.ods", "sample.xls"].forEach(function(inName) {
	t.test("sheets: " + inName + " values, kinds, formulas, merges, second sheet", function() {
		var m = Sheets.fromWorkbook(XLSX, workbook(inName));
		assert.strictEqual(m.sheets.length, 2);
		var s = m.sheets[0];
		assert.strictEqual(s.name, "Prices");
		assert.deepStrictEqual(s.text[0], ["Item", "Price", "Qty", "Total"]);
		assert.strictEqual(s.text[2][1], "12.25");
		assert.strictEqual(s.kinds[2][1], "n");
		assert.strictEqual(s.kinds[2][0], "s");
		assert.strictEqual(s.formulas["4,3"], "=SUM(D2:D4)");
		assert.deepStrictEqual(s.merges, [{r: 4, c: 0, rs: 1, cs: 3}]);
		assert.strictEqual(m.sheets[1].text[1][0], "findme in sheet two");
	});
});

t.test("sheets: references and typed input", function() {
	assert.strictEqual(Sheets.colName(0), "A");
	assert.strictEqual(Sheets.colName(27), "AB");
	assert.strictEqual(Sheets.cellRef(9, 2), "C10");
	assert.deepStrictEqual(Sheets.parseRef("$AB$12"), {r: 11, c: 27});
	assert.strictEqual(Sheets.parseRef("12"), null);
	assert.deepStrictEqual(Sheets.parseInput("12,5"), {t: "n", v: 12.5});
	assert.deepStrictEqual(Sheets.parseInput("=A1*2"), {f: "A1*2"});
	assert.deepStrictEqual(Sheets.parseInput("true"), {t: "b", v: true});
	assert.deepStrictEqual(Sheets.parseInput(""), {t: "z"});
	assert.deepStrictEqual(Sheets.parseInput("12 apples"), {t: "s", v: "12 apples"});
});

["xlsx", "ods", "csv"].forEach(function(inFormat) {
	t.test("sheets: an edited copy written as " + inFormat + " reads back with the edits", function() {
		var wb = workbook("sample.xlsx");
		Sheets.setCell(XLSX, wb, 0, 1, 2, "40");
		Sheets.setCell(XLSX, wb, 0, 6, 0, "new row");
		var out = Sheets.write(XLSX, wb, inFormat);
		var back = XLSX.read(new Uint8Array(out), {type: "array", dense: true});
		var m = Sheets.fromWorkbook(XLSX, back);
		assert.strictEqual(m.sheets[0].text[1][2], "40");
		assert.strictEqual(m.sheets[0].text[6][0], "new row");
		assert.strictEqual(m.sheets[0].text[2][0], "Book", "the rest is kept");
	});
});

// --- Streamed sheets ----------------------------------------------------------

function chunks(inText, inSize) {
	var out = [];
	for (var i = 0; i < inText.length; i += inSize) {
		out.push(inText.slice(i, i + inSize));
	}
	return out;
}

t.test("stream: an xlsx sheet fed in tiny chunks reads as the full reader does", async function() {
	var p = await parts("sample.xlsx");
	var sst = new Stream.SharedStrings();
	chunks(p["xl/sharedStrings.xml"], 7).forEach(function(c) {
		sst.feed(c);
	});
	var rows = {};
	var parser = new Stream.XlsxSheetParser({sst: sst.end(), styles: Stream.readStyles(p["xl/styles.xml"]),
		format: function(f, v) {
			return XLSX.SSF.format(f, v);
		}, onRow: function(r, texts, kinds) {
			rows[r] = {texts: texts, kinds: kinds};
		}});
	chunks(p["xl/worksheets/sheet1.xml"], 5).forEach(function(c) {
		parser.feed(c);
	});
	parser.end();
	var full = Sheets.fromWorkbook(XLSX, workbook("sample.xlsx")).sheets[0];
	for (var r = 0; r < full.text.length; r++) {
		for (var c = 0; c < 4; c++) {
			var expect = full.text[r] && full.text[r][c] !== undefined ? full.text[r][c] : undefined;
			assert.strictEqual(rows[r] && rows[r].texts[c], expect, "row " + r + " col " + c);
		}
	}
	assert.strictEqual(rows[1].kinds[1], "n");
	assert.deepStrictEqual(parser.merges, [{r: 4, c: 0, rs: 1, cs: 3}]);
	assert.ok(parser.colWidths[0] > parser.colWidths[1], "the wide first column");
});

t.test("stream: an ods content.xml fed in tiny chunks, with sheets, merges and repeats", async function() {
	var p = await parts("sample.ods", ["content.xml"]);
	var sheets = [];
	var parser = new Stream.OdsParser({
		onSheet: function(n) {
			sheets.push({name: n, rows: {}});
		},
		onRow: function(r, texts) {
			sheets[sheets.length - 1].rows[r] = texts.slice(0);
		},
		onSheetEnd: function(info) {
			sheets[sheets.length - 1].info = info;
		}
	});
	chunks(p["content.xml"], 9).forEach(function(c) {
		parser.feed(c);
	});
	parser.end();
	assert.deepStrictEqual(sheets.map(function(s) {
		return s.name;
	}), ["Prices", "Notes"]);
	assert.deepStrictEqual(sheets[0].rows[0], ["Item", "Price", "Qty", "Total"]);
	assert.strictEqual(sheets[0].rows[3][3], "30");
	assert.deepStrictEqual(sheets[0].info.merges, [{r: 4, c: 0, rs: 1, cs: 3}]);
	assert.strictEqual(sheets[1].rows[1][0], "findme in sheet two");
	assert.ok(sheets[0].info.colWidths[0] > sheets[0].info.colWidths[1]);
});

t.test("stream: the cell budget stops a sheet and marks it cut", function() {
	var rows = "";
	for (var i = 1; i <= 50; i++) {
		rows += "<row r=\"" + i + "\"><c r=\"A" + i + "\"><v>" + i + "</v></c><c r=\"B" + i + "\" t=\"inlineStr\"><is><t>x&amp;y</t></is></c></row>";
	}
	var got = [];
	var p = new Stream.XlsxSheetParser({sst: [], maxCells: 20, onRow: function(r, texts) {
		got.push(texts);
	}});
	p.feed("<worksheet><sheetData>" + rows + "</sheetData></worksheet>");
	p.end();
	assert.strictEqual(got.length, 10);
	assert.strictEqual(got[0][1], "x&y");
	assert.ok(p.truncated);
});

t.test("stream: numbers in the General format", function() {
	assert.strictEqual(Stream.generalNumber(12), "12");
	assert.strictEqual(Stream.generalNumber(0.1 + 0.2), "0.3");
	assert.strictEqual(Stream.generalNumber(1234567.891), "1234567.891");
	assert.strictEqual(Stream.colOfRef("AA10"), 26);
	assert.deepStrictEqual(Stream.rangeOf("B2:D5"), {r: 1, c: 1, r1: 4, c1: 3});
	assert.strictEqual(Stream.odsText("<text:p>a<text:s text:c=\"2\"/>b</text:p><text:p>c</text:p>"), "a  b\nc");
});

// --- Library -------------------------------------------------------------------

t.test("library: names joined by commas are put back together", function() {
	var names = Library.splitNames("README,Report, final.docx,a.xlsx,notes");
	assert.deepStrictEqual(names.map(function(n) {
		return n.name;
	}), ["README,Report, final.docx", "a.xlsx", "notes"]);
	assert.ok(names[0].joined);
	assert.strictEqual(names[0].alone, " final.docx");
});

t.test("library: the scanner walks folders, skips hidden ones and checks rebuilt names", async function() {
	var tree = {
		"/m": {dirs: "Docs,.hidden,DCIM", files: "a.odt,song.mp3"},
		"/m/Docs": {dirs: "", files: "Report, final.docx,b.csv"},
		"/m/.hidden": {dirs: "", files: "secret.docx"},
		"/m/DCIM": {dirs: "", files: "x.jpg"}
	};
	var existing = {"/m/Docs/Report, final.docx": true};
	var docs = await new Promise(function(resolve) {
		new Library.Scanner({
			list: function(inDir, inCallback) {
				setTimeout(function() {
					inCallback(tree[inDir] ? null : new Error("no"), tree[inDir]);
				}, 1);
			},
			exists: function(inPath, inCallback) {
				setTimeout(function() {
					inCallback(!!existing[inPath]);
				}, 1);
			},
			onDone: resolve
		}).start("/m");
	});
	assert.deepStrictEqual(docs.map(function(d) {
		return d.path;
	}), ["/m/a.odt", "/m/Docs/b.csv", "/m/Docs/Report, final.docx"]);
	assert.strictEqual(docs[2].kind, "docx");
});

t.test("library: launch targets, URLs and copy names", function() {
	assert.strictEqual(Library.targetToPath("file:///media/internal/My%20Doc.odt"), "/media/internal/My Doc.odt");
	assert.strictEqual(Library.targetToPath("/media/internal/a/../b.docx"), "/media/internal/b.docx");
	assert.strictEqual(Library.targetToPath("https://example.com/a.docx"), null);
	assert.strictEqual(Library.targetToPath("relative.docx"), null);
	assert.strictEqual(Library.targetToPath("file:///media/internal/%E0%A4%A"), null);
	assert.strictEqual(Library.pathToUrl("/media/internal/a b#c.odt"), "file:///media/internal/a%20b%23c.odt");
	assert.strictEqual(Library.copyName("Budget.xlsx", "xlsx", []), "Budget (edited).xlsx");
	assert.strictEqual(Library.copyName("Budget (edited).xlsx", "xlsx", ["budget (edited).xlsx"]), "Budget (edited 2).xlsx");
	assert.strictEqual(Library.copyName("Old.xls", "xlsx", []), "Old (edited).xlsx");
	var recent = Library.addRecent([{path: "/a"}, {path: "/b"}], {path: "/b"}, 2);
	assert.deepStrictEqual(recent.map(function(r) {
		return r.path;
	}), ["/b", "/a"]);
	assert.strictEqual(Library.formatSize(1536), "1.5 KB");
});

// --- Engine ------------------------------------------------------------------------

function engine(inLibs) {
	return new Engine(Object.assign({JSZip: JSZip, XLSX: XLSX}, inLibs || {}));
}

var expectKinds = {"sample.odt": ["odt", "text"], "sample.docx": ["docx", "text"], "sample.xlsx": ["xlsx", "sheet"],
	"sample.xls": ["xls", "sheet"], "sample.ods": ["ods", "sheet"], "sample.pptx": ["pptx", "slides"],
	"sample.odp": ["odp", "slides"], "sample.md": ["md", "text"], "sample.txt": ["txt", "text"], "semicolon.csv": ["csv", "sheet"]};

Object.keys(expectKinds).forEach(function(inName) {
	t.test("engine: opens " + inName, async function() {
		var r = await engine().open({name: inName, buffer: buffer(inName)});
		assert.deepStrictEqual([r.kind, r.family], expectKinds[inName]);
	});
});

t.test("engine: pictures come out of the package with their type", async function() {
	var r = await engine().open({name: "sample.odt", buffer: buffer("sample.odt")});
	var img = r.images[r.model.images[0]];
	assert.strictEqual(img.mime, "image/png");
	assert.strictEqual(img.bytes[1], 0x50, "PNG signature");
});

t.test("engine: a large Word document goes to the fast reader", async function() {
	var r = await engine({noDocxPreview: true}).open({name: "sample.docx", buffer: buffer("sample.docx")});
	assert.ok(r.simplified);
	assert.strictEqual(r.model.blocks[0].h, 1);
	assert.ok(r.images["word/media/image1.png"].bytes.length > 0);
});

t.test("engine: a streamed workbook matches the full reader and keeps sending rows", async function() {
	var updates = [];
	var finished;
	var done = new Promise(function(resolve) {
		finished = resolve;
	});
	var r = await engine({forceStream: true}).open({name: "sample.ods", buffer: buffer("sample.ods"), onProgress: function(u) {
		updates.push(u);
		if (u.finished) {
			finished();
		}
	}});
	await done;
	assert.ok(r.streamed);
	assert.deepStrictEqual(r.model.sheets[0].text[0], ["Item", "Price", "Qty", "Total"]);
	var second = updates.filter(function(u) {
		return u.sheet === 1 && u.rows;
	})[0];
	assert.strictEqual(second.rows[1][1][0], "findme in sheet two");
});

t.test("engine: binary legacy and PDF files are refused with a clear message", async function() {
	var ole = new Uint8Array(4096);
	[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1].forEach(function(b, i) {
		ole[i] = b;
	});
	await assert.rejects(engine().open({name: "old.doc", buffer: ole.buffer}), /before 2007/);
	var pdf = new TextEncoder().encode("%PDF-1.4\n...");
	await assert.rejects(engine().open({name: "x.pdf", buffer: pdf.buffer}), /PDF document, and no app/);
});

t.test("engine: an edited CSV is saved as CSV with the edits", async function() {
	var e = engine();
	await e.open({name: "semicolon.csv", buffer: buffer("semicolon.csv")});
	var out = await e.save({edits: [{sheet: 0, r: 2, c: 2, text: "new; note"}], format: "csv"});
	var text = new TextDecoder().decode(out);
	assert.ok(text.indexOf("Jones;7;\"new; note\"") >= 0, text);
});

t.test("engine: an edited xlsx is saved with the edit, from the workbook it read", async function() {
	var e = engine();
	await e.open({name: "sample.xlsx", buffer: buffer("sample.xlsx")});
	var out = await e.save({edits: [{sheet: 0, r: 1, c: 0, text: "Pencil"}], format: "xlsx"});
	var m = Sheets.fromWorkbook(XLSX, XLSX.read(out, {type: "array", dense: true}));
	assert.strictEqual(m.sheets[0].text[1][0], "Pencil");
});

// --- The service's writer ------------------------------------------------------------

var tmpRoot = path.join(ROOT, "test-output", "writer-root") + "/";

t.test("service: writes a new file and never replaces one", function() {
	fs.rmSync(tmpRoot, {recursive: true, force: true});
	fs.mkdirSync(tmpRoot, {recursive: true});
	var target = tmpRoot + "copy.csv";
	var r = writer.writeNew(tmpRoot, target, Buffer.from("a,b\n").toString("base64"));
	assert.strictEqual(r.returnValue, true);
	assert.strictEqual(fs.readFileSync(target, "utf8"), "a,b\n");
	var again = writer.writeNew(tmpRoot, target, Buffer.from("OVERWRITE").toString("base64"));
	assert.strictEqual(again.returnValue, false);
	assert.strictEqual(again.errorCode, "EEXIST");
	assert.strictEqual(fs.readFileSync(target, "utf8"), "a,b\n", "the existing file is untouched");
});

t.test("service: refuses paths outside the root, hidden names and other types", function() {
	var data = Buffer.from("x").toString("base64");
	["/etc/passwd.csv", tmpRoot + "../escape.csv", tmpRoot + "a/../../b.csv", tmpRoot + ".hidden.csv", tmpRoot + ".dir/a.csv",
		tmpRoot + "script.sh", tmpRoot + "x.xlsx\u0000.sh", tmpRoot, "relative.csv"].forEach(function(p) {
		var r = writer.writeNew(tmpRoot, p, data);
		assert.strictEqual(r.errorCode, "EINVAL", p);
	});
	assert.strictEqual(writer.writeNew(tmpRoot, tmpRoot + "empty.csv", "").errorCode, "EINVAL");
	assert.ok(!fs.existsSync(tmpRoot + "../escape.csv"));
});

t.test("service: writes copies of every format the app edits", function() {
	["c.xlsx", "c.ods", "c.csv", "c.odt", "c.docx", "c.pptx", "c.odp", "c.txt", "c.md"].forEach(function(n) {
		var r = writer.writeNew(tmpRoot, tmpRoot + n, Buffer.from("x").toString("base64"));
		assert.strictEqual(r.returnValue, true, n);
	});
	assert.strictEqual(writer.writeNew(tmpRoot, tmpRoot + "c.html", Buffer.from("x").toString("base64")).errorCode, "EINVAL");
});

t.test("service: a missing folder fails cleanly and leaves nothing behind", function() {
	var r = writer.writeNew(tmpRoot, tmpRoot + "nofolder/x.ods", Buffer.from("x").toString("base64"));
	assert.strictEqual(r.returnValue, false);
	assert.strictEqual(r.errorCode, "ENOENT");
});

t.run();
