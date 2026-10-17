/*
	Unit tests of editing: the XML parser's offsets, saving edited cells
	into the workbook's own file (SheetPatch.js), working formulas out
	again (Formulas.js), editing OpenDocument paragraphs (TextEdit.js),
	and the engine's copies of each. Node only.

		node tools/tests/edit-tests.js
		TEST_FILTER=ods node tools/tests/edit-tests.js
*/
var fs = require("fs");
var path = require("path");
var t = require("./lib/harness.js");
var assert = t.assert;

var ROOT = path.join(__dirname, "..", "..");
var FIX = path.join(__dirname, "fixtures");
var JSZip = require(path.join(ROOT, "lib/jszip/jszip.min.js"));
var XLSX = require(path.join(ROOT, "lib/sheetjs/xlsx.full.min.js"));
var HyperFormula = require(path.join(ROOT, "lib/hyperformula/hyperformula.full.min.js"));
var Xml = require(path.join(ROOT, "source/core/Xml.js"));
var Odf = require(path.join(ROOT, "source/core/Odf.js"));
var Sheets = require(path.join(ROOT, "source/core/Sheets.js"));
var SheetPatch = require(path.join(ROOT, "source/core/SheetPatch.js"));
var Formulas = require(path.join(ROOT, "source/core/Formulas.js"));
var TextEdit = require(path.join(ROOT, "source/core/TextEdit.js"));
var Engine = require(path.join(ROOT, "source/core/Engine.js"));
var DocxCopy = require(path.join(ROOT, "source/core/DocxCopy.js"));
var Pptx = require(path.join(ROOT, "source/core/Pptx.js"));

function buffer(inName) {
	var b = fs.readFileSync(path.join(FIX, inName));
	return b.buffer.slice(b.byteOffset, b.byteOffset + b.length);
}

function engine() {
	return new Engine({JSZip: JSZip, XLSX: XLSX, HyperFormula: HyperFormula, noDocxPreview: true});
}

//* Every entry of a zip: {name: {text (uncompressed), method}}.
async function entries(inBytes) {
	var zip = await JSZip.loadAsync(inBytes);
	var out = {};
	var names = Object.keys(zip.files);
	for (var i = 0; i < names.length; i++) {
		var f = zip.files[names[i]];
		if (!f.dir) {
			out[names[i]] = await f.async("string");
		}
	}
	return out;
}

//* The parts named, as text.
async function parts(inName, inList) {
	var zip = await JSZip.loadAsync(fs.readFileSync(path.join(FIX, inName)));
	var out = {};
	for (var i = 0; i < inList.length; i++) {
		var f = zip.file(inList[i]);
		if (f) {
			out[inList[i]] = await f.async("string");
		}
	}
	return out;
}

//* Every element's offsets, depth first, for comparing two parses.
function offsetsOf(inEl, inOut) {
	var out = inOut || [];
	out.push([inEl.n, inEl.p, inEl.ci, inEl.ce, inEl.e]);
	for (var i = 0; i < inEl.c.length; i++) {
		if (typeof inEl.c[i] === "object") {
			offsetsOf(inEl.c[i], out);
		} else {
			out.push(["#text", inEl.co[i][0], inEl.co[i][1]]);
		}
	}
	return out;
}

function applyDiff(inOld, inHunks) {
	var s = "", p = 0;
	inHunks.forEach(function(h) {
		assert.ok(h.s >= p, "hunks in order, not overlapping");
		s += inOld.slice(p, h.s) + h.ins;
		p = h.e;
	});
	return s + inOld.slice(p);
}

// --- XML offsets ----------------------------------------------------------------

t.test("xml offsets: elements, content and text chunks map back to the source", function() {
	var s = "<?xml version=\"1.0\"?>﻿<r xmlns:t=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"><t:p a=\"1\">he&amp;llo<t:s/>" +
		"<t:span>x</t:span></t:p><q><![CDATA[c<d]]></q><u>open</r>";
	var r = Xml.parse(s, {offsets: true});
	var p = r.c[0];
	assert.strictEqual(p.n, "text:p");
	assert.strictEqual(s.slice(p.p, p.e), "<t:p a=\"1\">he&amp;llo<t:s/><t:span>x</t:span></t:p>");
	assert.strictEqual(s.slice(p.ci, p.ce), "he&amp;llo<t:s/><t:span>x</t:span>");
	assert.strictEqual(s.slice(p.co[0][0], p.co[0][1]), "he&amp;llo");
	assert.strictEqual(p.c[0], "he&llo");
	var empty = p.c[1];
	assert.strictEqual(empty.ci, empty.e, "an empty element has no content");
	assert.strictEqual(s.slice(r.c[1].co[0][0], r.c[1].co[0][1]), "<![CDATA[c<d]]>");
	assert.strictEqual(s.slice(r.c[2].p, r.c[2].e), "<u>open", "unclosed: ends where its parent's end tag starts");
	// A piece parsed again on its own, in its scope and at its place, reads the same.
	var again = Xml.parse(s.slice(p.p, p.e), {offsets: true, scope: p.ns, base: p.p});
	assert.deepStrictEqual(offsetsOf(again), offsetsOf(p));
});

t.test("xml: splice applies changes in order and refuses overlapping ones", function() {
	assert.strictEqual(Xml.splice("abcdef", [{s: 4, e: 5, v: "X"}, {s: 1, e: 1, v: "1"}, {s: 1, e: 1, v: "2"}]), "a12bcdXf");
	assert.throws(function() {
		Xml.splice("abcdef", [{s: 1, e: 4, v: ""}, {s: 2, e: 3, v: ""}]);
	}, /overlapping/);
	assert.strictEqual(Xml.escape("a<b>&\"c\""), "a&lt;b&gt;&amp;&quot;c&quot;");
});

// --- Workbooks: Office Open XML -------------------------------------------------

var SHEET_HEAD = "<?xml version=\"1.0\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";

function xlsxParts(inSheetXml, inWorkbookExtra) {
	return {
		"xl/workbook.xml": "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
			"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>" +
			"<sheet name=\"S\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" + (inWorkbookExtra || "") + "</workbook>",
		"xl/_rels/workbook.xml.rels": "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
			"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
			"<Relationship Id=\"rId9\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.xml\"/>" +
			"</Relationships>",
		"[Content_Types].xml": "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
			"<Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/></Types>",
		"xl/worksheets/sheet1.xml": inSheetXml
	};
}

t.test("xlsx patch: replaced, new and cleared cells; styles kept; rows and cells in order", function() {
	var sheet = SHEET_HEAD + "<dimension ref=\"A1:B3\"/><sheetData>" +
		"<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"s\"><v>0</v></c><c r=\"C1\" s=\"4\"><v>9</v></c></row>" +
		"<row r=\"3\" s=\"7\" customFormat=\"1\"><c r=\"B3\"><v>2</v></c></row>" +
		"</sheetData><mergeCells count=\"1\"><mergeCell ref=\"A5:B5\"/></mergeCells></worksheet>";
	var out = SheetPatch.xlsx(xlsxParts(sheet), [
		{sheet: 0, r: 0, c: 0, text: "Fountain <pen> & ink"},   // replaced, keeps s="3"
		{sheet: 0, r: 0, c: 1, text: "12,5"},                   // new between A1 and C1
		{sheet: 0, r: 0, c: 2, text: ""},                       // cleared: keeps its style
		{sheet: 0, r: 1, c: 3, text: "TRUE"},                   // a new row between 1 and 3
		{sheet: 0, r: 2, c: 0, text: "x"},                      // new, takes the row's style
		{sheet: 0, r: 9, c: 1, text: "=B3*2"}                   // a new row at the end
	]);
	var x = out["xl/worksheets/sheet1.xml"];
	assert.ok(x.indexOf("<c r=\"A1\" s=\"3\" t=\"inlineStr\"><is><t xml:space=\"preserve\">Fountain &lt;pen&gt; &amp; ink</t></is></c>" +
		"<c r=\"B1\"><v>12.5</v></c><c r=\"C1\" s=\"4\"/></row>") > 0, x);
	assert.ok(x.indexOf("</row><row r=\"2\"><c r=\"D2\" t=\"b\"><v>1</v></c></row><row r=\"3\"") > 0, x);
	assert.ok(x.indexOf("<c r=\"A3\" s=\"7\" t=\"inlineStr\">") > 0, "a new cell in a formatted row takes its style");
	assert.ok(x.indexOf("<row r=\"10\"><c r=\"B10\"><f>B3*2</f></c></row></sheetData>") > 0, x);
	assert.ok(x.indexOf("<mergeCells count=\"1\"><mergeCell ref=\"A5:B5\"/></mergeCells>") > 0, "the rest of the sheet is kept");
	assert.ok(x.indexOf("<dimension ref=\"A1:D10\"/>") > 0, "the used range grows");
	Xml.parse(x);
	// The workbook is worked out again when opened; the calculation chain is dropped.
	assert.ok(/<calcPr fullCalcOnLoad="1"\/>/.test(out["xl/workbook.xml"]), out["xl/workbook.xml"]);
	assert.strictEqual(out["xl/calcChain.xml"], null);
	assert.ok(!/calcChain/.test(out["xl/_rels/workbook.xml.rels"]));
	assert.ok(!/calcChain/.test(out["[Content_Types].xml"]));
});

t.test("xlsx patch: an empty sheet, an existing calcPr, a prefixed namespace", function() {
	var out = SheetPatch.xlsx(xlsxParts(SHEET_HEAD + "<sheetData/></worksheet>", "<calcPr calcId=\"191029\" fullCalcOnLoad=\"0\"/>"),
		[{sheet: 0, r: 1, c: 1, text: "5"}]);
	assert.ok(out["xl/worksheets/sheet1.xml"].indexOf("<sheetData><row r=\"2\"><c r=\"B2\"><v>5</v></c></row></sheetData>") > 0);
	assert.ok(/<calcPr calcId="191029" fullCalcOnLoad="1"\/>/.test(out["xl/workbook.xml"]), out["xl/workbook.xml"]);
	var prefixed = "<x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheetData>" +
		"<x:row r=\"1\"><x:c r=\"A1\"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>";
	var out2 = SheetPatch.xlsx(xlsxParts(prefixed), [{sheet: 0, r: 0, c: 1, text: "hi"}]);
	assert.ok(out2["xl/worksheets/sheet1.xml"].indexOf("<x:c r=\"B1\" t=\"inlineStr\"><x:is><x:t xml:space=\"preserve\">hi</x:t></x:is></x:c></x:row>") > 0,
		out2["xl/worksheets/sheet1.xml"]);
});

t.test("xlsx patch: cells sharing an edited formula get it written out; results replace cached values", function() {
	var sheet = SHEET_HEAD + "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><f t=\"shared\" ref=\"B1:B3\" si=\"0\">A1*2</f><v>2</v></c></row>" +
		"<row r=\"2\"><c r=\"A2\"><v>2</v></c><c r=\"B2\"><f t=\"shared\" si=\"0\"/><v>4</v></c></row>" +
		"<row r=\"3\"><c r=\"A3\"><v>3</v></c><c r=\"B3\" s=\"2\"><f t=\"shared\" si=\"0\"/><v>6</v></c></row></sheetData></worksheet>";
	var formulas = {"1,1": "A2*2", "2,1": "A3*2"};
	var out = SheetPatch.xlsx(xlsxParts(sheet), [
		{sheet: 0, r: 0, c: 1, text: "=A1*10", value: {v: 10, t: "n"}},
		{sheet: 0, r: 2, c: 0, text: "7"},
		{sheet: 0, r: 2, c: 1, valueOnly: true, value: {v: 14, t: "n"}}
	], {formulaAt: function(s, r, c) {
		return formulas[r + "," + c];
	}});
	var x = out["xl/worksheets/sheet1.xml"];
	assert.ok(x.indexOf("<c r=\"B1\"><f>A1*10</f><v>10</v></c>") > 0, x);
	assert.ok(x.indexOf("<c r=\"B2\"><f>A2*2</f><v>4</v></c>") > 0, "a sharer gets its own formula: " + x);
	assert.ok(x.indexOf("<c r=\"B3\" s=\"2\"><f>A3*2</f><v>14</v></c>") > 0, "a sharer with a new result: " + x);
	assert.throws(function() {
		SheetPatch.xlsx(xlsxParts(sheet), [{sheet: 0, r: 0, c: 1, text: "1"}], {});
	}, /shared/, "without the formulas a shared formula is not broken silently");
	// A result that is a function the calculator does not know is left to the reading program.
	var out2 = SheetPatch.xlsx(xlsxParts(sheet), [{sheet: 0, r: 1, c: 1, valueOnly: true, value: {v: "#NAME?", t: "e", name: true}}]);
	assert.ok(out2["xl/worksheets/sheet1.xml"].indexOf("<c r=\"B2\"><f t=\"shared\" si=\"0\"/></c>") > 0, out2["xl/worksheets/sheet1.xml"]);
});

// --- Workbooks: OpenDocument ----------------------------------------------------

var ODS_HEAD = "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
	"xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" " +
	"xmlns:calcext=\"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0\"><office:body><office:spreadsheet>";
var ODS_TAIL = "</office:spreadsheet></office:body></office:document-content>";

t.test("ods patch: repeated rows and cells are split around the edited cell", function() {
	var content = ODS_HEAD + "<table:table table:name=\"S\"><table:table-column table:number-columns-repeated=\"5\"/>" +
		"<table:table-row table:style-name=\"ro1\"><table:table-cell office:value-type=\"string\"><text:p>a</text:p></table:table-cell>" +
		"<table:table-cell table:style-name=\"ce2\" table:number-columns-repeated=\"4\"/></table:table-row>" +
		"<table:table-row table:style-name=\"ro1\" table:number-rows-repeated=\"1000\"><table:table-cell table:number-columns-repeated=\"5\"/></table:table-row>" +
		"</table:table>" + ODS_TAIL;
	var out = SheetPatch.ods({"content.xml": content}, [
		{sheet: 0, r: 0, c: 2, text: "3.5"},
		{sheet: 0, r: 4, c: 1, text: "=SUM(A1:C1)*2"},
		{sheet: 0, r: 4, c: 3, text: "hello & bye"}
	])["content.xml"];
	var row1 = "<table:table-row table:style-name=\"ro1\"><table:table-cell office:value-type=\"string\"><text:p>a</text:p></table:table-cell>" +
		"<table:table-cell table:style-name=\"ce2\"/>" +
		"<table:table-cell table:style-name=\"ce2\" office:value-type=\"float\" office:value=\"3.5\"><text:p>3.5</text:p></table:table-cell>" +
		"<table:table-cell table:style-name=\"ce2\" table:number-columns-repeated=\"2\"/></table:table-row>";
	assert.ok(out.indexOf(row1) > 0, out);
	assert.ok(out.indexOf("<table:table-row table:style-name=\"ro1\" table:number-rows-repeated=\"3\"><table:table-cell table:number-columns-repeated=\"5\"/></table:table-row>" +
		"<table:table-row table:style-name=\"ro1\"><table:table-cell/>" +
		"<table:table-cell table:formula=\"of:=SUM([.A1:.C1])*2\"/><table:table-cell/>" +
		"<table:table-cell office:value-type=\"string\"><text:p>hello &amp; bye</text:p></table:table-cell><table:table-cell/></table:table-row>" +
		"<table:table-row table:style-name=\"ro1\" table:number-rows-repeated=\"996\">") > 0, out);
	// 1 + (3 + 1 + 996) rows: as many as before.
});

t.test("ods patch: results replace cached values; formulas, spans and comments are kept", function() {
	var content = ODS_HEAD + "<table:table table:name=\"S\"><table:table-row>" +
		"<table:table-cell office:value-type=\"float\" office:value=\"2\" calcext:value-type=\"float\"><text:p>2</text:p></table:table-cell>" +
		"<table:table-cell table:style-name=\"ce1\" table:formula=\"of:=[.A1]*3\" office:value-type=\"float\" office:value=\"6\" calcext:value-type=\"float\">" +
		"<office:annotation><text:p>note</text:p></office:annotation><text:p>6</text:p></table:table-cell>" +
		"<table:table-cell table:number-columns-spanned=\"2\" office:value-type=\"string\"><text:p>wide</text:p></table:table-cell>" +
		"<table:covered-table-cell/></table:table-row></table:table>" + ODS_TAIL;
	var out = SheetPatch.ods({"content.xml": content}, [
		{sheet: 0, r: 0, c: 0, text: "5"},
		{sheet: 0, r: 0, c: 1, valueOnly: true, value: {v: 15, t: "n"}},
		{sheet: 0, r: 0, c: 2, text: "still wide"}
	])["content.xml"];
	assert.ok(out.indexOf("<table:table-cell table:style-name=\"ce1\" table:formula=\"of:=[.A1]*3\" office:value-type=\"float\" office:value=\"15\">" +
		"<office:annotation><text:p>note</text:p></office:annotation><text:p>15</text:p></table:table-cell>") > 0, out);
	assert.ok(out.indexOf("<table:table-cell table:number-columns-spanned=\"2\" office:value-type=\"string\"><text:p>still wide</text:p></table:table-cell><table:covered-table-cell/>") > 0, out);
});

t.test("ods: formulas in A1 syntax become OpenFormula", function() {
	var f = SheetPatch.toOpenFormula;
	assert.strictEqual(f("A1+B2"), "of:=[.A1]+[.B2]");
	assert.strictEqual(f("SUM($A$1:B10,3)"), "of:=SUM([.$A$1:.B10];3)");
	assert.strictEqual(f("Sheet2!C3*'My sheet'!A1:A4"), "of:=[$Sheet2.C3]*[$'My sheet'.A1:.A4]");
	assert.strictEqual(f("IF(A1=\"x,A1\",LOG10(100),1)"), "of:=IF([.A1]=\"x,A1\";LOG10(100);1)");
	assert.strictEqual(f("SUM(A:A)"), "of:=SUM([.A:.A])");
});

t.test("ods patch: rows past the end of the table are added", function() {
	var content = ODS_HEAD + "<table:table table:name=\"S\"><table:table-row><table:table-cell/></table:table-row></table:table>" + ODS_TAIL;
	var out = SheetPatch.ods({"content.xml": content}, [{sheet: 0, r: 3, c: 2, text: "x"},
		{sheet: 0, r: 5, c: 0, valueOnly: true, value: {v: 1, t: "n"}}])["content.xml"];
	assert.ok(out.indexOf("<table:table-row><table:table-cell/></table:table-row>" +
		"<table:table-row table:number-rows-repeated=\"2\"><table:table-cell/></table:table-row>" +
		"<table:table-row><table:table-cell table:number-columns-repeated=\"2\"/><table:table-cell office:value-type=\"string\"><text:p>x</text:p></table:table-cell></table:table-row>" +
		"</table:table>") > 0, out);
});

// --- Formulas -------------------------------------------------------------------

t.test("formulas: an edit gives new values to the formulas that use it, formatted as the cells say", function() {
	var wb = XLSX.read(new Uint8Array(buffer("sample.xlsx")), {type: "array", dense: true, cellFormula: true, cellNF: true});
	var calc = Formulas.create(HyperFormula, XLSX, wb);
	// Prices: B2 1.5, C2 4, D2 =B2*C2; D5 =SUM(D2:D4).
	var changed = calc.set(0, 1, 1, "2");
	var byRef = {};
	changed.forEach(function(c) {
		byRef[Sheets.cellRef(c.r, c.c)] = c;
	});
	assert.strictEqual(changed[0].r, 1, "the cell itself first");
	assert.strictEqual(byRef.D2.v, 8);
	assert.strictEqual(byRef.D5.v, 62.5);
	var own = calc.set(0, 2, 2, "=B3*10");
	assert.strictEqual(own[0].v, 122.5, "a new formula has its value");
	// D3 = 12.25 * 122.5; D5 = 8 + 1500.625 + 30.
	assert.strictEqual(calc.values()["0,4,3"].v, 1538.625);
	var err = calc.set(0, 3, 2, "=1/0");
	assert.strictEqual(err[0].t, "e");
	assert.strictEqual(err[0].text, "#DIV/0!");
	assert.strictEqual(calc.set(0, 0, 0, "=NOSUCHFUNCTION(1)")[0].t, "e");
	assert.ok(Formulas.contentOf({t: "s", v: "=not a formula"}) === "'=not a formula", "text that looks like a formula stays text");
	calc.destroy();
});

// --- OpenDocument paragraphs ----------------------------------------------------

t.test("text diff: the changes, in order, turn the old text into the new", function() {
	var seed = 7;
	var rnd = function(n) {
		seed = (seed * 1103515245 + 12345) % 2147483648;
		return seed % n;
	};
	var word = function() {
		var w = "";
		for (var i = 0; i < 1 + rnd(6); i++) {
			w += "abcde \t\n".charAt(rnd(8));
		}
		return w;
	};
	for (var k = 0; k < 400; k++) {
		var a = "", b = "";
		for (var j = 0; j < rnd(12); j++) {
			a += word();
		}
		b = a;
		for (var m = 0; m < rnd(4); m++) {
			var p = rnd(b.length + 1);
			b = b.slice(0, p) + (rnd(2) ? word() : "") + b.slice(p + rnd(4));
		}
		assert.strictEqual(applyDiff(a, TextEdit.diff(a, b)), b, JSON.stringify([a, b]));
	}
	// Two changes far apart stay two.
	assert.strictEqual(TextEdit.diff("Plain, bold words.", "Plain, very bold words!").length, 2);
});

t.test("text encode: spaces that would collapse, tabs and line breaks", function() {
	assert.strictEqual(TextEdit.encode("a b", "text:"), "a b");
	assert.strictEqual(TextEdit.encode("a   b", "text:"), "a <text:s text:c=\"2\"/>b");
	assert.strictEqual(TextEdit.encode(" lead", "t:"), "<t:s/>lead");
	assert.strictEqual(TextEdit.encode("a\tb\nc <&>", "text:"), "a<text:tab/>b<text:line-break/>c &lt;&amp;&gt;");
	assert.strictEqual(TextEdit.encode("x￼y\u0001z\r", "text:"), "xyz", "objects and characters XML cannot hold are dropped");
});

var ODT_HEAD = "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
	"xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" " +
	"xmlns:xlink=\"http://www.w3.org/1999/xlink\"><office:body><office:text>";
var ODT_TAIL = "</office:text></office:body></office:document-content>";

function session(inBody) {
	var xml = ODT_HEAD + inBody + ODT_TAIL;
	var track = {};
	var model = Odf.text({"content.xml": xml}, {track: track});
	return {model: model, s: TextEdit.session(xml, track), track: track};
}

//* After edits the session's tree has the offsets a fresh parse of its text gives.
function checkTree(inSession) {
	var fresh = Xml.parse(inSession.xml, {offsets: true});
	assert.deepStrictEqual(offsetsOf(inSession.root), offsetsOf(fresh));
}

t.test("odt edit: formatting of unchanged words, links, notes and bookmarks are kept", function() {
	var x = session("<text:p>Plain, <text:span text:style-name=\"B\">bold words</text:span> and a " +
		"<text:a xlink:href=\"http://x/\">link</text:a><text:bookmark text:name=\"m\"/> and a note" +
		"<text:note><text:note-citation>1</text:note-citation><text:note-body><text:p>Note.</text:p></text:note-body></text:note>.</text:p>" +
		"<text:p>Second</text:p>");
	var p0 = x.model.blocks[0];
	assert.strictEqual(p0.pt, "Plain, bold words and a link and a note￼.");
	var r = x.s.edit(p0.id, "Plain, very bold words and a link and a note￼!");
	assert.ok(x.s.xml.indexOf("<text:p>Plain, very <text:span text:style-name=\"B\">bold words</text:span> and a " +
		"<text:a xlink:href=\"http://x/\">link</text:a><text:bookmark text:name=\"m\"/> and a note<text:note>") > 0, x.s.xml);
	assert.ok(x.s.xml.indexOf("</text:note>!</text:p>") > 0, x.s.xml);
	assert.strictEqual(r.block.pt, "Plain, very bold words and a link and a note￼!");
	checkTree(x.s);
	// The note's paragraph is still editable, and the second one too: their places moved.
	var ids = [];
	(function walk(bs) {
		bs.forEach(function(b) {
			if (b.id !== undefined) {
				ids.push(b.id);
			}
			if (b.items) {
				b.items.forEach(function(i) {
					walk(i.b);
				});
			}
		});
	})(x.model.blocks);
	assert.strictEqual(ids.length, 3);
	x.s.edit(ids[1], "Second, longer");
	x.s.edit(ids[2], "Note, edited.");
	checkTree(x.s);
	assert.ok(x.s.xml.indexOf("<text:p>Note, edited.</text:p>") > 0, "a paragraph inside a note");
	assert.ok(x.s.xml.indexOf("<text:p>Second, longer</text:p>") > 0);
	assert.strictEqual(x.s.edit(ids[1], "Second, longer"), null, "the same text: nothing to do");
});

t.test("odt edit: replacing a word takes its formatting; deleting into an object removes it", function() {
	var x = session("<text:p>a <text:span text:style-name=\"I\">old</text:span> word" +
		"<draw:frame><draw:image xlink:href=\"Pictures/p.png\"/></draw:frame>end</text:p>");
	assert.strictEqual(x.model.blocks[0].pt, "a old word￼end");
	x.s.edit(0, "a new word￼end");
	assert.ok(x.s.xml.indexOf("<text:span text:style-name=\"I\">new</text:span>") > 0, x.s.xml);
	x.s.edit(0, "a new wordend");
	assert.ok(x.s.xml.indexOf("draw:frame") < 0, "the picture went with its character: " + x.s.xml);
	checkTree(x.s);
});

t.test("odt edit: empty and self-closing paragraphs, spaces, fields; repeated table rows are not editable", function() {
	var x = session("<text:p/><text:p></text:p><text:p>x <text:date>17 Oct</text:date> y</text:p>" +
		"<table:table xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"><table:table-row table:number-rows-repeated=\"2\">" +
		"<table:table-cell><text:p>rep</text:p></table:table-cell></table:table-row>" +
		"<table:table-row><table:table-cell><text:p>one</text:p></table:table-cell></table:table-row></table:table>");
	x.s.edit(0, "First  line");
	x.s.edit(1, "\tSecond");
	assert.ok(x.s.xml.indexOf("<text:p>First <text:s/>line</text:p><text:p><text:tab/>Second</text:p>") > 0, x.s.xml);
	assert.strictEqual(x.model.blocks[2].pt, "x 17 Oct y");
	x.s.edit(2, "x 17 Oc y");
	assert.ok(x.s.xml.indexOf("<text:p>x  y</text:p>") > 0 || x.s.xml.indexOf("<text:p>x <text:s/>y</text:p>") > 0,
		"a field that lost a character is removed: " + x.s.xml);
	var table = x.model.blocks[3];
	assert.strictEqual(table.rows[0].cells[0].b[0].id, undefined, "a paragraph shown twice is not editable");
	assert.ok(table.rows[2].cells[0].b[0].id !== undefined);
	checkTree(x.s);
});

// --- The engine ---------------------------------------------------------------

t.test("engine: an edited .odt is the original with only its paragraph changed", async function() {
	var e = engine();
	var r = await e.open({name: "sample.odt", buffer: buffer("sample.odt")});
	assert.ok(r.editable);
	var p = r.model.blocks.filter(function(b) {
		return b.id !== undefined && /bold words/.test(b.pt);
	})[0];
	var res = await e.edit({para: {id: p.id, text: p.pt.replace("bold words", "bold, edited words")}});
	assert.ok(res.block);
	var bytes = await e.save({format: "odt"});
	var before = await entries(new Uint8Array(buffer("sample.odt")));
	var after = await entries(bytes);
	assert.deepStrictEqual(Object.keys(after).sort(), Object.keys(before).sort());
	Object.keys(before).forEach(function(k) {
		if (k !== "content.xml") {
			assert.strictEqual(after[k], before[k], k + " unchanged");
		}
	});
	assert.ok(/bold, edited words/.test(after["content.xml"]));
	// The mimetype first and stored, as OpenDocument asks.
	var b = Buffer.from(bytes);
	assert.strictEqual(b.readUInt32LE(0), 0x04034b50);
	assert.strictEqual(b.readUInt16LE(8), 0, "stored");
	assert.strictEqual(b.toString("latin1", 30, 30 + b.readUInt16LE(26)), "mimetype");
	// It opens again, with the edit.
	var e2 = engine();
	var r2 = await e2.open({name: "copy.odt", buffer: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)});
	assert.ok(r2.model.blocks.some(function(x) {
		return x.pt && /bold, edited words/.test(x.pt);
	}));
});

[["sample.xlsx", "xlsx"], ["sample.ods", "ods"]].forEach(function(inCase) {
	t.test("engine: an edited ." + inCase[1] + " keeps the rest of the file and has the new results", async function() {
		var e = engine();
		await e.open({name: inCase[0], buffer: buffer(inCase[0])});
		var shown = await e.edit({cell: {sheet: 0, r: 1, c: 1, text: "2"}});
		assert.ok(shown.cells.some(function(c) {
			return c.r === 1 && c.c === 3 && c.text === "8";
		}), JSON.stringify(shown.cells));
		await e.edit({cell: {sheet: 0, r: 7, c: 0, text: "Added"}});
		var bytes = await e.save({format: inCase[1], edits: [{sheet: 0, r: 1, c: 1, text: "2"}, {sheet: 0, r: 7, c: 0, text: "Added"}]});
		var before = await entries(new Uint8Array(buffer(inCase[0])));
		var after = await entries(bytes);
		var changed = Object.keys(before).filter(function(k) {
			return after[k] !== before[k];
		}).sort();
		assert.deepStrictEqual(changed, inCase[1] === "xlsx" ? ["xl/workbook.xml", "xl/worksheets/sheet1.xml"] : ["content.xml"]);
		var wb = XLSX.read(Buffer.from(bytes), {type: "buffer", cellFormula: true});
		var ws = wb.Sheets.Prices;
		assert.strictEqual(ws.B2.v, 2);
		assert.strictEqual(ws.D2.v, 8, "the formula's cached result is the new one");
		assert.strictEqual(ws.D5.v, 62.5);
		assert.strictEqual(ws.A8.v, "Added");
		assert.strictEqual(ws.B3.v, 12.25, "the rest is kept");
		assert.ok(ws.D2.f, "formulas are kept");
	});
});

t.test("engine: an .xls is saved as .xlsx through SheetJS, with the results", async function() {
	var e = engine();
	await e.open({name: "sample.xls", buffer: buffer("sample.xls")});
	await e.edit({cell: {sheet: 0, r: 1, c: 1, text: "2"}});
	var bytes = await e.save({format: "xlsx", edits: [{sheet: 0, r: 1, c: 1, text: "2"}]});
	var ws = XLSX.read(Buffer.from(bytes), {type: "buffer"}).Sheets.Prices;
	assert.strictEqual(ws.B2.v, 2);
	assert.strictEqual(ws.D2.v, 8);
});

t.test("engine: plain text and Markdown are saved as UTF-8", async function() {
	var e = engine();
	var bytes = await e.save({format: "md", text: "# Title\nété"});
	assert.strictEqual(Buffer.from(bytes).toString("utf8"), "# Title\nété");
});

// --- Slides ------------------------------------------------------------------

var A_NS = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";

//* A DrawingML paragraph inside a text body, parsed with offsets: {xml, p}.
function drawingPara(inInner) {
	var xml = "<a:txBody " + A_NS + "><a:p>" + inInner + "</a:p></a:txBody>";
	return {xml: xml, p: Xml.child(Xml.parse(xml, {offsets: true}), "a:p")};
}

function editDrawing(inInner, inText) {
	var d = drawingPara(inInner);
	var list = TextEdit.splices(d.xml, d.p, inText, TextEdit.drawingDialect);
	var out = list ? Xml.splice(d.xml, list) : d.xml;
	Xml.parse(out);
	return out.replace(/^<a:txBody[^>]*><a:p>/, "").replace(/<\/a:p><\/a:txBody>$/, "");
}

t.test("pptx paragraph: words keep their run, a new line is a break and a run like it", function() {
	var bold = "<a:r><a:rPr b=\"1\"/><a:t>bold</a:t></a:r>";
	var inner = "<a:r><a:rPr lang=\"en-US\"/><a:t>Plain and </a:t></a:r>" + bold + "<a:r><a:rPr lang=\"en-US\"/><a:t> end.</a:t></a:r>" +
		"<a:endParaRPr lang=\"en-US\" sz=\"1800\"/>";
	assert.strictEqual(Pptx.plainText(Pptx.tokens(drawingPara(inner).p)), "Plain and bold end.");
	assert.strictEqual(editDrawing(inner, "Plain and bold end!"),
		"<a:r><a:rPr lang=\"en-US\"/><a:t>Plain and </a:t></a:r>" + bold + "<a:r><a:rPr lang=\"en-US\"/><a:t> end!</a:t></a:r><a:endParaRPr lang=\"en-US\" sz=\"1800\"/>");
	assert.strictEqual(editDrawing(inner, "Plain and very bold end."),
		"<a:r><a:rPr lang=\"en-US\"/><a:t>Plain and very </a:t></a:r>" + bold + "<a:r><a:rPr lang=\"en-US\"/><a:t> end.</a:t></a:r><a:endParaRPr lang=\"en-US\" sz=\"1800\"/>");
	assert.strictEqual(editDrawing(inner, "Plain and bo\nld end."),
		"<a:r><a:rPr lang=\"en-US\"/><a:t>Plain and </a:t></a:r><a:r><a:rPr b=\"1\"/><a:t>bo</a:t></a:r><a:br/><a:r><a:rPr b=\"1\"/><a:t>ld</a:t></a:r>" +
		"<a:r><a:rPr lang=\"en-US\"/><a:t> end.</a:t></a:r><a:endParaRPr lang=\"en-US\" sz=\"1800\"/>",
		"the line after the break keeps the run's properties");
	assert.strictEqual(editDrawing(inner, "Plain and bold end."), "<a:r><a:rPr lang=\"en-US\"/><a:t>Plain and </a:t></a:r>" + bold +
		"<a:r><a:rPr lang=\"en-US\"/><a:t> end.</a:t></a:r><a:endParaRPr lang=\"en-US\" sz=\"1800\"/>", "the same text: nothing changes");
});

t.test("pptx paragraph: breaks, fields, an empty paragraph and escaping", function() {
	var inner = "<a:r><a:t>one</a:t></a:r><a:br><a:rPr/></a:br><a:r><a:t>two</a:t></a:r>" +
		"<a:fld id=\"{1}\" type=\"slidenum\"><a:t>7</a:t></a:fld><a:endParaRPr sz=\"2000\"/>";
	assert.strictEqual(Pptx.plainText(Pptx.tokens(drawingPara(inner).p)), "one\ntwo7");
	assert.strictEqual(editDrawing(inner, "onetwo7"), "<a:r><a:t>one</a:t></a:r><a:r><a:t>two</a:t></a:r>" +
		"<a:fld id=\"{1}\" type=\"slidenum\"><a:t>7</a:t></a:fld><a:endParaRPr sz=\"2000\"/>", "a deleted break goes");
	assert.strictEqual(editDrawing(inner, "one\ntwo"), "<a:r><a:t>one</a:t></a:r><a:br><a:rPr/></a:br><a:r><a:t>two</a:t></a:r>" +
		"<a:endParaRPr sz=\"2000\"/>", "a field that lost its text goes");
	assert.strictEqual(editDrawing("<a:endParaRPr lang=\"de-DE\" sz=\"2000\"><a:solidFill/></a:endParaRPr>", "New <b> & more\nline"),
		"<a:r><a:rPr lang=\"de-DE\" sz=\"2000\"><a:solidFill/></a:rPr><a:t>New &lt;b&gt; &amp; more</a:t></a:r><a:br/>" +
		"<a:r><a:rPr lang=\"de-DE\" sz=\"2000\"><a:solidFill/></a:rPr><a:t>line</a:t></a:r>" +
		"<a:endParaRPr lang=\"de-DE\" sz=\"2000\"><a:solidFill/></a:endParaRPr>", "an empty paragraph gets runs formatted as its end");
});

t.test("engine: an edited .pptx is the original with only the slide changed; the slide is read again", async function() {
	var e = engine();
	var r = await e.open({name: "sample.pptx", buffer: buffer("sample.pptx")});
	assert.ok(r.editable);
	var ids = [];
	r.model.slides.forEach(function(s, si) {
		s.items.forEach(function(it) {
			(it.b || []).forEach(function(b) {
				if (b.id !== undefined) {
					ids.push({slide: si, id: b.id, pt: b.pt});
				}
			});
		});
	});
	assert.deepStrictEqual(ids.map(function(x) {
		return x.pt;
	}), ["Fixture slides", "First point", "Second point", "Yellow box", "findme on the last slide"], "the slides' own text, not the master's");
	var res = await e.edit({para: {id: 1, text: "First point, edited"}});
	assert.strictEqual(res.slide.index, 0);
	assert.ok(JSON.stringify(res.slide.model).indexOf("First point, edited") > 0);
	var again = await e.edit({para: {id: 4, text: "findme on the very last slide"}});
	assert.strictEqual(again.slide.index, 2, "ids still the same after the deck was read again");
	var bytes = await e.save({format: "pptx"});
	var before = await entries(new Uint8Array(buffer("sample.pptx")));
	var after = await entries(bytes);
	var changed = Object.keys(before).filter(function(k) {
		return after[k] !== before[k];
	}).sort();
	assert.deepStrictEqual(changed, ["ppt/slides/slide1.xml", "ppt/slides/slide3.xml"]);
	var e2 = engine();
	var r2 = await e2.open({name: "copy.pptx", buffer: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)});
	assert.ok(JSON.stringify(r2.model).indexOf("First point, edited") > 0);
});

t.test("engine: text a slide's master draws is shown but not editable", async function() {
	var zip = await JSZip.loadAsync(buffer("sample.pptx"));
	var path = "ppt/slideMasters/slideMaster1.xml";
	var master = await zip.file(path).async("string");
	var box = "<p:sp><p:nvSpPr><p:cNvPr id=\"99\" name=\"Footer box\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>" +
		"<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"3000000\" cy=\"400000\"/></a:xfrm><a:prstGeom prst=\"rect\"/></p:spPr>" +
		"<p:txBody><a:bodyPr/><a:p><a:r><a:t>Master footer</a:t></a:r></a:p></p:txBody></p:sp>";
	zip.file(path, master.replace("</p:spTree>", box + "</p:spTree>"));
	var bytes = await zip.generateAsync({type: "uint8array"});
	var e = engine();
	var r = await e.open({name: "m.pptx", buffer: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)});
	var found = [];
	r.model.slides[0].items.forEach(function(it) {
		(it.b || []).forEach(function(b) {
			if (b.r && JSON.stringify(b.r).indexOf("Master footer") >= 0) {
				found.push(b);
			}
		});
	});
	assert.ok(found.length >= 1, "the master's text is drawn on the slide");
	found.forEach(function(b) {
		assert.strictEqual(b.id, undefined, "but it is the master's, not the slide's: not editable");
	});
});

t.test("engine: an edited .odp is the original with only content.xml changed", async function() {
	var e = engine();
	var r = await e.open({name: "sample.odp", buffer: buffer("sample.odp")});
	assert.ok(r.editable);
	var first = null;
	(function walk(bs) {
		(bs || []).forEach(function(b) {
			if (!first && b.id !== undefined && /Yellow/.test(b.pt)) {
				first = b;
			}
			(b.items || []).forEach(walk);
		});
	})([].concat.apply([], r.model.slides.map(function(s) {
		return [].concat.apply([], s.items.map(function(it) {
			return it.b || [];
		}));
	})));
	assert.ok(first, "the box's paragraph is editable");
	var res = await e.edit({para: {id: first.id, text: "Yellow box, edited"}});
	assert.strictEqual(res.block.pt, "Yellow box, edited");
	var bytes = await e.save({format: "odp"});
	var before = await entries(new Uint8Array(buffer("sample.odp")));
	var after = await entries(bytes);
	assert.deepStrictEqual(Object.keys(before).filter(function(k) {
		return after[k] !== before[k];
	}), ["content.xml"]);
	assert.ok(/Yellow box, edited/.test(after["content.xml"]));
});

// --- Word copies from the .docx editor ------------------------------------------

async function docxWith(inCustom) {
	var zip = await JSZip.loadAsync(buffer("sample.docx"));
	if (inCustom) {
		zip.file("docProps/custom.xml", inCustom);
		var rels = await zip.file("_rels/.rels").async("string");
		zip.file("_rels/.rels", rels.replace("</Relationships>", "<Relationship Id=\"rIdC\" " +
			"Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties\" Target=\"docProps/custom.xml\"/></Relationships>"));
		var types = await zip.file("[Content_Types].xml").async("string");
		zip.file("[Content_Types].xml", types.replace("</Types>", "<Override PartName=\"/docProps/custom.xml\" " +
			"ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/></Types>"));
	}
	return zip.generateAsync({type: "uint8array"});
}

function customXml(inProps) {
	return "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" " +
		"xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" + inProps.map(function(p, i) {
		return "<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"" + (i + 2) + "\" name=\"" + p + "\"><vt:lpwstr>x</vt:lpwstr></property>";
	}).join("") + "</Properties>";
}

t.test("docx copy: the editor's stamps are removed, the document's own properties kept", async function() {
	var stamped = customXml(["SuperdocVersion", "SuperdocDocumentOrigin", "DocumentGuid"]);
	// A document without custom properties: the part, its relationship and its content type go.
	var out = await entries(await DocxCopy.clean(JSZip, await docxWith(stamped), await docxWith(null)));
	assert.strictEqual(out["docProps/custom.xml"], undefined);
	assert.ok(!/custom\.xml/.test(out["_rels/.rels"]), out["_rels/.rels"]);
	assert.ok(!/custom\.xml/.test(out["[Content_Types].xml"]));
	assert.ok(/Fixture heading/.test(out["word/document.xml"]), "the document itself is untouched");
	// One with a property of its own keeps it.
	var own = customXml(["Client", "SuperdocVersion", "DocumentGuid"]);
	var out2 = await entries(await DocxCopy.clean(JSZip, await docxWith(own), await docxWith(customXml(["Client"]))));
	assert.ok(/name="Client"/.test(out2["docProps/custom.xml"]));
	assert.ok(!/Superdoc|DocumentGuid/.test(out2["docProps/custom.xml"]), out2["docProps/custom.xml"]);
	assert.ok(/custom\.xml/.test(out2["_rels/.rels"]));
});

t.run();
