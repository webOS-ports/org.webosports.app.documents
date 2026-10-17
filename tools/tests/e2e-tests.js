/*
	End-to-end tests: the real app (index-desktop.html, Enyo 1 from
	enyo-1.0-ports) in headless Chromium, with the webOS runtime, the luna
	bus and the file system mocked (mock/webos-mock.js). Files under
	/media/internal are served from fixtures/ (and test-output/e2e-root/
	for documents made here), through the app's own file: reads.

		node tools/tests/e2e-tests.js [--screenshots]
		TEST_FILTER=sheet node tools/tests/e2e-tests.js

	Every test also fails on a page error, a console.error, an unmocked
	request, or a luna call the app's requiredPermissions would not allow.
*/
var fs = require("fs");
var path = require("path");
var t = require("./lib/harness.js");
var browserLib = require("./lib/browser.js");
var serverLib = require("./lib/server.js");
var assert = t.assert;

var ROOT = path.join(__dirname, "..", "..");
var FIX = path.join(__dirname, "fixtures");
var OUT = path.join(ROOT, "test-output");
var E2E_ROOT = path.join(OUT, "e2e-root");
var SHOTS = process.argv.indexOf("--screenshots") >= 0;
var XLSX = require(path.join(ROOT, "lib/sheetjs/xlsx.full.min.js"));
var mockSource = fs.readFileSync(path.join(__dirname, "mock", "webos-mock.js"), "utf8");
var appinfo = JSON.parse(fs.readFileSync(path.join(ROOT, "appinfo.json"), "utf8"));

var server, browser, page;

// --- Documents made for the tests --------------------------------------------

function makeE2eRoot() {
	fs.rmSync(E2E_ROOT, {recursive: true, force: true});
	fs.mkdirSync(path.join(E2E_ROOT, "Work"), {recursive: true});
	// A big sheet: 20000 rows, for the grid's virtualisation.
	var lines = ["Id,Name,Value"];
	for (var i = 1; i <= 20000; i++) {
		lines.push(i + ",name " + i + "," + (i * 3));
	}
	fs.writeFileSync(path.join(E2E_ROOT, "big.csv"), lines.join("\n") + "\n");
	// A name with a comma (the File Manager's service joins names with commas).
	fs.copyFileSync(path.join(FIX, "sample.docx"), path.join(E2E_ROOT, "Work", "Report, final.docx"));
	// A binary Word 97 file: only its header, enough to be recognised.
	var ole = Buffer.alloc(2048);
	Buffer.from([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]).copy(ole);
	fs.writeFileSync(path.join(E2E_ROOT, "legacy.doc"), ole);
	// PDFs, for the PDF app: one by its name, one behind a ".docx" name.
	var pdf = Buffer.from("%PDF-1.4\n1 0 obj << >> endobj\ntrailer << >>\n%%EOF\n");
	fs.writeFileSync(path.join(E2E_ROOT, "Manual 1.pdf"), pdf);
	fs.writeFileSync(path.join(E2E_ROOT, "scan.docx"), pdf);
}

function fixtureFiles() {
	return fs.readdirSync(FIX).filter(function(f) {
		return fs.statSync(path.join(FIX, f)).isFile();
	});
}

// --- The mocked device --------------------------------------------------------

function acg() {
	var groups = {};
	fs.readdirSync(path.join(FIX, "acg")).forEach(function(f) {
		var g = JSON.parse(fs.readFileSync(path.join(FIX, "acg", f), "utf8"));
		for (var k in g) {
			if (g.hasOwnProperty(k)) {
				groups[k] = (groups[k] || []).concat(g[k]);
			}
		}
	});
	// The installer gives an app the group of its own service.
	groups["org.webosports.app.documents.service.group"] = ["org.webosports.app.documents.service/*"];
	return {groups: groups, granted: appinfo.requiredPermissions.concat(["org.webosports.app.documents.service.group"]),
		// Enyo 1's test hook, which it asks for at start on every device (and is told is not there).
		exempt: ["^luna://com\\.palm\\.telluriumservice/"]};
}

/**
	A LuneOS device with the fixtures on its USB storage. inOptions:
	launchParams, service (true: the app's service is installed), luna (more
	rules), storage, palmSystem.
*/
function scenario(inOptions) {
	var o = inOptions || {};
	var files = fixtureFiles().concat(["big.csv", "Manual 1.pdf", "scan.docx", "legacy.doc"]).join(",");
	var luna = {
		"luna://org.webosports.service.filemanager/GetDirectories": {cases: [
			{when: {dir: "/media/internal"}, json: {dirs: "Work,.thumbnails,src", files: files}},
			{when: {dir: "/media/internal/Work"}, json: {dirs: "", files: "Report, final.docx,todo.txt.bak"}},
			{when: {dir: "/media/internal/src"}, json: {dirs: "", files: "sample.fodt,sample.fods,sample.fodp"}}
		], error: {errorCode: -1, errorText: "ENOENT"}},
		"luna://org.webosports.service.filemanager/getFileSize": {cases: [
			{when: {path: "/media/internal/Work/Report, final.docx"}, json: {size: 8404}}
		], error: {errorCode: -1, errorText: "ENOENT"}},
		"luna://com.palm.applicationManager/open": {json: {}},
		"luna://com.palm.telluriumservice/subscribeToCommands": {error: {errorCode: -1, errorText: "Service does not exist"}},
		"luna://org.webosports.app.documents.service/info": o.service ? {json: {id: "org.webosports.app.documents.service"}} :
			{error: {errorCode: -1, errorText: "Service does not exist: org.webosports.app.documents.service."}},
		"luna://org.webosports.app.documents.service/writeNew": o.service ? {json: {path: "x", bytes: 1}} :
			{error: {errorCode: -1, errorText: "Service does not exist"}},
		// LibreOffice's service: asked without a document it says badRequest; legacy.doc becomes sample.odt.
		"luna://org.webosports.service.office/convert": o.office ? {cases: [
			{when: {source: "/media/internal/legacy.doc", format: "odt"}, json: {target: "/media/internal/sample.odt", format: "odt", cached: false}}
		], error: {errorCode: "badRequest", errorText: "source and format are required"}} :
			{error: {errorCode: -1, errorText: "Service does not exist: org.webosports.service.office."}}
	};
	Object.assign(luna, o.luna || {});
	return {
		webos: true,
		palmSystem: Object.assign({identifier: "org.webosports.app.documents 1000", launchParams: JSON.stringify(o.launchParams || {})},
			o.palmSystem || {}),
		luna: luna,
		acg: acg(),
		rewrite: [
			{from: "^file:///media/internal/(big\\.csv|legacy\\.doc|Manual(%20| )1\\.pdf|scan\\.docx|Work/)", to: server.url + "/test-output/e2e-root/$1"},
			{from: "^file:///media/internal/", to: server.url + "/tools/tests/fixtures/"}
		],
		http: [],
		storage: o.storage || {clear: true}
	};
}

async function open(inConfig, inSize) {
	if (page) {
		await page.close();
	}
	var size = inSize || {width: 1024, height: 768};
	page = await browser.newPage({width: size.width, height: size.height,
		initScripts: ["window.__MOCK_CONFIG__ = " + JSON.stringify(inConfig) + ";", mockSource]});
	await page.goto(server.url + "/index-desktop.html");
	await page.waitFor("window.documentsApp && documentsApp.hasNode()");
	return page;
}

function calls(inKind, inPattern) {
	return page.eval("__MOCK__.calls.filter(function(c){return c.kind===" + JSON.stringify(inKind) + " && new RegExp(" +
		JSON.stringify(inPattern) + ").test(c.url);})");
}

//* Waits until the viewer has shown the document (or an error).
async function waitOpened() {
	return page.waitFor("(function(){var v=documentsApp.$.viewer; return (v.doc && v.timings.total && !v.renderingAll) ? v.doc.kind : " +
		"(v.lastError && !v.doc ? 'error' : null);})()", 20000);
}

async function openDoc(inPath) {
	await page.eval("documentsApp.openDocument(null, " + JSON.stringify(inPath) + "); 1");
	return waitOpened();
}

function text(inSelector) {
	return page.eval("(function(){var e=document.querySelector(" + JSON.stringify(inSelector) + ");return e?e.textContent:null;})()");
}

function count(inSelector) {
	return page.eval("document.querySelectorAll(" + JSON.stringify(inSelector) + ").length");
}

async function libraryNames() {
	await page.waitFor("!documentsApp.$.library.scanning", 10000);
	return page.eval("Array.prototype.map.call(document.querySelectorAll('.library-name'), function(e){return e.textContent;})");
}

//* A real tap on an Enyo control, found by its path in the app (e.g. "documentsApp.$.viewer.$.findNext").
async function tapControl(inPath) {
	var c = await page.waitFor("(function(){var n=" + inPath + ".hasNode();if(!n)return null;var r=n.getBoundingClientRect();" +
		"return r.width&&r.height?{x:r.left+r.width/2,y:r.top+r.height/2}:null;})()", 5000);
	await page.send("Input.dispatchMouseEvent", {type: "mousePressed", x: c.x, y: c.y, button: "left", clickCount: 1});
	await page.send("Input.dispatchMouseEvent", {type: "mouseReleased", x: c.x, y: c.y, button: "left", clickCount: 1});
	await browserLib.sleep(100);
}

async function key(inKey) {
	await page.key(inKey);
	await browserLib.sleep(150);
}

// --- Library ---------------------------------------------------------------------

t.test("library: every document on the device, hidden folders skipped, others ignored", async function() {
	await open(scenario());
	var names = await libraryNames();
	["sample.docx", "sample.odt", "sample.xlsx", "sample.xls", "sample.ods", "sample.pptx", "sample.odp", "sample.md",
		"sample.txt", "semicolon.csv", "big.csv", "sample.fodt"].forEach(function(n) {
		assert.ok(names.indexOf(n) >= 0, n + " listed: " + names.join(", "));
	});
	assert.ok(names.indexOf("legacy.doc") < 0, "a .doc is not offered");
	assert.ok(names.indexOf("todo.txt.bak") < 0);
	assert.ok(!(await calls("luna", "GetDirectories")).some(function(c) {
		return /thumbnails/.test(c.params.dir);
	}), "hidden folders are not walked");
});

t.test("library: a name with a comma is rebuilt and checked", async function() {
	await open(scenario());
	var names = await libraryNames();
	assert.ok(names.indexOf("Report, final.docx") >= 0, names.join(", "));
	assert.strictEqual((await calls("luna", "getFileSize")).length, 1);
});

t.test("library: filter by kind and search by name", async function() {
	await open(scenario());
	await libraryNames();
	await page.tap(".library-filters .enyo-grouped-toolbutton:nth-child(3)");
	var sheets = await page.eval("Array.prototype.map.call(document.querySelectorAll('.library-name'), function(e){return e.textContent;})");
	assert.ok(sheets.length >= 5 && sheets.every(function(n) {
		return /\.(xlsx|xls|ods|csv|fods)$/.test(n);
	}), sheets.join(", "));
	await page.tap(".library-filters .enyo-grouped-toolbutton:nth-child(1)");
	await page.eval("documentsApp.$.library.$.search.setValue('semi'); documentsApp.$.library.searchChanged(); 1");
	await page.waitFor("document.querySelectorAll('.library-name').length === 1");
	assert.strictEqual(await text(".library-name"), "semicolon.csv");
});

t.test("library: tapping a document opens it; it is then listed as recent", async function() {
	await open(scenario());
	await libraryNames();
	await page.eval("documentsApp.$.library.$.search.setValue('sample.odt'); documentsApp.$.library.searchChanged(); 1");
	await page.waitFor("document.querySelectorAll('.library-name').length === 1");
	await page.tap(".library-item");
	assert.strictEqual(await waitOpened(), "odt");
	await key("Escape");
	await page.waitFor("documentsApp.$.pane.getViewName() === 'library'");
	await page.eval("documentsApp.$.library.$.search.setValue(''); documentsApp.$.library.searchChanged(); 1");
	await page.waitFor("document.querySelector('.library-divider') && /Recent/.test(document.querySelector('.library-divider').textContent)");
	assert.strictEqual(await text(".library-item .library-name"), "sample.odt");
	var stored = await page.eval("JSON.parse(localStorage.getItem('documents.prefs')).recent[0].path");
	assert.strictEqual(stored, "/media/internal/sample.odt");
});

t.test("library: a PDF is listed, and a tap hands it to the application/pdf app", async function() {
	await open(scenario());
	await libraryNames();
	await page.eval("documentsApp.$.library.$.search.setValue('manual'); documentsApp.$.library.searchChanged(); 1");
	await page.waitFor("document.querySelectorAll('.library-name').length === 1");
	assert.strictEqual(await text(".library-name"), "Manual 1.pdf");
	assert.strictEqual(await count(".library-icon-pdf"), 1);
	await page.tap(".library-item");
	await page.waitFor("__MOCK__.calls.some(function(c){return /applicationManager\\/open/.test(c.url);})");
	var c = await calls("luna", "applicationManager/open");
	assert.strictEqual(c.length, 1);
	assert.strictEqual(c[0].params.target, "file:///media/internal/Manual%201.pdf");
	assert.strictEqual(c[0].params.mime, "application/pdf");
	assert.strictEqual(c[0].acg, "granted");
	assert.strictEqual(await page.eval("documentsApp.$.pane.getViewName()"), "library", "Documents stays on its library");
	assert.strictEqual((await calls("file", "Manual")).length, 0, "the PDF is not read here");
});

t.test("opening: a PDF behind a .docx name is handed on too", async function() {
	await open(scenario());
	await page.eval("documentsApp.openDocument(null, '/media/internal/scan.docx'); 1");
	await page.waitFor("__MOCK__.calls.some(function(c){return /applicationManager\\/open/.test(c.url);})");
	var c = await calls("luna", "applicationManager/open");
	assert.strictEqual(c[0].params.target, "file:///media/internal/scan.docx");
	assert.strictEqual(c[0].params.mime, "application/pdf");
	await page.waitFor("documentsApp.$.pane.getViewName() === 'library'");
});

t.test("opening: the application manager's empty failure after a launch counts as handed on", async function() {
	await open(scenario({luna: {"luna://com.palm.applicationManager/open": {error: {errorText: ""}}}}));
	await page.eval("documentsApp.openDocument(null, '/media/internal/Manual 1.pdf'); 1");
	await page.waitFor("__MOCK__.calls.some(function(c){return /applicationManager\\/open/.test(c.url);})");
	await page.eval("new Promise(function(r){setTimeout(r, 300);})");
	assert.strictEqual(await page.eval("documentsApp.$.pane.getViewName()"), "library");
	assert.strictEqual((await calls("file", "Manual")).length, 0);
});

t.test("opening: with no app for PDFs, the viewer says so", async function() {
	await open(scenario({luna: {"luna://com.palm.applicationManager/open":
		{error: {errorText: "No handler for file:///media/internal/Manual%201.pdf"}}}}));
	assert.strictEqual(await openDoc("/media/internal/Manual 1.pdf"), "error");
	assert.ok(/no app to show PDF/.test(await text(".viewer-status-text")));
	assert.strictEqual(await openDoc("/media/internal/scan.docx"), "error");
	assert.ok(/no app to show PDF/.test(await text(".viewer-status-text")));
});

// --- Opening -----------------------------------------------------------------------

t.test("launch: {target: file URL} opens that document straight away", async function() {
	await open(scenario({launchParams: {target: "file:///media/internal/sample.docx"}}));
	assert.strictEqual(await waitOpened(), "docx");
	await page.waitFor("document.querySelector('.doc-docx section.docx')");
	assert.ok(/Fixture heading/.test(await text(".doc-docx")));
	assert.strictEqual(await page.eval("documentsApp.$.pane.getViewName()"), "viewer");
});

t.test("launch: a relaunch with another target opens it in the open card", async function() {
	await open(scenario({launchParams: {target: "/media/internal/sample.txt"}}));
	assert.strictEqual(await waitOpened(), "txt");
	await page.eval("__MOCK__.relaunch({target: 'file:///media/internal/sample.md'}); 1");
	await page.waitFor("documentsApp.$.viewer.doc && documentsApp.$.viewer.doc.kind === 'md'");
	assert.ok(/Markdown fixture/.test(await text(".doc-markdown-body h1")));
});

t.test("odt: text, styles, list, table and picture on the page", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.odt");
	assert.strictEqual(await text(".doc-page h1"), "Fixture heading");
	assert.strictEqual(await count(".doc-page ol"), 2);
	assert.strictEqual(await page.eval("getComputedStyle(document.querySelector('.doc-page ol')).listStyleType"), "decimal");
	assert.strictEqual(await page.eval("document.querySelector('.doc-table td[colspan=\"2\"]').textContent"), "Spanning two columns");
	await page.waitFor("(function(){var i=document.querySelector('img.doc-img');return i && i.complete && i.naturalWidth === 40;})()");
	assert.ok(/^blob:/.test(await page.eval("document.querySelector('img.doc-img').src")));
	assert.strictEqual(await page.eval("document.querySelector('.doc-link').hasAttribute('href')"), false, "links never navigate");
	assert.ok(/The footnote text/.test(await text(".doc-notes")));
});

t.test("docx, pptx, odp, md, txt, csv, xls, flat odf: each opens and shows its content", async function() {
	await open(scenario());
	var cases = [
		["/media/internal/sample.docx", "docx", ".doc-docx", /Spanning two columns/],
		["/media/internal/sample.pptx", "pptx", ".doc-slides", /Fixture slides[\s\S]*Yellow box[\s\S]*findme on the last slide/],
		["/media/internal/sample.odp", "odp", ".doc-slides", /Fixture slides[\s\S]*Yellow box/],
		["/media/internal/sample.md", "md", ".doc-markdown-body", /Markdown fixture/],
		["/media/internal/sample.txt", "txt", ".doc-plaintext", /Second line with findme/],
		["/media/internal/semicolon.csv", "csv", ".sheet-cells", /Smith; J\./],
		["/media/internal/sample.xls", "xls", ".sheet-cells", /Book/],
		["/media/internal/src/sample.fodt", "odt", ".doc-page", /Fixture heading/],
		["/media/internal/src/sample.fods", "ods", ".sheet-cells", /Lamp/]
	];
	for (var i = 0; i < cases.length; i++) {
		var c = cases[i];
		assert.strictEqual(await openDoc(c[0]), c[1], c[0]);
		assert.ok(c[3].test(await text(c[2])), c[0] + ": " + (await text(c[2])));
	}
	assert.strictEqual(await count(".doc-slide-frame"), 0, "the last document replaced the slides");
});

t.test("slides: three slides, scaled to the width, with the picture", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.pptx");
	assert.strictEqual(await count(".doc-slide-frame"), 3);
	var w = await page.eval("document.querySelector('.doc-slide-frame').getBoundingClientRect().width");
	assert.ok(w > 900 && w <= 1024, "fitted: " + w);
	await page.waitFor("(function(){var i=document.querySelector('.doc-slide-image img');return i && i.complete && i.naturalWidth===64;})()");
	assert.strictEqual(await page.eval("document.querySelector('.doc-bullet').textContent"), "•");
});

t.test("markdown: raw HTML is shown as text and never runs", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.md");
	assert.strictEqual(await count(".doc-markdown-body script"), 0);
	assert.ok(/<script>alert\("not run"\)<\/script>/.test(await text(".doc-markdown-body p")));
	assert.strictEqual(await page.eval("document.querySelector('.doc-markdown-body a').hasAttribute('href')"), false);
});

t.test("errors: a binary Word 97 file and a missing file say what is wrong", async function() {
	await open(scenario());
	assert.strictEqual(await openDoc("/media/internal/legacy.doc"), "error");
	assert.ok(/before 2007/.test(await text(".viewer-status-text")));
	assert.strictEqual(await openDoc("/media/internal/nothere.docx"), "error");
	assert.ok(/could not be read/.test(await text(".viewer-status-text")));
	// The page logs nothing as an error for this: the message is the answer.
});

t.test("links: a tap asks first, then opens the browser through luna", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.odt");
	await page.tap(".doc-link");
	await page.waitFor("documentsApp.$.viewer.$.linkDialog.isOpen");
	assert.ok(/webos-ports\.org/.test(await page.eval("documentsApp.$.viewer.$.linkText.content")));
	assert.strictEqual((await calls("luna", "applicationManager/open")).length, 0, "nothing opened before the answer");
	await page.eval("documentsApp.$.viewer.openLink(); 1");
	var c = await calls("luna", "applicationManager/open");
	assert.strictEqual(c.length, 1);
	assert.strictEqual(c[0].params.target, "https://www.webos-ports.org/");
	assert.strictEqual(c[0].acg, "granted");
});

// --- Reading ----------------------------------------------------------------------

t.test("find: every match marked, accents and case ignored, next and previous", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.odt");
	await page.tap(".viewer-toolbar .viewer-tool:nth-child(5)");
	await page.waitFor("documentsApp.$.viewer.$.findBar.showing");
	await page.type("findme");
	await page.waitFor("document.querySelectorAll('mark.find-hit').length === 3", 5000);
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.findCount.content"), "1 of 3");
	assert.strictEqual(await text("mark.find-current"), "findme");
	await tapControl("documentsApp.$.viewer.$.findNext");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.findCount.content"), "2 of 3");
	assert.strictEqual(await text("mark.find-current"), "FINDME");
	await page.eval("documentsApp.$.viewer.findNext(); 1");
	assert.strictEqual(await text("mark.find-current"), "Findmé");
	await key("Escape");
	assert.strictEqual(await count("mark.find-hit"), 0, "back closes find and clears the marks");
	assert.strictEqual(await page.eval("documentsApp.$.pane.getViewName()"), "viewer", "and stays in the document");
	await key("Escape");
	await page.waitFor("documentsApp.$.pane.getViewName() === 'library'");
});

t.test("zoom: in, out, and back to fit", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.odt");
	var fit = await page.eval("documentsApp.$.viewer.zoom");
	await page.tap(".viewer-toolbar .viewer-tool:nth-child(3)");
	var bigger = await page.eval("documentsApp.$.viewer.zoom");
	assert.ok(bigger > fit, bigger + " > " + fit);
	assert.strictEqual(await page.eval("document.querySelector('.viewer-content').style.zoom"), String(bigger));
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.zoomLabel.caption"), Math.round(bigger * 100) + "%");
	await page.tap(".viewer-toolbar .viewer-tool:nth-child(1)");
	await page.tap(".viewer-toolbar .viewer-tool:nth-child(1)");
	assert.ok((await page.eval("documentsApp.$.viewer.zoom")) < bigger);
	await page.tap(".viewer-toolbar .viewer-tool:nth-child(2)");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.zoom"), fit);
});

t.test("night reading: dark pages, kept for the next start", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.odt");
	await page.tap(".viewer-toolbar .viewer-tool:nth-child(6)");
	assert.ok(await page.eval("documentsApp.$.viewer.hasClass('night')"));
	var filter = await page.eval("getComputedStyle(document.querySelector('.doc-page')).filter");
	assert.ok(/invert/.test(filter), filter);
	var imgFilter = await page.eval("getComputedStyle(document.querySelector('img.doc-img')).filter");
	assert.ok(/invert\(1\)/.test(imgFilter), "pictures keep their colours: " + imgFilter);
	await page.reload();
	await page.waitFor("window.documentsApp && documentsApp.hasNode()");
	assert.ok(await page.eval("documentsApp.$.viewer.hasClass('night') && documentsApp.hasClass('night')"));
});

t.test("document info: title, author, type and size", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.odt");
	await page.tap(".viewer-toolbar .viewer-tool:nth-child(7)");
	await page.waitFor("documentsApp.$.viewer.$.infoDialog.isOpen");
	var info = await page.eval("documentsApp.$.viewer.$.infoText.hasNode().textContent");
	assert.ok(/Fixture Text/.test(info) && /Documents tests/.test(info) && /OpenDocument text/.test(info) && /KB/.test(info), info);
	await key("Escape");
	assert.ok(!(await page.eval("documentsApp.$.viewer.$.infoDialog.isOpen")), "back closes the dialog first");
});

// --- Spreadsheets ---------------------------------------------------------------------

t.test("sheet: cells, number alignment, frozen header row, sheet tabs", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.xlsx");
	assert.ok(/Pen/.test(await text(".sheet-cells")));
	var right = await page.eval("Array.prototype.filter.call(document.querySelectorAll('.sheet-cells .c'), function(e){return e.textContent==='12.25';})[0].className");
	assert.ok(/\bn\b/.test(right), right);
	assert.strictEqual(await page.eval("document.querySelector('.sheet-colinner .fz').textContent"), "Item", "row 1 stays on top");
	assert.strictEqual(await page.eval("Array.prototype.filter.call(document.querySelectorAll('.sheet-cells .c'), function(e){return e.textContent==='Item';}).length"), 0);
	assert.strictEqual(await count(".sheet-tab"), 2);
	await page.tap(".sheet-tab:nth-child(2)");
	await page.waitFor("/Second sheet/.test(document.querySelector('.sheet-colinner').textContent + document.querySelector('.sheet-cells').textContent)");
	assert.ok(/findme in sheet two/.test(await text(".sheet-cells")));
});

t.test("sheet: tapping a cell shows its formula; merged cells span", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.ods");
	var merged = await page.eval("(function(){var e=document.querySelector('.sheet-cells .c.m');return e?{t:e.textContent,w:e.getBoundingClientRect().width}:null;})()");
	assert.strictEqual(merged.t, "Merged total label");
	assert.ok(merged.w > 150, "three columns wide: " + merged.w);
	await page.eval("documentsApp.$.viewer.$.sheet.$.grid.select(4, 3); 1");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.sheet.$.ref.content"), "D5");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.sheet.$.value.content"), "=SUM(D2:D4)");
});

t.test("sheet: 20000 rows, only the cells near the screen are built", async function() {
	await open(scenario());
	await openDoc("/media/internal/big.csv");
	await page.waitFor("!documentsApp.$.viewer.streaming && documentsApp.$.viewer.$.sheet.currentSheet().rows === 20001", 20000);
	var cells = await count(".sheet-cells .c");
	assert.ok(cells > 30 && cells < 1500, "built " + cells + " cells");
	await page.eval("documentsApp.$.viewer.$.sheet.$.grid.scrollToCell(20000, 0); 1");
	await page.waitFor("/name 20000/.test(document.querySelector('.sheet-cells').textContent)");
	assert.ok(/20001/.test(await text(".sheet-rowinner")));
	assert.ok((await count(".sheet-cells .c")) < 1500);
});

t.test("sheet: find goes to the cell", async function() {
	await open(scenario());
	await openDoc("/media/internal/big.csv");
	await page.waitFor("!documentsApp.$.viewer.streaming", 20000);
	await page.eval("documentsApp.$.viewer.openFind(); 1");
	await page.waitFor("document.activeElement === documentsApp.$.viewer.$.findInput.$.input.hasNode()");
	await page.type("name 15000");
	await page.waitFor("documentsApp.$.viewer.$.findCount.content === '1 of 1'", 5000);
	await page.waitFor("document.querySelector('.sheet-hit') && /name 15000/.test(document.querySelector('.sheet-cells').textContent)");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.sheet.$.ref.content"), "B15001");
});

t.test("sheet: app menu freezes and unfreezes the first row; zoom scales the grid", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.xlsx");
	await page.eval("documentsApp.$.appMenu.open(); 1");
	await page.waitFor("documentsApp.$.appMenu.isOpen");
	assert.ok(await page.eval("documentsApp.$.menuFreeze.showing"), "the freeze item is offered for a workbook");
	assert.strictEqual(await page.eval("documentsApp.$.menuFreeze.caption"), "Unfreeze First Row");
	await tapControl("documentsApp.$.menuFreeze");
	await page.waitFor("document.querySelectorAll('.sheet-colinner .fz').length === 0");
	assert.ok(/Item/.test(await text(".sheet-cells")), "row 1 is back among the cells");
	assert.strictEqual(await page.eval("JSON.parse(localStorage.getItem('documents.prefs')).freeze"), false);
	var w1 = await page.eval("document.querySelector('.sheet-space').getBoundingClientRect().width");
	await page.eval("documentsApp.$.viewer.zoomIn(); documentsApp.$.viewer.zoomIn(); 1");
	var w2 = await page.eval("document.querySelector('.sheet-space').getBoundingClientRect().width");
	assert.ok(w2 > w1 * 1.1, w1 + " -> " + w2);
	await page.eval("documentsApp.$.appMenu.close(); documentsApp.$.pane.selectViewByName('library'); 1");
	await page.waitFor("documentsApp.$.pane.getViewName() === 'library' && !documentsApp.$.viewer.doc");
	await page.eval("documentsApp.$.appMenu.open(); 1");
	assert.ok(!(await page.eval("documentsApp.$.menuFreeze.showing")), "not offered on the library");
	assert.ok(await page.eval("documentsApp.$.menuFind.disabled"));
});

t.test("sheet editing: not offered when the app's service is not installed", async function() {
	await open(scenario({service: false}));
	await openDoc("/media/internal/sample.xlsx");
	await page.waitFor("(function(){return __MOCK__.calls.some(function(c){return /service\\/info/.test(c.url);});})()");
	await browserLib.sleep(200);
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.sheet.$.editButton.showing"), false);
});

t.test("sheet editing: edit a cell, save a copy next to the original, never over it", async function() {
	await open(scenario({service: true}));
	await openDoc("/media/internal/sample.xlsx");
	await page.waitFor("documentsApp.$.viewer.$.sheet.$.editButton.showing");
	var c = await page.eval("(function(){var e=Array.prototype.filter.call(document.querySelectorAll('.sheet-cells .c'), function(x){return x.textContent==='Pen';})[0];var r=e.getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2};})()");
	await page.send("Input.dispatchMouseEvent", {type: "mousePressed", x: c.x, y: c.y, button: "left", clickCount: 1});
	await page.send("Input.dispatchMouseEvent", {type: "mouseReleased", x: c.x, y: c.y, button: "left", clickCount: 1});
	await page.waitFor("documentsApp.$.viewer.$.sheet.$.ref.content === 'A2'");
	await page.tap(".sheet-edit-button");
	await page.waitFor("documentsApp.$.viewer.$.sheet.$.editor.showing");
	await page.eval("documentsApp.$.viewer.$.sheet.$.editor.setValue('Fountain pen'); 1");
	await page.key("Enter");
	await page.waitFor("/Fountain pen/.test(document.querySelector('.sheet-cells').textContent)");
	await page.waitFor("documentsApp.$.viewer.$.saveButton.showing");
	await page.tap(".viewer-save");
	await page.waitFor("documentsApp.$.saveDialog.isOpen");
	assert.strictEqual(await page.eval("documentsApp.$.saveName.getValue()"), "sample (edited).xlsx");
	await page.eval("documentsApp.doSave(); 1");
	await page.waitFor("!documentsApp.$.saveDialog.isOpen", 10000);
	var w = await calls("luna", "service/writeNew");
	assert.strictEqual(w.length, 1);
	assert.strictEqual(w[0].params.path, "/media/internal/sample (edited).xlsx");
	assert.strictEqual(w[0].acg, "granted");
	var wb = XLSX.read(Buffer.from(w[0].params.data, "base64"), {type: "buffer"});
	assert.strictEqual(wb.Sheets.Prices.A2.v, "Fountain pen");
	assert.strictEqual(wb.Sheets.Prices.B3.v, 12.25, "the rest of the workbook is in the copy");
	assert.ok(/Saved as sample \(edited\)\.xlsx/.test(await page.eval("documentsApp.$.toaster.content")));
});

t.test("sheet editing: an existing name is reported, nothing is overwritten", async function() {
	await open(scenario({service: true, luna: {"luna://org.webosports.app.documents.service/writeNew":
		{error: {errorCode: "EEXIST", errorText: "EEXIST: file already exists"}}}}));
	await openDoc("/media/internal/semicolon.csv");
	await page.waitFor("documentsApp.$.viewer.$.sheet.$.editButton.showing");
	await page.eval("var s=documentsApp.$.viewer.$.sheet; s.$.grid.select(2,2); s.startEdit(); s.$.editor.setValue('x'); s.commitEdit(); 1");
	await page.eval("documentsApp.$.viewer.saveCopy(); 1");
	await page.waitFor("documentsApp.$.saveDialog.isOpen");
	assert.strictEqual(await page.eval("documentsApp.$.saveName.getValue()"), "semicolon (edited).csv");
	await page.eval("documentsApp.doSave(); 1");
	await page.waitFor("documentsApp.$.saveError.showing");
	assert.ok(/already exists/.test(await page.eval("documentsApp.$.saveError.content")));
	assert.ok(await page.eval("documentsApp.$.saveDialog.isOpen"), "the dialog stays for another name");
});

t.test("sheet editing: a formula's result and the results that use the edited cell are shown at once", async function() {
	await open(scenario({service: true}));
	await openDoc("/media/internal/sample.xlsx");
	await page.waitFor("documentsApp.$.viewer.$.sheet.$.editButton.showing");
	// Prices: B2 1.5, C2 4, D2 =B2*C2, D5 =SUM(D2:D4).
	await page.eval("var s=documentsApp.$.viewer.$.sheet; s.$.grid.select(1,1); s.startEdit(); s.$.editor.setValue('2'); s.commitEdit(); 1");
	await page.waitFor("documentsApp.$.viewer.$.sheet.currentSheet().text[1][3] === '8'", 15000);
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.sheet.currentSheet().text[4][3]"), "62.5");
	await page.eval("var s=documentsApp.$.viewer.$.sheet; s.$.grid.select(6,0); s.startEdit(); s.$.editor.setValue('=B2*10'); s.commitEdit(); 1");
	await page.waitFor("documentsApp.$.viewer.$.sheet.currentSheet().text[6][0] === '20'", 15000);
	if (SHOTS) {
		await page.screenshot(path.join(OUT, "shots", "edit-xlsx.png"));
	}
	await page.waitFor("/62\\.5/.test(document.querySelector('.sheet-cells').textContent)");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.sheet.cellSource(documentsApp.$.viewer.$.sheet.currentSheet(), 6, 0)"), "=B2*10",
		"the cell bar shows the formula");
	await page.eval("documentsApp.$.viewer.saveCopy(); 1");
	await page.waitFor("documentsApp.$.saveDialog.isOpen");
	assert.ok(!/not the formatting/.test(await page.eval("documentsApp.$.saveNote.content")), "an .xlsx copy keeps the formatting");
	await page.eval("documentsApp.doSave(); 1");
	await page.waitFor("!documentsApp.$.saveDialog.isOpen", 10000);
	var w = await calls("luna", "service/writeNew");
	var ws = XLSX.read(Buffer.from(w[0].params.data, "base64"), {type: "buffer", cellFormula: true}).Sheets.Prices;
	assert.strictEqual(ws.D2.v, 8);
	assert.strictEqual(ws.A7.f, "B2*10");
	assert.strictEqual(ws.A7.v, 20);
});

t.test("sheet editing: the 123 key makes the cell editor a field for digits (a keypad's), and it is remembered", async function() {
	await open(scenario({service: true}));
	await openDoc("/media/internal/sample.xlsx");
	await page.waitFor("documentsApp.$.viewer.$.sheet.$.editButton.showing");
	await page.eval("var s=documentsApp.$.viewer.$.sheet; s.$.grid.select(6,0); s.startEdit(); s.$.editor.setValue('12'); 1");
	await page.waitFor("documentsApp.$.viewer.$.sheet.$.digitsButton.showing");
	assert.strictEqual(await page.eval("document.querySelector('.sheet-editor input').type"), "text");
	await page.tap(".sheet-digits-button");
	assert.strictEqual(await page.eval("document.querySelector('.sheet-editor input').type"), "tel");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.sheet.$.editor.getValue()"), "12", "what was typed stays");
	assert.strictEqual(await page.eval("localStorage.getItem('documents.cellDigits')"), "1");
	await page.eval("documentsApp.$.viewer.$.sheet.commitEdit(); 1");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.sheet.$.digitsButton.showing"), false, "only while editing");
});

t.test("text editing: not offered without the app's service", async function() {
	await open(scenario());
	await openDoc("/media/internal/sample.odt");
	await page.eval("new Promise(function(r){setTimeout(r, 500);})");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.editButton.showing"), false);
});

t.test("text editing: a tapped paragraph is edited in place and saved as a copy of the .odt", async function() {
	await open(scenario({service: true}));
	await openDoc("/media/internal/sample.odt");
	await page.waitFor("documentsApp.$.viewer.$.editButton.showing");
	await page.eval("documentsApp.$.viewer.toggleEdit(); 1");
	await page.waitFor("document.querySelector('.doc-editing [data-pid]')");
	var pid = await page.eval("Array.prototype.filter.call(document.querySelectorAll('[data-pid]'), function(x){return /bold words/.test(x.textContent);})[0].getAttribute('data-pid')");
	await page.tap("[data-pid='" + pid + "']");
	await page.waitFor("document.querySelector('.doc-para-input')");
	assert.strictEqual(await page.eval("document.querySelector('.doc-para-input').value"), "Plain, bold words and red underlined text, E=mc2.");
	if (SHOTS) {
		await page.screenshot(path.join(OUT, "shots", "edit-odt.png"));
	}
	await page.eval("var a=document.querySelector('.doc-para-input'); a.value='Plain, very bold words and red underlined text, E=mc2!'; 1");
	await page.tap(".doc-para-done");
	await page.waitFor("!document.querySelector('.doc-para-input') && /very bold/.test(document.querySelector('.viewer-content').textContent)", 10000);
	// The bold run is still bold: the paragraph was built again from the edited document.
	assert.ok(await page.eval("Array.prototype.some.call(document.querySelectorAll('[data-pid] span'), function(s){return s.textContent==='bold words' && /bold|700/.test(s.style.fontWeight);})"));
	await page.waitFor("documentsApp.$.viewer.$.saveButton.showing");
	// Back leaves edit mode first.
	await page.eval("documentsApp.$.viewer.back(); 1");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.editMode"), false);
	await page.tap(".viewer-save");
	await page.waitFor("documentsApp.$.saveDialog.isOpen");
	assert.strictEqual(await page.eval("documentsApp.$.saveName.getValue()"), "sample (edited).odt");
	await page.eval("documentsApp.doSave(); 1");
	await page.waitFor("!documentsApp.$.saveDialog.isOpen", 10000);
	var w = await calls("luna", "service/writeNew");
	assert.strictEqual(w.length, 1);
	assert.strictEqual(w[0].params.path, "/media/internal/sample (edited).odt");
	var JSZip = require(path.join(ROOT, "lib/jszip/jszip.min.js"));
	var zip = await JSZip.loadAsync(Buffer.from(w[0].params.data, "base64"));
	var content = await zip.file("content.xml").async("string");
	assert.ok(/Plain, very <text:span text:style-name="T1">bold words<\/text:span>/.test(content), "the bold run is kept");
	assert.ok(zip.file("Pictures/") !== undefined || Object.keys(zip.files).some(function(n) {
		return /^Pictures\//.test(n);
	}), "the pictures are in the copy");
});

[["sample.pptx", "ppt/slides/slide1.xml"], ["sample.odp", "content.xml"]].forEach(function(inCase) {
	t.test("slide editing: a paragraph on a slide of " + inCase[0] + ", edited under the slide and saved as a copy", async function() {
		await open(scenario({service: true}));
		await openDoc("/media/internal/" + inCase[0]);
		await page.waitFor("documentsApp.$.viewer.$.editButton.showing");
		await page.eval("documentsApp.$.viewer.toggleEdit(); 1");
		var pid = await page.eval("Array.prototype.filter.call(document.querySelectorAll('.doc-slide [data-pid]'), function(x){return /Yellow box/.test(x.textContent);})[0].getAttribute('data-pid')");
		// On the second slide, below the screen: brought into view first (the scroller is Enyo's).
		await page.eval("documentsApp.$.viewer.scrollToNode(document.querySelector(\".doc-slide [data-pid='" + pid + "']\")); 1");
		await page.tap(".doc-slide [data-pid='" + pid + "']");
		await page.waitFor("document.querySelector('.doc-slide-editor .doc-para-input')");
		assert.strictEqual(await page.eval("document.querySelector('.doc-para-input').value"), "Yellow box");
		assert.ok(await page.eval("document.querySelector('.doc-slide-editor').parentNode.previousSibling.classList.contains('doc-slide-frame')"),
			"the box is under its slide, not in the scaled slide");
		assert.ok(await page.eval("document.querySelector('.doc-para-editing') !== null"), "the paragraph is marked");
		if (SHOTS) {
			await page.screenshot(path.join(OUT, "shots", "edit-slide-" + inCase[0].replace(".", "-") + ".png"));
		}
		await page.eval("document.querySelector('.doc-para-input').value = 'Yellow box, edited'; 1");
		await page.tap(".doc-para-done");
		await page.waitFor("!document.querySelector('.doc-para-input') && /Yellow box, edited/.test(document.querySelector('.viewer-content').textContent)", 10000);
		assert.strictEqual(await page.eval("document.querySelectorAll('.doc-slide-frame').length"), await page.eval("documentsApp.$.viewer.slideBoxes.length"));
		assert.strictEqual(await page.eval("document.querySelectorAll('.doc-slide-number')[1].textContent"), "2", "slides keep their numbers");
		await page.waitFor("documentsApp.$.viewer.$.saveButton.showing");
		await page.eval("documentsApp.$.viewer.back(); documentsApp.$.viewer.saveCopy(); 1");
		await page.waitFor("documentsApp.$.saveDialog.isOpen");
		assert.strictEqual(await page.eval("documentsApp.$.saveName.getValue()"), inCase[0].replace(".", " (edited)."));
		await page.eval("documentsApp.doSave(); 1");
		await page.waitFor("!documentsApp.$.saveDialog.isOpen", 10000);
		var w = await calls("luna", "service/writeNew");
		var JSZip = require(path.join(ROOT, "lib/jszip/jszip.min.js"));
		var zip = await JSZip.loadAsync(Buffer.from(w[0].params.data, "base64"));
		var all = "";
		var names = Object.keys(zip.files).filter(function(n) {
			return /slides\/slide\d+\.xml$|^content\.xml$/.test(n);
		});
		for (var i = 0; i < names.length; i++) {
			all += await zip.file(names[i]).async("string");
		}
		assert.ok(/Yellow box, edited/.test(all), "the copy has the new text");
	});
});

t.test("text editing: a plain text file in one box, saved as a copy", async function() {
	await open(scenario({service: true}));
	await openDoc("/media/internal/sample.txt");
	await page.waitFor("documentsApp.$.viewer.$.editButton.showing");
	await page.eval("documentsApp.$.viewer.toggleEdit(); 1");
	await page.waitFor("document.querySelector('.doc-text-input')");
	await page.eval("var a=document.querySelector('.doc-text-input'); a.value='Edited\\n' + a.value; a.dispatchEvent(new Event('input')); 1");
	await page.waitFor("documentsApp.$.viewer.$.saveButton.showing");
	await page.eval("documentsApp.$.viewer.toggleEdit(); 1");
	await page.waitFor("!document.querySelector('.doc-text-input') && /^Edited/.test(document.querySelector('.doc-plaintext').textContent)");
	await page.eval("documentsApp.$.viewer.saveCopy(); 1");
	await page.waitFor("documentsApp.$.saveDialog.isOpen");
	assert.strictEqual(await page.eval("documentsApp.$.saveName.getValue()"), "sample (edited).txt");
	await page.eval("documentsApp.doSave(); 1");
	await page.waitFor("!documentsApp.$.saveDialog.isOpen", 10000);
	var w = await calls("luna", "service/writeNew");
	var saved = Buffer.from(w[0].params.data, "base64").toString("utf8");
	assert.ok(/^Edited\n/.test(saved), saved.slice(0, 40));
	assert.ok(saved.indexOf(fs.readFileSync(path.join(FIX, "sample.txt"), "utf8").slice(0, 20)) > 0, "the rest of the text");
});

t.test("word editing: the .docx editor, its own toolbar keys, and a copy without the editor's stamps", async function() {
	await open(scenario({service: true}));
	await openDoc("/media/internal/sample.docx");
	await page.waitFor("documentsApp.$.viewer.$.editButton.showing");
	await page.tap(".viewer-toolbar .viewer-tool:last-child");
	await page.waitFor("documentsApp.$.viewer.$.docx.isOpen() && documentsApp.$.viewer.$.docx.showing && documentsApp.$.viewer.editMode", 60000);
	assert.ok(/Fixture heading/.test(await text(".docx-editor-host")));
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.findButton.disabled"), true, "find is the editor's while it is open");
	// Type at the start, then make a word bold with the bar's key.
	await page.eval("(function(){var e=documentsApp.$.viewer.$.docx.superdoc.activeEditor; e.commands.setTextSelection({from: 3, to: 3}); " +
		"e.commands.insertContent('Edited '); e.commands.setTextSelection({from: 3, to: 9}); return 1;})()");
	await page.tap(".docx-b");
	await page.waitFor("documentsApp.$.viewer.$.saveButton.showing");
	if (SHOTS) {
		await page.screenshot(path.join(OUT, "shots", "edit-docx.png"));
	}
	await page.tap(".viewer-save");
	await page.waitFor("documentsApp.$.saveDialog.isOpen");
	assert.strictEqual(await page.eval("documentsApp.$.saveName.getValue()"), "sample (edited).docx");
	await page.eval("documentsApp.doSave(); 1");
	await page.waitFor("!documentsApp.$.saveDialog.isOpen", 20000);
	var w = await calls("luna", "service/writeNew");
	assert.strictEqual(w[0].params.path, "/media/internal/sample (edited).docx");
	var JSZip = require(path.join(ROOT, "lib/jszip/jszip.min.js"));
	var zip = await JSZip.loadAsync(Buffer.from(w[0].params.data, "base64"));
	var doc = await zip.file("word/document.xml").async("string");
	assert.ok(/Edited/.test(doc), "the typed text is in the copy");
	assert.ok(/<w:b\/>|<w:b /.test(doc), "and the bold");
	assert.ok(/Plain, /.test(doc), "and the rest of the document");
	assert.strictEqual(zip.file("docProps/custom.xml"), null, "no editor stamps");
	// Back: read only first, the editor stays.
	await page.eval("documentsApp.$.viewer.back(); 1");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.editMode"), false);
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.docx.$.bar.showing"), false);
});

t.test("word editing on a phone: the page fits the screen, the zoom keys work the editor's zoom", async function() {
	await open(scenario({service: true}), {width: 480, height: 800});
	await openDoc("/media/internal/sample.docx");
	await page.waitFor("documentsApp.$.viewer.$.editButton.showing");
	await page.eval("documentsApp.$.viewer.toggleEdit(); 1");
	await page.waitFor("documentsApp.$.viewer.$.docx.isOpen() && documentsApp.$.viewer.editMode", 60000);
	await page.waitFor("documentsApp.$.viewer.$.docx.zoomPercent() < 100", 10000);
	var m = await page.eval("({doc: document.documentElement.scrollWidth, bar: document.querySelector('.docx-editor-bar').scrollWidth, " +
		"label: documentsApp.$.viewer.$.zoomLabel.caption, zoom: documentsApp.$.viewer.$.docx.zoomPercent()})");
	assert.ok(m.doc <= 480 && m.bar <= 480, JSON.stringify(m));
	assert.strictEqual(m.label, m.zoom + "%", "the label shows the editor's zoom");
	await page.eval("documentsApp.$.viewer.zoomIn(); 1");
	await page.waitFor("documentsApp.$.viewer.$.docx.zoomPercent() > " + m.zoom);
	if (SHOTS) {
		await page.screenshot(path.join(OUT, "shots", "edit-docx-phone.png"));
	}
});

t.test("conversion: with LibreOffice's service a .doc is listed, converted and shown under its own name", async function() {
	await open(scenario({service: true, office: true}));
	await page.waitFor("FileTypes.converter === true", 10000);
	var names = await libraryNames();
	assert.ok(names.indexOf("legacy.doc") >= 0, "a .doc is listed: " + names.join(", "));
	assert.strictEqual(await openDoc("/media/internal/legacy.doc"), "odt");
	assert.ok(/Fixture heading/.test(await text(".viewer-content")));
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.title.content"), "legacy.doc");
	assert.ok(/Word 97-2003 document, converted by LibreOffice/.test(await page.eval("documentsApp.$.viewer.$.subtitle.content")));
	var c = await calls("luna", "service.office/convert");
	assert.ok(c.some(function(x) {
		return x.params.source === "/media/internal/legacy.doc" && x.params.format === "odt" && x.acg === "granted";
	}), JSON.stringify(c));
	// A copy goes next to the .doc, as OpenDocument.
	await page.eval("documentsApp.$.viewer.saveCopy(); 1");
	await page.waitFor("documentsApp.$.saveDialog.isOpen");
	assert.strictEqual(await page.eval("documentsApp.$.saveName.getValue()"), "legacy (edited).odt");
});

t.test("conversion: without the service a .doc is not listed, and opened it says why", async function() {
	await open(scenario());
	await libraryNames();
	assert.strictEqual(await page.eval("FileTypes.converter"), false);
	assert.strictEqual(await openDoc("/media/internal/legacy.doc"), "error");
	assert.ok(/binary format from before 2007/.test(await page.eval("documentsApp.$.viewer.$.statusText.content")));
});

// --- Layouts ------------------------------------------------------------------------

[["tablet", 1024, 768], ["phone", 480, 800], ["square", 450, 380]].forEach(function(inSize) {
	t.test("layout " + inSize[0] + " " + inSize[1] + "x" + inSize[2] + ": nothing wider than the screen, tools fit", async function() {
		// With the service, so the edit tool is in the toolbar too.
		await open(scenario({service: true}), {width: inSize[1], height: inSize[2]});
		await libraryNames();
		var lib = await page.eval("({doc: document.documentElement.scrollWidth, toolbar: document.querySelector('.library-toolbar').scrollWidth})");
		assert.ok(lib.doc <= inSize[1] && lib.toolbar <= inSize[1], JSON.stringify(lib));
		// Each filter's caption fits inside its button (the theme's 16px ends take room from the width).
		var clipped = await page.eval("Array.prototype.filter.call(document.querySelectorAll('.library-filters .enyo-grouped-toolbutton'), " +
			"function(b){var r=document.createRange();r.selectNodeContents(b);var t=r.getBoundingClientRect().width;" +
			"var cs=getComputedStyle(b);var inner=b.getBoundingClientRect().width-parseFloat(cs.borderLeftWidth)-parseFloat(cs.borderRightWidth);" +
			"return t>inner+0.5;}).map(function(b){return b.textContent;})");
		assert.deepStrictEqual(clipped, [], "filter captions fit");
		await openDoc("/media/internal/sample.odt");
		await page.waitFor("documentsApp.$.viewer.$.editButton.showing");
		var reflow = await page.eval("documentsApp.$.viewer.reflow");
		assert.strictEqual(reflow, inSize[1] < 700, "phones read reflowed text");
		var m = await page.eval("({page: document.querySelector('.doc-page').getBoundingClientRect().right, tools: document.querySelector('.viewer-toolbar').scrollWidth, " +
			"lastTool: document.querySelector('.viewer-toolbar .viewer-tool:last-child').getBoundingClientRect().right})");
		assert.ok(m.page <= inSize[1] + 1, "page fits: " + JSON.stringify(m));
		assert.ok(m.lastTool <= inSize[1], "toolbar fits: " + JSON.stringify(m));
		await openDoc("/media/internal/sample.xlsx");
		var g = await page.eval("({grid: document.querySelector('.sheet-grid').getBoundingClientRect().right, h: document.querySelector('.sheet-grid').getBoundingClientRect().height})");
		assert.ok(g.grid <= inSize[1] && g.h > 100, JSON.stringify(g));
		if (SHOTS) {
			await page.screenshot(path.join(OUT, "shots", "layout-" + inSize[0] + ".png"));
		}
	});
});

// --- Run --------------------------------------------------------------------------------

(async function() {
	makeE2eRoot();
	server = await serverLib.start(ROOT);
	if (!fs.existsSync(path.join(ROOT, "enyo-framework"))) {
		var enyo = process.env.ENYO || path.join(ROOT, "..", "enyo-1.0-ports", "framework");
		fs.symlinkSync(enyo, path.join(ROOT, "enyo-framework"));
	}
	browser = await browserLib.launch();
	await t.run({
		afterEach: async function(inTest) {
			var errors = page.errors.filter(function(e) {
				return !/Tellurium/.test(e);
			});
			assert.deepStrictEqual(errors, [], "page errors");
			var bad = await page.eval("__MOCK__.calls.filter(function(c){return c.unmocked || (c.acg && /^denied|unknown/.test(c.acg));})" +
				".map(function(c){return c.kind+' '+c.url+' '+(c.acg||'');})");
			assert.deepStrictEqual(bad, [], "unmocked or denied calls");
			if (SHOTS) {
				await page.screenshot(path.join(OUT, "shots", inTest.name.replace(/[^a-z0-9]+/gi, "-") + ".png"));
			}
		},
		onFailure: async function(inTest) {
			if (page) {
				await page.screenshot(path.join(OUT, "shots", "FAIL-" + inTest.name.replace(/[^a-z0-9]+/gi, "-") + ".png"));
			}
		},
		after: async function() {
			await browser.close();
			await server.stop();
		}
	});
})();
