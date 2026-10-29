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
var Cloud = require(path.join(ROOT, "source/services/CloudService.js"));

var server, browser, page;

// --- Documents made for the tests --------------------------------------------

function makeE2eRoot() {
	fs.rmSync(E2E_ROOT, {recursive: true, force: true});
	fs.mkdirSync(path.join(E2E_ROOT, "Work"), {recursive: true});
	// What the office service's editing session hands out: a 256 px tile, all one colour, and a saved copy
	fs.mkdirSync(path.join(E2E_ROOT, ".cache"), {recursive: true});
	var tile = Buffer.alloc(256 * 256 * 4);
	for (var px = 0; px < tile.length; px += 4) {
		tile[px] = 10;
		tile[px + 1] = 200;
		tile[px + 2] = 30;
		tile[px + 3] = 255;
	}
	fs.writeFileSync(path.join(E2E_ROOT, ".cache", "tile.raw"), tile);
	fs.copyFileSync(path.join(FIX, "sample.pptx"), path.join(E2E_ROOT, ".cache", "saved.pptx"));
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
	// What the connector's downloadFile leaves in the cloud cache, where the app asks for it
	[[CLOUD.plan, "sample.odt"], [CLOUD.budget, "sample.xlsx"], [CLOUD.notes, "sample.txt"]].forEach(function(inPair) {
		var local = path.join(E2E_ROOT, Cloud.cachePath(CLOUD.dropbox, inPair[0]).replace(/^\/media\/internal\//, ""));
		fs.mkdirSync(path.dirname(local), {recursive: true});
		fs.copyFileSync(path.join(FIX, inPair[1]), local);
	});
}

// --- The cloud: a Dropbox account and a Box one, through their connectors -------------

var CLOUD = {
	dropbox: {id: "acc1", provider: "Dropbox", name: "Herrie", service: "com.palm.service.dropbox"},
	box: {id: "acc2", provider: "Box", name: "work@example.com", service: "com.palm.service.boxnet"},
	plan: {id: "id:plan", type: "file", name: "Plan.odt", size: 9000, modified: "2026-05-01T10:00:00Z", path: "/Work/Plan.odt"},
	budget: {id: "id:budget", type: "file", name: "Budget.xlsx", size: 7000, modified: "2026-05-02T10:00:00Z", path: "/Work/Budget.xlsx"},
	notes: {id: "id:notes", type: "file", name: "notes.txt", size: 12, path: "/Work/notes.txt"}
};

function cloudLuna() {
	var account = function(inAccount, inPhotos) {
		var providers = [{capability: "DOCUMENTS", implementation: "palm://" + inAccount.service + "/"}];
		if (inPhotos) {
			providers.unshift({capability: "PHOTO.UPLOAD", implementation: "palm://com.palm.service.photos"});
		}
		// Dropbox's template has its picture (with a quote in its path, which must not end the
		// row's style attribute); Box's has none and keeps the cloud.
		var icon = inPhotos ? {loc_32x32: "/usr/palm/public/accounts/com.palm.dropbox/images/d'b (32).png",
			loc_48x48: "/usr/palm/public/accounts/com.palm.dropbox/images/d'b (48).png"} : undefined;
		return {_id: inAccount.id, alias: inAccount.name === "Herrie" ? "Herrie" : "", username: inAccount.name,
			loc_name: inAccount.provider, icon: icon, capabilityProviders: providers};
	};
	return {
		"luna://com.palm.service.accounts/listAccounts": {cases: [{when: {capability: "DOCUMENTS"}, json: {results: [
			account(CLOUD.dropbox, true), account(CLOUD.box),
			// Photos only: not listed
			{_id: "acc3", username: "pics", loc_name: "Flickr", capabilityProviders: [
				{capability: "PHOTO.UPLOAD", implementation: "palm://com.palm.service.flickr/"}]}
		]}}], error: {errorCode: -1, errorText: "capability needed"}},
		"luna://com.palm.service.dropbox/listFolder": {cases: [
			{when: {accountId: "acc1", path: "/Work"}, json: {entries: [
				CLOUD.plan, CLOUD.budget,
				{id: "id:old", type: "file", name: "Budget (old).xlsx", size: 7000, path: "/Work/Budget (old).xlsx"},
				CLOUD.notes,
				{id: "id:photo", type: "file", name: "photo.jpg", size: 30000, path: "/Work/photo.jpg"},
				{id: "id:reports", type: "folder", name: "Reports", path: "/Work/Reports"}
			]}},
			{when: {accountId: "acc1", path: "/Work/Reports"}, json: {entries: []}},
			{when: {accountId: "acc1"}, json: {entries: [
				{id: "id:work", type: "folder", name: "Work", path: "/Work"},
				{id: "id:song", type: "file", name: "song.mp3", size: 4000000, path: "/song.mp3"},
				{id: "id:welcome", type: "file", name: "Welcome.pdf", size: 120000, path: "/Welcome.pdf"}
			]}}
		], error: {errorCode: "MISSING_ARGS", errorText: "unexpected listFolder"}},
		// Slow enough to see it busy
		"luna://com.palm.service.dropbox/downloadFile": {delay: 300, json: {path: "ignored"}},
		"luna://com.palm.service.dropbox/uploadFile": {delay: 100, json: {id: "id:new", name: "x", size: 1}},
		"luna://com.palm.service.boxnet/listFolder": {error: {errorCode: "NO_STORED_CREDENTIALS"}}
	};
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
	// The cloud connectors' perm.json files (luneos-synergy-connectors) give this app their groups.
	var connectors = ["synergy-dropbox.operation", "synergy-box.operation"];
	return {groups: groups, granted: appinfo.requiredPermissions.concat(["org.webosports.app.documents.service.group"], connectors),
		// Enyo 1's test hook, which it asks for at start on every device (and is told is not there).
		exempt: ["^luna://com\\.palm\\.telluriumservice/"]};
}

/**
	A LuneOS device with the fixtures on its USB storage. inOptions:
	launchParams, service (true: the app's service is installed), cloud
	(true: cloud accounts, see cloudLuna), luna (more rules), storage,
	palmSystem.
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
		// The app's own service lists with whole names; the File Manager's (above) joins them with commas.
		"luna://org.webosports.app.documents.service/list": o.service ? {cases: [
			{when: {dir: "/media/internal"}, json: {dirs: ["Work", ".thumbnails", "src"], files: files.split(",")}},
			{when: {dir: "/media/internal/Work"}, json: {dirs: [], files: ["Report, final.docx", "todo.txt.bak"]}},
			{when: {dir: "/media/internal/src"}, json: {dirs: [], files: ["sample.fodt", "sample.fods", "sample.fodp"]}}
		], error: {errorCode: "ENOENT", errorText: "ENOENT"}} : {error: {errorCode: -1, errorText: "Service does not exist"}},
		"luna://org.webosports.app.documents.service/size": o.service ? {cases: [
			{when: {path: "/media/internal/Work/Report, final.docx"}, json: {size: 8404}}
		], error: {errorCode: "ENOENT", errorText: "ENOENT"}} : {error: {errorCode: -1, errorText: "Service does not exist"}},
		// LibreOffice's service: asked without a document it says badRequest; legacy.doc becomes sample.odt.
		"luna://org.webosports.service.office/convert": o.office ? {cases: [
			{when: {source: "/media/internal/legacy.doc", format: "odt"}, json: {target: "/media/internal/sample.odt", format: "odt", cached: false}}
		], error: {errorCode: "badRequest", errorText: "source and format are required"}} :
			{error: {errorCode: -1, errorText: "Service does not exist: org.webosports.service.office."}},
		// Its editing sessions (o.lok): a service with them refuses an open without a document
		"luna://org.webosports.service.office/open": o.lok ? {cases: [
			{when: {source: "/media/internal/sample.pptx"}, json: {session: "s1", kind: "presentation", parts: 2, part: 0,
				width: 28000, height: 15750, tileMode: "rgba"}}
		], error: {errorCode: "badRequest", errorText: "need \"source\" (a path)"}} :
			o.office ? {error: {errorCode: -1, errorText: "Unknown method \"open\" for category \"/\""}} :
			{error: {errorCode: -1, errorText: "Service does not exist: org.webosports.service.office."}}
	};
	if (o.lok) {
		var ok = {json: {}};
		Object.assign(luna, {
			// The subscription's answer, then LibreOffice's callbacks
			"luna://org.webosports.service.office/events": {responses: [
				{json: {subscribed: true}},
				{delay: 400, json: {session: "s1", event: "cursor", type: 1, payload: "1440, 1440, 0, 400"}},
				{delay: 100, json: {session: "s1", event: "state", type: 8, payload: ".uno:Bold=true"}},
				{delay: 100, json: {session: "s1", event: "invalidate", type: 0, payload: "0, 0, 5000, 5000, 0"}}
			]},
			"luna://org.webosports.service.office/paintTile": {json: {file: "/media/internal/.cache/tile.raw", tileMode: "rgba"}},
			"luna://org.webosports.service.office/postMouse": ok,
			"luna://org.webosports.service.office/postKey": ok,
			"luna://org.webosports.service.office/extTextInput": ok,
			"luna://org.webosports.service.office/uno": ok,
			"luna://org.webosports.service.office/setPart": {json: {parts: 2, part: 1, width: 28000, height: 15750}},
			"luna://org.webosports.service.office/save": {json: {target: "/media/internal/.cache/saved.pptx"}},
			"luna://org.webosports.service.office/close": ok
		});
	}
	if (o.cloud) {
		Object.assign(luna, cloudLuna());
	}
	Object.assign(luna, o.luna || {});
	return {
		webos: true,
		palmSystem: Object.assign({identifier: "org.webosports.app.documents 1000", launchParams: JSON.stringify(o.launchParams || {})},
			o.palmSystem || {}),
		luna: luna,
		acg: acg(),
		rewrite: [
			{from: "^file:///media/internal/(big\\.csv|legacy\\.doc|Manual(%20| )1\\.pdf|scan\\.docx|Work/|\\.cache/|\\.documents/)", to: server.url + "/test-output/e2e-root/$1"},
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

t.test("library: with the app's service, listed through it, whole names and no File Manager", async function() {
	await open(scenario({service: true}));
	var names = await libraryNames();
	["sample.docx", "sample.xlsx", "sample.fodt", "Report, final.docx"].forEach(function(n) {
		assert.ok(names.indexOf(n) >= 0, n + " listed: " + names.join(", "));
	});
	assert.ok((await calls("luna", "documents.service/list")).length >= 3);
	assert.strictEqual((await calls("luna", "GetDirectories")).length, 0);
	assert.strictEqual((await calls("luna", "getFileSize")).length, 0, "a whole name needs no check");
	assert.ok(!(await calls("luna", "documents.service/list")).some(function(c) {
		return /thumbnails/.test(c.params.dir);
	}), "hidden folders are not walked");
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
	var info = await page.eval("documentsApp.$.viewer.$.infoDialog.hasNode().textContent");
	assert.ok(/sample\.odt/.test(info) && /Fixture Text/.test(info) && /Documents tests/.test(info) &&
		/OpenDocument text/.test(info) && /KB/.test(info), info);
	await key("Escape");
	assert.ok(!(await page.eval("documentsApp.$.viewer.$.infoDialog.isOpen")), "back closes the dialog first");
});

t.test("document info: opened first on a small screen, it fits and is centred, all of it reachable", async function() {
	await open(scenario(), {width: 450, height: 380});
	await openDoc("/media/internal/sample.odt");
	await page.eval("documentsApp.$.viewer.showInfo(); 1");
	await page.waitFor("documentsApp.$.viewer.$.infoDialog.isOpen");
	var m = JSON.parse(await page.eval("(function(){var d=documentsApp.$.viewer.$.infoDialog.hasNode().getBoundingClientRect(); " +
		"var b=documentsApp.$.viewer.$.infoDialog.hasNode().querySelector('.enyo-button').getBoundingClientRect(); " +
		"var sc=documentsApp.$.viewer.$.infoScroller; var g=documentsApp.$.viewer.$.infoGroups.hasNode(); " +
		"return JSON.stringify({top:d.top, bottom:d.bottom, left:d.left, right:d.right, h:innerHeight, w:innerWidth, " +
		"done:b.bottom, rows:g.querySelectorAll('.viewer-info-row').length, groups:g.offsetHeight, room:sc.hasNode().offsetHeight});})()"));
	if (SHOTS) {
		await page.eval("new Promise(function(r){setTimeout(r, 600);})");
		await page.screenshot(path.join(OUT, "shots", "info-small-open.png"));
	}
	assert.ok(m.top >= 0 && m.bottom <= m.h && m.done <= m.h, "inside the window, Done too: " + JSON.stringify(m));
	assert.ok(Math.abs(m.top - (m.h - m.bottom)) <= 2, "centred: " + JSON.stringify(m));
	assert.ok(m.rows >= 6, "the rows: " + JSON.stringify(m));
	var info = await page.eval("documentsApp.$.viewer.$.infoGroups.hasNode().textContent");
	assert.ok(/LibreOffice 26\.2\.6\.3/.test(info) && !/\$/.test(info), "the generator, readable: " + info);
	// More than fits: the groups scroll inside the dialog, and the last row can be reached
	assert.ok(m.groups > m.room, "taller than the room: " + JSON.stringify(m));
	await page.eval("documentsApp.$.viewer.$.infoScroller.scrollToBottom(); 1");
	await page.eval("new Promise(function(r){setTimeout(r, 800);})");
	var last = JSON.parse(await page.eval("(function(){var r=documentsApp.$.viewer.$.infoGroups.hasNode().querySelectorAll('.viewer-info-row'); " +
		"var l=r[r.length-1].getBoundingClientRect(); var s=documentsApp.$.viewer.$.infoScroller.hasNode().getBoundingClientRect(); " +
		"return JSON.stringify({top:l.top, bottom:l.bottom, st:s.top, sb:s.bottom});})()"));
	assert.ok(last.top >= last.st - 1 && last.bottom <= last.sb + 1, "the last row scrolled into view: " + JSON.stringify(last));
	if (SHOTS) {
		await page.screenshot(path.join(OUT, "shots", "info-small.png"));
	}
	await key("Escape");
});

t.test("about: on a small screen it fits, is centred and shows its button", async function() {
	await open(scenario(), {width: 450, height: 380});
	await page.eval("documentsApp.showAbout(); 1");
	await page.waitFor("documentsApp.$.about.isOpen");
	var m = JSON.parse(await page.eval("(function(){var d=documentsApp.$.about.hasNode().getBoundingClientRect(); " +
		"var b=documentsApp.$.about.hasNode().querySelector('.enyo-button').getBoundingClientRect(); " +
		"return JSON.stringify({top:d.top, bottom:d.bottom, h:innerHeight, ok:b.bottom, text:documentsApp.$.about.hasNode().textContent});})()"));
	assert.ok(m.top >= 0 && m.bottom <= m.h && m.ok <= m.h, "inside the window, OK too: " + JSON.stringify(m));
	assert.ok(Math.abs(m.top - (m.h - m.bottom)) <= 2, "centred: " + JSON.stringify(m));
	assert.ok(/Version 0\.1\.0/.test(m.text) && /GNU General Public License/.test(m.text), m.text);
	if (SHOTS) {
		await page.eval("new Promise(function(r){setTimeout(r, 600);})");
		await page.screenshot(path.join(OUT, "shots", "about-small.png"));
	}
	await page.eval("documentsApp.closeAbout(); 1");
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

t.test("LibreOffice editing: a session's tiles, taps, typing, Bold, the next slide and a copy", async function() {
	await open(scenario({service: true, office: true, lok: true}));
	await page.waitFor("FileTypes.lokSessions === true", 10000);
	await openDoc("/media/internal/sample.pptx");
	await page.eval("documentsApp.$.appMenu.open(); 1");
	await page.waitFor("documentsApp.$.appMenu.isOpen");
	assert.strictEqual(await page.eval("documentsApp.$.menuLok.showing"), true, "offered in the app menu");
	await tapControl("documentsApp.$.menuLok");
	await page.waitFor("documentsApp.$.viewer.$.lok.isOpen() && documentsApp.$.viewer.$.lok.showing", 10000);
	// A tile is asked for and drawn
	await page.waitFor("(function(){var c=document.querySelector('.lok-tile'); if(!c) return false; " +
		"var d=c.getContext('2d').getImageData(10,10,1,1).data; return d[0]===10 && d[1]===200 && d[2]===30;})()", 10000);
	var paints = await calls("luna", "service.office/paintTile");
	assert.ok(paints.length > 0 && paints[0].params.session === "s1" && paints[0].params.part === 0 &&
		paints[0].params.pixelWidth === 256, JSON.stringify(paints[0] && paints[0].params));
	// The events: the cursor drawn, Bold shown as on
	await page.waitFor("document.querySelector('.lok-cursor').style.display !== 'none'", 5000);
	await page.waitFor("documentsApp.$.viewer.$.lok.$.bold.hasClass('viewer-tool-on')", 5000);
	// A tap on the page: a click in LibreOffice's twips
	await page.tap(".lok-page");
	await page.waitFor("__MOCK__.calls.filter(function(c){return /postMouse/.test(c.url)}).length >= 2", 5000);
	var mouse = await calls("luna", "service.office/postMouse");
	assert.strictEqual(mouse[0].params.type, "down");
	assert.strictEqual(mouse[1].params.type, "up");
	assert.ok(mouse[0].params.x > 0 && mouse[0].params.y > 0 && mouse[0].params.x < 28000, JSON.stringify(mouse[0].params));
	// Repaints and taps are not edits
	await page.waitFor("__MOCK__.calls.filter(function(c){return /paintTile/.test(c.url)}).length > " + paints.length, 5000);
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.saveButton.showing"), false, "not edited yet");
	// Typing: the text box takes it and sends it as committed text
	await page.eval("document.querySelector('.lok-tile').__kept = true; document.querySelector('.lok-input').focus(); 1");
	await page.send("Input.insertText", {text: "Hi"});
	await page.waitFor("__MOCK__.calls.some(function(c){return /extTextInput/.test(c.url) && c.params.text === 'Hi' && c.params.type === 'commit'})", 5000);
	// The Save button shows with the first edit; the page keeps its tiles (no new layout for the same width)
	await page.waitFor("documentsApp.$.viewer.$.saveButton.showing", 5000);
	await page.eval("new Promise(function(r){setTimeout(r, 300);})");
	assert.strictEqual(await page.eval("document.querySelector('.lok-tile').__kept === true"), true, "the tiles kept");
	// Bold, through the key bar
	await page.eval("documentsApp.$.viewer.$.lok.uno(documentsApp.$.viewer.$.lok.$.bold); 1");
	var unos = await calls("luna", "service.office/uno");
	assert.ok(unos.some(function(u) {
		return u.params.command === ".uno:Bold";
	}), JSON.stringify(unos));
	// Night reading darkens LibreOffice's page too
	assert.strictEqual(await page.eval("getComputedStyle(document.querySelector('.lok-page')).filter"), "none");
	await page.eval("documentsApp.$.viewer.toggleNight(); 1");
	var lokFilter = await page.eval("getComputedStyle(document.querySelector('.lok-page')).filter");
	await page.eval("documentsApp.$.viewer.toggleNight(); 1");
	assert.ok(/invert/.test(lokFilter), "night: " + lokFilter);
	// The next slide
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.lok.$.partLabel.content"), "1 of 2");
	await page.eval("documentsApp.$.viewer.$.lok.nextPart(); 1");
	await page.waitFor("documentsApp.$.viewer.$.lok.$.partLabel.content === '2 of 2'", 5000);
	// Save a Copy: the session's saved file, next to the original
	await page.waitFor("documentsApp.$.viewer.$.saveButton.showing", 5000);
	await page.eval("documentsApp.$.viewer.saveCopy(); 1");
	await page.waitFor("documentsApp.$.saveDialog.isOpen");
	assert.strictEqual(await page.eval("documentsApp.$.saveName.getValue()"), "sample (edited).pptx");
	await page.eval("documentsApp.doSave(); 1");
	await page.waitFor("!documentsApp.$.saveDialog.isOpen", 10000);
	var w = await calls("luna", "service/writeNew");
	assert.strictEqual(w[0].params.path, "/media/internal/sample (edited).pptx");
	assert.ok(Buffer.from(w[0].params.data, "base64").equals(fs.readFileSync(path.join(FIX, "sample.pptx"))), "the saved bytes");
	// Back: the session is closed, the slides shown again
	await page.eval("documentsApp.$.viewer.back(); 1");
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.lok.isOpen()"), false);
	assert.ok((await calls("luna", "service.office/close")).length === 1);
	assert.strictEqual(await page.eval("documentsApp.$.viewer.$.scroller.showing"), true);
	if (SHOTS) {
		await page.screenshot(path.join(OUT, "shots", "lok-closed.png"));
	}
});

t.test("LibreOffice editing: a long press selects the word and offers cut, copy and paste", async function() {
	await open(scenario({service: true, office: true, lok: true}));
	await page.waitFor("FileTypes.lokSessions === true", 10000);
	await openDoc("/media/internal/sample.pptx");
	await page.eval("documentsApp.$.viewer.editWithLibreOffice(); 1");
	await page.waitFor("documentsApp.$.viewer.$.lok.isOpen() && !!document.querySelector('.lok-tile')", 10000);
	var p = JSON.parse(await page.eval("(function(){var r=document.querySelector('.lok-page').getBoundingClientRect(); " +
		"return JSON.stringify([Math.round(r.left + 60), Math.round(r.top + 40)]);})()"));
	var mice = (await calls("luna", "service.office/postMouse")).length;
	// Held, as a finger does: no tap, but a double click (the word), and the menu
	await page.send("Input.dispatchTouchEvent", {type: "touchStart", touchPoints: [{x: p[0], y: p[1]}]});
	await page.eval("new Promise(function(r){setTimeout(r, 700);})");
	await page.send("Input.dispatchTouchEvent", {type: "touchEnd", touchPoints: []});
	await page.waitFor("documentsApp.$.viewer.$.lok.$.editMenu.isOpen", 5000);
	var mouse = (await calls("luna", "service.office/postMouse")).slice(mice);
	assert.deepStrictEqual(mouse.map(function(c) {
		return c.params.type + c.params.count;
	}), ["down1", "up1", "down2", "up2"], "a double click (a click, then one counted 2), and no tap besides");
	var items = await page.eval("documentsApp.$.viewer.$.lok.$.editMenu.getControls().map(function(c){return c.caption;}).join(',')");
	assert.ok(/Cut/.test(items) && /Copy/.test(items) && /Paste/.test(items), items);
	if (SHOTS) {
		await page.screenshot(path.join(OUT, "shots", "lok-edit-menu.png"));
	}
	// Copy: LibreOffice's own
	var copy = await page.eval("(function(){var m=documentsApp.$.viewer.$.lok.$.editMenu; var c=m.getControls().filter(function(x){return x.caption==='Copy';})[0]; " +
		"var r=c.hasNode().getBoundingClientRect(); return {x:r.left+r.width/2, y:r.top+r.height/2};})()");
	await page.send("Input.dispatchMouseEvent", {type: "mousePressed", x: copy.x, y: copy.y, button: "left", clickCount: 1});
	await page.send("Input.dispatchMouseEvent", {type: "mouseReleased", x: copy.x, y: copy.y, button: "left", clickCount: 1});
	await page.waitFor("__MOCK__.calls.some(function(c){return /service.office\\/uno/.test(c.url) && c.params.command === '.uno:Copy'})", 5000);
	// On the selection, a long press keeps it: no click
	await page.eval("documentsApp.$.viewer.$.lok.event({event: 'selection', payload: '0, 0, 28000, 15750'}); 1");
	mice = (await calls("luna", "service.office/postMouse")).length;
	await page.send("Input.dispatchMouseEvent", {type: "mousePressed", x: p[0], y: p[1], button: "right", clickCount: 1});
	await page.send("Input.dispatchMouseEvent", {type: "mouseReleased", x: p[0], y: p[1], button: "right", clickCount: 1});
	await page.waitFor("documentsApp.$.viewer.$.lok.$.editMenu.isOpen", 5000);
	assert.strictEqual((await calls("luna", "service.office/postMouse")).length, mice, "the selection kept");
});

t.test("LibreOffice editing: not offered by a service without sessions", async function() {
	await open(scenario({service: true, office: true}));
	await page.waitFor("FileTypes.converter === true", 10000);
	await openDoc("/media/internal/sample.pptx");
	await page.eval("new Promise(function(r){setTimeout(r, 300);})");
	await page.eval("documentsApp.$.appMenu.open(); 1");
	await page.waitFor("documentsApp.$.appMenu.isOpen");
	assert.strictEqual(await page.eval("documentsApp.$.menuLok.showing"), false);
});

// --- Cloud ------------------------------------------------------------------------

//* The names the cloud view lists, once it has its answer.
async function cloudNames() {
	await page.waitFor("!documentsApp.$.cloud.loading", 5000);
	return page.eval("Array.prototype.map.call(document.querySelectorAll('.cloud .library-name'), function(e){return e.textContent;})");
}

//* A real tap on the cloud view's row called inName.
async function tapCloud(inName) {
	var names = await cloudNames();
	assert.ok(names.indexOf(inName) >= 0, inName + " listed: " + names.join(", "));
	await page.tap(".cloud .cloud-item:nth-child(" + (names.indexOf(inName) + 1) + ")");
}

//* From the library, through the Cloud button, into Dropbox's Work folder.
async function cloudWorkFolder() {
	await libraryNames();
	await page.tap(".library-cloud");
	await page.waitFor("documentsApp.$.pane.getViewName() === 'cloud'");
	await tapCloud("Dropbox");
	await tapCloud("Work");
	await cloudNames();
}

async function back() {
	await page.eval("documentsApp.goBack(); 1");
	await browserLib.sleep(50);
}

t.test("cloud: library, Cloud, an account, a folder; a document is downloaded into the cache and shown", async function() {
	await open(scenario({cloud: true}));
	await libraryNames();
	await page.tap(".library-cloud");
	await page.waitFor("documentsApp.$.pane.getViewName() === 'cloud'");
	assert.deepStrictEqual(await cloudNames(), ["Box", "Dropbox"], "the accounts with documents, not the photos-only one");
	// Each with its provider's picture from Accounts, the cloud where there is none
	assert.deepStrictEqual(await page.eval("Array.prototype.map.call(document.querySelectorAll('.cloud .library-icon'), " +
		"function(e){return e.className.replace(/^library-icon /, '') + ' ' + e.style.backgroundImage;})"),
		["cloud-icon-account ", "cloud-icon-provider url(\"file:///usr/palm/public/accounts/com.palm.dropbox/images/d%27b%20%2848%29.png\")"]);
	var accounts = await calls("luna", "accounts/listAccounts");
	assert.deepStrictEqual(accounts[0].params, {capability: "DOCUMENTS"});
	assert.strictEqual(accounts[0].acg, "granted", "appinfo.json asks for accountsservice.operation");
	await tapCloud("Dropbox");
	assert.deepStrictEqual(await cloudNames(), ["Work", "Welcome.pdf"], "folders first; only documents");
	await tapCloud("Work");
	assert.deepStrictEqual(await cloudNames(), ["Reports", "Budget (old).xlsx", "Budget.xlsx", "notes.txt", "Plan.odt"]);
	assert.deepStrictEqual((await calls("luna", "dropbox/listFolder")).map(function(c) {
		return c.params;
	}), [{accountId: "acc1"}, {accountId: "acc1", path: "/Work"}], "the top folder without a path, then the entry's path");
	assert.strictEqual(await text(".cloud-header .library-title"), "Work");
	assert.strictEqual(await text(".cloud-subtitle"), "Dropbox \u00b7 Work");
	// Filtered as the library is; folders stay
	await page.tap(".cloud-toolbar .library-filters .enyo-grouped-toolbutton:nth-child(3)");
	assert.deepStrictEqual(await cloudNames(), ["Reports", "Budget (old).xlsx", "Budget.xlsx"]);
	await page.tap(".cloud-toolbar .library-filters .enyo-grouped-toolbutton:nth-child(1)");
	await tapCloud("Plan.odt");
	await page.waitFor("documentsApp.$.cloud.isBusy() && /Downloading Plan\\.odt/.test(documentsApp.$.cloud.$.busyText.content)");
	assert.strictEqual(await waitOpened(), "odt");
	var d = await calls("luna", "dropbox/downloadFile");
	assert.strictEqual(d.length, 1);
	assert.deepStrictEqual(d[0].params, {accountId: "acc1", path: "/Work/Plan.odt", localPath: Cloud.cachePath(CLOUD.dropbox, CLOUD.plan)});
	assert.strictEqual(d[0].acg, "granted");
	assert.ok(/^\/media\/internal\/\.documents\/cloud\/acc1\/[0-9a-f]{8}\/Plan\.odt$/.test(d[0].params.localPath), d[0].params.localPath);
	assert.strictEqual(await text(".viewer-title"), "Plan.odt");
	assert.ok(/ \u00b7 Dropbox \u00b7 Work$/.test(await text(".viewer-subtitle")), await text(".viewer-subtitle"));
	assert.strictEqual(await page.eval("documentsApp.$.cloud.isBusy()"), false);
	// It is recent, and says where it is from
	var recent = await page.eval("JSON.parse(localStorage.getItem('documents.prefs')).recent[0]");
	assert.strictEqual(recent.path, d[0].params.localPath);
	assert.deepStrictEqual([recent.cloud.account.id, recent.cloud.folderPath], ["acc1", "/Work"]);
});

t.test("cloud: back goes from the document to its folder, up the folders, to the accounts and the library", async function() {
	await open(scenario({cloud: true}));
	await cloudWorkFolder();
	await tapCloud("Reports");
	assert.deepStrictEqual(await cloudNames(), []);
	assert.ok(/empty/.test(await text(".cloud-status")), await text(".cloud-status"));
	assert.strictEqual(await text(".cloud-subtitle"), "Dropbox \u00b7 Work \u203a Reports");
	await back();
	assert.deepStrictEqual(await cloudNames(), ["Reports", "Budget (old).xlsx", "Budget.xlsx", "notes.txt", "Plan.odt"]);
	await tapCloud("notes.txt");
	assert.strictEqual(await waitOpened(), "txt");
	await back();
	assert.strictEqual(await page.eval("documentsApp.$.pane.getViewName()"), "cloud", "back from the document: its folder");
	assert.strictEqual(await text(".cloud-header .library-title"), "Work");
	// The toolbar's arrow goes up too
	await tapControl("documentsApp.$.cloud.$.backButton");
	assert.deepStrictEqual(await cloudNames(), ["Work", "Welcome.pdf"]);
	await back();
	assert.deepStrictEqual(await cloudNames(), ["Box", "Dropbox"]);
	assert.strictEqual((await calls("luna", "accounts/listAccounts")).length, 1, "the accounts are not asked for again");
	await back();
	assert.strictEqual(await page.eval("documentsApp.$.pane.getViewName()"), "library");
	// Back while a document downloads gives up waiting for it
	await page.tap(".library-cloud");
	await tapCloud("Dropbox");
	await tapCloud("Work");
	await tapCloud("Plan.odt");
	await page.waitFor("documentsApp.$.cloud.isBusy()");
	await back();
	assert.strictEqual(await page.eval("documentsApp.$.cloud.isBusy()"), false);
	await browserLib.sleep(500);
	assert.strictEqual(await page.eval("documentsApp.$.pane.getViewName()"), "cloud", "the download that came late is not opened");
	assert.strictEqual(await text(".cloud-header .library-title"), "Work");
});

t.test("cloud: a recently opened cloud document opens its downloaded copy and says where it is from", async function() {
	var local = Cloud.cachePath(CLOUD.dropbox, CLOUD.plan);
	await open(scenario({cloud: true, storage: {clear: true, set: {"documents.prefs": {recent: [{path: local, name: "Plan.odt", kind: "odt",
		opened: Date.now(), cloud: {account: CLOUD.dropbox, folderPath: "/Work", folderName: "Work", place: "Dropbox \u00b7 Work"}}]}}}}));
	await libraryNames();
	assert.ok(/^Dropbox \u00b7 /.test(await text(".library-item .library-sub")), await text(".library-item .library-sub"));
	await page.tap(".library-item");
	assert.strictEqual(await waitOpened(), "odt");
	assert.ok(/ \u00b7 Dropbox \u00b7 Work$/.test(await text(".viewer-subtitle")), await text(".viewer-subtitle"));
	assert.strictEqual((await calls("luna", "downloadFile")).length, 0, "no download");
	await back();
	assert.strictEqual(await page.eval("documentsApp.$.pane.getViewName()"), "library", "opened from the library, back to it");
});

t.test("cloud: errors and no accounts are said plainly; Try Again asks again", async function() {
	await open(scenario({cloud: true}));
	await libraryNames();
	await page.tap(".library-cloud");
	await tapCloud("Box");
	assert.deepStrictEqual(await cloudNames(), []);
	assert.strictEqual(await text(".cloud-status .library-scanning, .cloud-status div"), "Sign in to this account again in Accounts.");
	assert.ok(await page.eval("documentsApp.$.cloud.$.retryButton.showing"));
	await tapControl("documentsApp.$.cloud.$.retryButton");
	await cloudNames();
	assert.strictEqual((await calls("luna", "boxnet/listFolder")).length, 2);
	// A download that fails is reported, and the folder stays
	await back();
	await tapCloud("Dropbox");
	await page.eval("__MOCK__.set('luna', 'luna://com.palm.service.dropbox/downloadFile', {error: {errorCode: 'PERMISSION_DENIED', " +
		"errorText: \"caller 'x' is not permitted\"}}); 1");
	await tapCloud("Welcome.pdf");
	await page.waitFor("documentsApp.$.toaster.showing");
	assert.ok(/Welcome\.pdf could not be downloaded\. This account's service does not let Documents use it\./.test(
		await page.eval("documentsApp.$.toaster.content")), await page.eval("documentsApp.$.toaster.content"));
	assert.strictEqual(await page.eval("documentsApp.$.pane.getViewName()"), "cloud");

	await open(scenario({cloud: true, luna: {"luna://com.palm.service.accounts/listAccounts": {json: {results: []}}}}));
	await libraryNames();
	await page.tap(".library-cloud");
	assert.deepStrictEqual(await cloudNames(), []);
	assert.ok(/No cloud accounts with documents/.test(await text(".cloud-status")));
});

t.test("cloud: Save a Copy of a downloaded workbook uploads it into the folder it came from", async function() {
	await open(scenario({cloud: true, service: true}));
	await cloudWorkFolder();
	await tapCloud("Budget.xlsx");
	assert.strictEqual(await waitOpened(), "xlsx");
	await page.waitFor("documentsApp.$.viewer.$.sheet.$.editButton.showing");
	await page.eval("var s=documentsApp.$.viewer.$.sheet; s.$.grid.select(1,0); s.startEdit(); s.$.editor.setValue('Fountain pen'); s.commitEdit(); 1");
	await page.waitFor("documentsApp.$.viewer.$.saveButton.showing");
	await page.tap(".viewer-save");
	await page.waitFor("documentsApp.$.saveDialog.isOpen");
	assert.strictEqual(await page.eval("documentsApp.$.saveName.getValue()"), "Budget (edited).xlsx");
	assert.ok(/same folder in Dropbox/.test(await page.eval("documentsApp.$.saveNote.content")));
	// A name the folder has: nothing is written or uploaded over it
	await page.eval("documentsApp.$.saveName.setValue('budget (OLD).xlsx'); documentsApp.doSave(); 1");
	await page.waitFor("documentsApp.$.saveError.showing");
	assert.strictEqual(await page.eval("documentsApp.$.saveError.content"), "A file with that name already exists.");
	assert.strictEqual((await calls("luna", "service/writeNew")).length, 0);
	// The connector refuses: said, and the dialog stays
	await page.eval("__MOCK__.set('luna', 'luna://com.palm.service.dropbox/uploadFile', {error: {errorCode: 'NO_STORED_CREDENTIALS'}}); 1");
	await page.eval("documentsApp.$.saveName.setValue('Budget (edited).xlsx'); documentsApp.doSave(); 1");
	await page.waitFor("documentsApp.$.saveError.showing && /uploaded/.test(documentsApp.$.saveError.content)");
	assert.strictEqual(await page.eval("documentsApp.$.saveError.content"),
		"The copy could not be uploaded. Sign in to this account again in Accounts.");
	assert.ok(await page.eval("documentsApp.$.saveDialog.isOpen"));
	// And then it takes it
	await page.eval("__MOCK__.set('luna', 'luna://com.palm.service.dropbox/uploadFile', {json: {id: 'id:new'}}); documentsApp.doSave(); 1");
	await page.waitFor("!documentsApp.$.saveDialog.isOpen", 10000);
	var w = await calls("luna", "service/writeNew");
	var u = await calls("luna", "dropbox/uploadFile");
	assert.strictEqual(w.length, 2);
	assert.strictEqual(u.length, 2);
	var dir = path.posix.dirname(Cloud.cachePath(CLOUD.dropbox, CLOUD.budget));
	assert.ok(w[1].params.path.indexOf(dir + "/copy-") === 0 && /-Budget \(edited\)\.xlsx$/.test(w[1].params.path), w[1].params.path);
	assert.notStrictEqual(w[0].params.path, w[1].params.path, "each try writes a copy of its own");
	assert.deepStrictEqual(u[1].params, {accountId: "acc1", path: "/Work", localPath: w[1].params.path, name: "Budget (edited).xlsx",
		mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
	assert.strictEqual(u[1].acg, "granted");
	var wb = XLSX.read(Buffer.from(w[1].params.data, "base64"), {type: "buffer"});
	assert.strictEqual(wb.Sheets.Prices.A2.v, "Fountain pen");
	// The folder was looked at before each upload, for the name
	assert.deepStrictEqual((await calls("luna", "dropbox/listFolder")).slice(-3).map(function(c) {
		return c.params.path;
	}), ["/Work", "/Work", "/Work"]);
	assert.strictEqual(await page.eval("documentsApp.$.toaster.content"), "Saved as Budget (edited).xlsx in Dropbox \u00b7 Work");
});

// --- Layouts ------------------------------------------------------------------------

[["tablet", 1024, 768], ["phone", 480, 800], ["square", 450, 380]].forEach(function(inSize) {
	t.test("layout " + inSize[0] + " " + inSize[1] + "x" + inSize[2] + ": nothing wider than the screen, tools fit", async function() {
		// With the service, so the edit tool is in the toolbar too.
		await open(scenario({service: true, cloud: true}), {width: inSize[1], height: inSize[2]});
		await libraryNames();
		var lib = await page.eval("({doc: document.documentElement.scrollWidth, toolbar: document.querySelector('.library-toolbar').scrollWidth, " +
			"cloud: document.querySelector('.library-cloud').getBoundingClientRect().right})");
		assert.ok(lib.doc <= inSize[1] && lib.toolbar <= inSize[1] && lib.cloud <= inSize[1], JSON.stringify(lib));
		// Each filter's caption fits inside its button (the theme's 16px ends take room from the width).
		var clipped = function(inView) {
			return page.eval("Array.prototype.filter.call(document.querySelectorAll('" + inView + " .library-filters .enyo-grouped-toolbutton'), " +
				"function(b){var r=document.createRange();r.selectNodeContents(b);var t=r.getBoundingClientRect().width;" +
				"var cs=getComputedStyle(b);var inner=b.getBoundingClientRect().width-parseFloat(cs.borderLeftWidth)-parseFloat(cs.borderRightWidth);" +
				"return t>inner+0.5;}).map(function(b){return b.textContent;})");
		};
		assert.deepStrictEqual(await clipped(".library"), [], "filter captions fit");
		// The cloud view's toolbar has its back key where the library has Cloud
		await page.eval("documentsApp.showCloud(); 1");
		await page.waitFor("!documentsApp.$.cloud.loading");
		var cloud = await page.eval("({doc: document.documentElement.scrollWidth, toolbar: document.querySelector('.cloud-toolbar').scrollWidth, " +
			"last: document.querySelector('.cloud-toolbar .library-refresh').getBoundingClientRect().right})");
		assert.ok(cloud.doc <= inSize[1] && cloud.toolbar <= inSize[1] && cloud.last <= inSize[1], JSON.stringify(cloud));
		assert.deepStrictEqual(await clipped(".cloud"), [], "the cloud view's filter captions fit");
		await page.eval("documentsApp.showLibrary(); 1");
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
