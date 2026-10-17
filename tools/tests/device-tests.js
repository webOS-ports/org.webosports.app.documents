/*
	On-device tests: the installed app in the real WAM on a LuneOS device,
	with the real File Manager service and the documents
	tools/test-device.sh copied to /media/internal/documents-test/. Run
	through that script, which installs and launches the app and forwards
	WAM's DevTools port; or by hand:

		DEVICE_DEBUG=http://127.0.0.1:19461 DEVICE_NAME=touchpad DEVICE_SERIAL=<serial> \
			APP_ID=org.webosports.app.documents node tools/tests/device-tests.js

	Each run records the screen, opens every fixture and the large
	documents (when they were pushed) and measures how long they take and
	the longest time the page did not answer (polled every 250 ms), opens a
	document the way Files does (applicationManager/open), and checks the
	layout at that screen size. Results: test-output/device/<name>/results.json
	and a screenshot per document.
*/
var fs = require("fs");
var path = require("path");
var execFile = require("child_process").execFile;
var t = require("./lib/harness.js");
var device = require("./lib/device.js");
var assert = t.assert;

var base = process.env.DEVICE_DEBUG || "http://127.0.0.1:19461";
var name = process.env.DEVICE_NAME || "device";
var serial = process.env.DEVICE_SERIAL || "";
var appId = process.env.APP_ID || "org.webosports.app.documents";
var DIR = "/media/internal/documents-test";
var outDir = path.join(__dirname, "..", "..", "test-output", "device", name);
var APP = new RegExp(appId.replace(/\./g, "\\.") + "/index\\.html");
var results = {device: name, appId: appId, documents: {}};

var page;

function sleep(inMs) {
	return new Promise(function(r) {
		setTimeout(r, inMs);
	});
}

function adb(inCommand) {
	return new Promise(function(resolve) {
		execFile("adb", ["-s", serial, "shell", inCommand], {timeout: 60000}, function(err, out) {
			resolve(String(out || ""));
		});
	});
}

function shot(inName) {
	return page.screenshot(path.join(outDir, inName + ".png"));
}

/**
	Opens a document and polls until it is shown and fully read; returns
	{timings, kind, pollFirst, longestBlock, gaps}. A long gap between two
	polls is time the page could not answer.
*/
async function openAndMeasure(inPath, inTimeout) {
	await page.eval("documentsApp.openDocument(null, " + JSON.stringify(inPath) + "); 1");
	var t0 = Date.now(), last = t0, longest = 0, gaps = [], first = null, r = null;
	while (Date.now() - t0 < (inTimeout || 240000)) {
		r = JSON.parse(await page.eval("(function(){var v=documentsApp.$.viewer; return JSON.stringify({path: v.path, t: v.timings, " +
			"kind: v.doc && v.doc.kind, err: v.lastError && !v.doc ? v.lastError : null, rendering: !!v.renderingAll, " +
			"streaming: !!v.streaming});})()"));
		var now = Date.now();
		longest = Math.max(longest, now - last);
		if (now - last > 1000) {
			gaps.push([last - t0, now - last]);
		}
		last = now;
		if (r.path === inPath && (r.kind || r.err)) {
			if (r.err) {
				break;
			}
			if (r.t.total && first === null) {
				first = now - t0;
			}
			if (r.t.total && !r.rendering && !r.streaming) {
				break;
			}
		}
		await sleep(250);
	}
	return {timings: r.t, kind: r.kind, error: r.err, pollFirst: first, longestBlock: longest, gaps: gaps};
}

t.test("device and screen", async function() {
	var info = await page.eval("({width:innerWidth,height:innerHeight,dpr:devicePixelRatio,screen:screen.width+'x'+screen.height," +
		"ua:navigator.userAgent.replace(/^.*?Chrome\\/([0-9.]+).*$/,'Chrome $1'),cores:navigator.hardwareConcurrency," +
		"palm:!!window.PalmSystem})");
	results.screen = info;
	console.log("     " + name + ": " + info.width + "x" + info.height + " CSS px, devicePixelRatio " + info.dpr + ", screen " +
		info.screen + ", " + info.ua + ", " + info.cores + " cores");
	assert.ok(info.palm, "not running under WAM (no PalmSystem)");
});

t.test("starts without errors and finds the documents on USB storage", async function() {
	await page.reload();
	await page.waitFor("window.documentsApp && documentsApp.hasNode()", 60000);
	var t0 = Date.now();
	await page.waitFor("!documentsApp.$.library.scanning", 120000);
	results.libraryScanMs = Date.now() - t0;
	var docs = await page.eval("documentsApp.$.library.docs.map(function(d){return d.path;})");
	results.libraryCount = docs.length;
	console.log("     library: " + docs.length + " documents in " + results.libraryScanMs + " ms");
	["sample.odt", "sample.docx", "sample.xlsx", "sample.pptx", "semicolon.csv"].forEach(function(f) {
		assert.ok(docs.indexOf(DIR + "/" + f) >= 0, f + " found");
	});
	await shot("01-library");
	assert.deepStrictEqual(page.errors, []);
});

var samples = [
	["sample.odt", "odt", ".doc-page", /Fixture heading/],
	["sample.docx", "docx", ".doc-docx", /Fixture heading/],
	["sample.xlsx", "xlsx", ".sheet-cells", /Book/],
	["sample.xls", "xls", ".sheet-cells", /Book/],
	["sample.ods", "ods", ".sheet-cells", /Book/],
	["sample.pptx", "pptx", ".doc-slides", /Fixture slides/],
	["sample.odp", "odp", ".doc-slides", /Fixture slides/],
	["sample.md", "md", ".doc-markdown-body", /Markdown fixture/],
	["sample.txt", "txt", ".doc-plaintext", /findme/],
	["semicolon.csv", "csv", ".sheet-cells", /Smith; J\./]
];

samples.forEach(function(s, i) {
	t.test("opens " + s[0], async function() {
		var m = await openAndMeasure(DIR + "/" + s[0], 60000);
		results.documents[s[0]] = m;
		console.log("     " + s[0] + ": " + JSON.stringify(m.timings) + ", longest unresponsive " + m.longestBlock + " ms");
		assert.strictEqual(m.kind, s[1], JSON.stringify(m.error));
		var text = await page.eval("(function(){var e=document.querySelector(" + JSON.stringify(s[2]) + ");return e?e.textContent:'';})()");
		assert.ok(s[3].test(text), s[0] + " shows its content");
		await sleep(400);
		await shot((10 + i) + "-" + s[0].replace(/\./g, "-"));
		assert.deepStrictEqual(page.errors, []);
	});
});

t.test("layout: page, toolbar and grid fit the screen", async function() {
	await openAndMeasure(DIR + "/sample.odt", 60000);
	var m = await page.eval("({w: innerWidth, page: document.querySelector('.doc-page').getBoundingClientRect().right, " +
		"tool: document.querySelector('.viewer-toolbar .viewer-tool:last-child').getBoundingClientRect().right, " +
		"title: document.querySelector('.viewer-title').getBoundingClientRect().right, reflow: documentsApp.$.viewer.reflow})");
	results.layout = m;
	assert.ok(m.page <= m.w + 1 && m.tool <= m.w && m.title <= m.w, JSON.stringify(m));
});

t.test("find in a document", async function() {
	await openAndMeasure(DIR + "/sample.odt", 60000);
	await page.eval("documentsApp.$.viewer.openFind(); documentsApp.$.viewer.$.findInput.setValue('findme'); documentsApp.$.viewer.runFind(); 1");
	assert.strictEqual(await page.eval("document.querySelectorAll('mark.find-hit').length"), 3);
	await shot("30-find");
	await page.eval("documentsApp.$.viewer.closeFind(); 1");
});

t.test("opened as Files opens it: applicationManager/open with a file URL", async function() {
	if (!serial) {
		console.log("     (skipped: no DEVICE_SERIAL)");
		return;
	}
	await page.eval("documentsApp.$.pane.selectViewByName('library'); 1");
	await adb("luna-send -n 1 -w 8000 luna://com.palm.applicationManager/open '{\"target\":\"file://" + DIR + "/sample.pptx\"}' > /tmp/documents-open.out 2>&1");
	var out = await adb("cat /tmp/documents-open.out; rm -f /tmp/documents-open.out");
	results.openedByMime = out.trim();
	console.log("     applicationManager/open answered: " + out.trim());
	// The handler for .pptx may be another app on this device (or none): say which.
	var shown = null;
	try {
		shown = await page.waitFor("documentsApp.$.viewer.path === " + JSON.stringify(DIR + "/sample.pptx") + " && documentsApp.$.viewer.doc && documentsApp.$.viewer.doc.kind", 30000);
	} catch (e) {}
	results.openedByMimeShown = shown;
	assert.strictEqual(shown, "pptx", "Documents was not the app that opened the .pptx (" + out.trim() + ")");
	await shot("40-opened-like-files");
});

t.test("a PDF goes to the app registered for application/pdf", async function() {
	if (!serial) {
		console.log("     (skipped: no DEVICE_SERIAL)");
		return;
	}
	var running = async function() {
		await adb("luna-send -n 1 -w 8000 luna://com.webos.service.applicationmanager/running '{}' > /tmp/documents-run.out 2>&1");
		return adb("cat /tmp/documents-run.out; rm -f /tmp/documents-run.out");
	};
	await adb("luna-send -n 1 -w 8000 luna://com.palm.applicationManager/listAllHandlersForMime '{\"mime\":\"application/pdf\"}' > /tmp/documents-h.out 2>&1");
	var handler = ((await adb("cat /tmp/documents-h.out; rm -f /tmp/documents-h.out")).match(/"activeHandler"[^}]*"appId": *"([^"]+)"/) || [])[1];
	results.pdfHandler = handler || null;
	if (!handler) {
		console.log("     (skipped: no app on this device handles application/pdf)");
		return;
	}
	await adb("luna-send -n 1 -w 8000 luna://com.webos.service.applicationmanager/closeByAppId '{\"id\":\"" + handler + "\"}' > /dev/null 2>&1; sleep 2");
	await page.eval("documentsApp.$.pane.selectViewByName('library'); documentsApp.openDocument(null, " + JSON.stringify(DIR + "/hand off.pdf") + "); 1");
	var started = false;
	for (var i = 0; i < 20 && !started; i++) {
		await sleep(1000);
		started = (await running()).indexOf("\"id\":\"" + handler + "\"") >= 0;
	}
	results.pdfHandedTo = started ? handler : null;
	console.log("     " + (started ? handler + " started" : handler + " did not start") + "; Documents shows its " + await page.eval("documentsApp.$.pane.getViewName()"));
	await adb("luna-send -n 1 -w 10000 luna://com.webos.surfacemanager/captureCompositorOutput '{\"output\":\"/tmp/documents-pdf.png\",\"format\":\"PNG\"}' > /dev/null 2>&1");
	await new Promise(function(resolve) {
		execFile("adb", ["-s", serial, "pull", "/tmp/documents-pdf.png", path.join(outDir, "45-pdf-handed-on.png")], function() {
			resolve();
		});
	});
	await adb("rm -f /tmp/documents-pdf.png; luna-send -n 1 -w 8000 luna://com.webos.service.applicationmanager/closeByAppId '{\"id\":\"" + handler + "\"}' > /dev/null 2>&1");
	assert.ok(started, handler + " was not started");
	assert.strictEqual(await page.eval("documentsApp.$.pane.getViewName()"), "library");
	assert.deepStrictEqual(page.errors, []);
});

["large.docx", "large.odt", "large.xlsx", "large.ods", "large.csv"].forEach(function(f) {
	t.test("large document: " + f, async function() {
		var present = serial ? await adb("ls " + DIR + "/" + f + " 2>/dev/null") : "";
		if (serial && !present.trim()) {
			console.log("     (skipped: " + f + " not on the device)");
			return;
		}
		var m = await openAndMeasure(DIR + "/" + f, 300000);
		results.documents[f] = m;
		console.log("     " + f + ": shown after " + (m.timings && m.timings.total) + " ms, all read after " +
			(m.timings && (m.timings.complete || m.timings.total)) + " ms, longest unresponsive " + m.longestBlock + " ms, gaps " +
			JSON.stringify(m.gaps));
		assert.ok(m.kind, JSON.stringify(m.error));
		await shot("50-" + f.replace(/\./g, "-"));
		assert.deepStrictEqual(page.errors, []);
	});
});

(async function() {
	fs.mkdirSync(outDir, {recursive: true});
	page = await device.attach(base, APP, 60000);
	await t.run({
		after: async function() {
			if (serial) {
				var mem = await adb("free -m | head -n 2");
				results.memoryAfter = mem.trim();
				var crashes = await adb("journalctl --no-pager --since '-30min' 2>/dev/null | grep -c WEBPROC_CRASH");
				results.webprocCrashesLast30min = +crashes.trim() || 0;
				console.log("     WAM renderer crashes in the last 30 minutes: " + results.webprocCrashesLast30min);
			}
			fs.writeFileSync(path.join(outDir, "results.json"), JSON.stringify(results, null, "\t") + "\n");
			console.log("results: " + path.join(outDir, "results.json"));
			page.close();
		}
	});
})().catch(function(e) {
	console.error(e);
	process.exit(1);
});
