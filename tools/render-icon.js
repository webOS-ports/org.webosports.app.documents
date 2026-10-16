/*
	Renders images/icon.svg to icon.png (64x64) and icon-256x256.png with
	headless Chromium, so the PNGs match the SVG exactly.

		node tools/render-icon.js
*/
var path = require("path");
var fs = require("fs");
var browser = require("./tests/lib/browser.js");

(async function() {
	var root = path.join(__dirname, "..");
	var svg = fs.readFileSync(path.join(root, "images", "icon.svg"), "utf8");
	var b = await browser.launch();
	try {
		var sizes = [[64, "icon.png"], [256, "icon-256x256.png"]];
		for (var i = 0; i < sizes.length; i++) {
			var s = sizes[i][0];
			var page = await b.newPage({width: s, height: s});
			await page.send("Emulation.setDefaultBackgroundColorOverride", {color: {r: 0, g: 0, b: 0, a: 0}});
			var html = "<html><body style='margin:0;background:transparent'><img width='" + s + "' height='" + s +
				"' src='data:image/svg+xml;base64," + Buffer.from(svg).toString("base64") + "'></body></html>";
			await page.goto("data:text/html;base64," + Buffer.from(html).toString("base64"));
			await page.screenshot(path.join(root, sizes[i][1]));
			await page.close();
		}
	} finally {
		await b.close();
	}
})().catch(function(e) {
	console.error(e);
	process.exit(1);
});
