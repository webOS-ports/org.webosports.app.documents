/**
	Which kind of document a file is: from its name first, then from its
	first bytes and, for zip containers, from the entries inside. A file
	whose name and content disagree (a ".docx" that is really an ODF text,
	say) is opened as what its content is.

	Kinds and the family that shows them:
		docx, odt, txt, md   "text"    (flowing pages)
		xlsx, xls, ods, csv  "sheet"   (a grid)
		pptx, odp            "slides"
	plus "doc", "ppt" and "rtf", which are recognised only to say they are
	not supported, and "unknown".

	Pure logic, no DOM: tested in node (tools/tests/unit-tests.js).
*/
(function(root) {
	var kinds = {
		docx: {family: "text", label: "Word document", mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
		odt: {family: "text", label: "OpenDocument text", mime: "application/vnd.oasis.opendocument.text"},
		txt: {family: "text", label: "Plain text", mime: "text/plain"},
		md: {family: "text", label: "Markdown text", mime: "text/markdown"},
		xlsx: {family: "sheet", label: "Excel workbook", mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
		xls: {family: "sheet", label: "Excel 97-2003 workbook", mime: "application/vnd.ms-excel"},
		ods: {family: "sheet", label: "OpenDocument spreadsheet", mime: "application/vnd.oasis.opendocument.spreadsheet"},
		csv: {family: "sheet", label: "Comma-separated values", mime: "text/csv"},
		pptx: {family: "slides", label: "PowerPoint presentation", mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation"},
		odp: {family: "slides", label: "OpenDocument presentation", mime: "application/vnd.oasis.opendocument.presentation"},
		doc: {family: null, label: "Word 97-2003 document", mime: "application/msword"},
		ppt: {family: null, label: "PowerPoint 97-2003 presentation", mime: "application/vnd.ms-powerpoint"},
		rtf: {family: null, label: "Rich text", mime: "application/rtf"}
	};

	//* File name extension -> kind.
	var extensions = {
		docx: "docx", docm: "docx", dotx: "docx",
		odt: "odt", ott: "odt", fodt: "odt",
		txt: "txt", text: "txt", log: "txt",
		md: "md", markdown: "md",
		xlsx: "xlsx", xlsm: "xlsx", xltx: "xlsx",
		xls: "xls",
		ods: "ods", ots: "ods", fods: "ods",
		csv: "csv", tsv: "csv",
		pptx: "pptx", pptm: "pptx", ppsx: "pptx", potx: "pptx",
		odp: "odp", otp: "odp", fodp: "odp",
		doc: "doc", ppt: "ppt", rtf: "rtf"
	};

	//* ODF "mimetype" entry -> kind.
	var odfMimes = {
		"application/vnd.oasis.opendocument.text": "odt",
		"application/vnd.oasis.opendocument.text-template": "odt",
		"application/vnd.oasis.opendocument.spreadsheet": "ods",
		"application/vnd.oasis.opendocument.spreadsheet-template": "ods",
		"application/vnd.oasis.opendocument.presentation": "odp",
		"application/vnd.oasis.opendocument.presentation-template": "odp"
	};

	function extensionOf(inName) {
		var name = String(inName || "").replace(/[?#].*$/, "");
		var slash = name.lastIndexOf("/");
		if (slash >= 0) {
			name = name.slice(slash + 1);
		}
		var dot = name.lastIndexOf(".");
		return dot > 0 ? name.slice(dot + 1).toLowerCase() : "";
	}

	function kindFromName(inName) {
		return extensions[extensionOf(inName)] || "unknown";
	}

	//* True for the kinds the viewer can open.
	function isSupported(inKind) {
		return !!(kinds[inKind] && kinds[inKind].family);
	}

	//* True when the name has an extension of a kind the viewer can open.
	function isDocumentName(inName) {
		return isSupported(kindFromName(inName));
	}

	/**
		What the first bytes say: "zip", "ole" (the 97-2003 compound file),
		"pdf", "rtf", "text" or "binary". inHead is a Uint8Array (a few KB
		is plenty).
	*/
	function sniff(inHead) {
		var b = inHead || [];
		if (b.length >= 4 && b[0] === 0x50 && b[1] === 0x4B && (b[2] === 3 || b[2] === 5 || b[2] === 7)) {
			return "zip";
		}
		if (b.length >= 8 && b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0 &&
				b[4] === 0xA1 && b[5] === 0xB1 && b[6] === 0x1A && b[7] === 0xE1) {
			return "ole";
		}
		if (b.length >= 5 && b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46 && b[4] === 0x2D) {
			return "pdf";
		}
		if (b.length >= 5 && b[0] === 0x7B && b[1] === 0x5C && b[2] === 0x72 && b[3] === 0x74 && b[4] === 0x66) {
			return "rtf";
		}
		return looksLikeText(b) ? "text" : "binary";
	}

	/**
		Text if it has no NUL bytes and few control characters (UTF-16 with
		a byte order mark counts as text too).
	*/
	function looksLikeText(inBytes) {
		var n = Math.min(inBytes.length, 4096);
		if (n >= 2 && ((inBytes[0] === 0xFF && inBytes[1] === 0xFE) || (inBytes[0] === 0xFE && inBytes[1] === 0xFF))) {
			return true;
		}
		var bad = 0;
		for (var i = 0; i < n; i++) {
			var c = inBytes[i];
			if (c === 0) {
				return false;
			}
			if (c < 9 || (c > 13 && c < 32 && c !== 27)) {
				bad++;
			}
		}
		return bad <= n / 50;
	}

	/**
		The kind of a zip container from its entry names and, for ODF, the
		content of its "mimetype" entry. Returns null for a zip that is no
		office document.
	*/
	function kindFromZip(inNames, inMimetype) {
		var mt = String(inMimetype || "").replace(/\s+/g, "");
		if (odfMimes[mt]) {
			return odfMimes[mt];
		}
		var has = {};
		for (var i = 0; i < inNames.length; i++) {
			has[inNames[i]] = true;
		}
		if (has["word/document.xml"]) {
			return "docx";
		}
		if (has["xl/workbook.xml"] || has["xl/workbook.bin"]) {
			return "xlsx";
		}
		if (has["ppt/presentation.xml"]) {
			return "pptx";
		}
		// An ODF package without the mimetype entry: guess from content.xml's neighbours.
		if (has["content.xml"] && has["META-INF/manifest.xml"]) {
			return null;
		}
		return null;
	}

	/**
		The final verdict for a file: {kind, family, label, mismatch} where
		mismatch says the name pointed elsewhere. inZip is optional:
		{names: [...], mimetype: "..."} once the container was opened.
	*/
	function resolve(inName, inHead, inZip) {
		var byName = kindFromName(inName);
		var magic = sniff(inHead);
		var kind = byName;
		if (magic === "zip") {
			var z = inZip ? kindFromZip(inZip.names || [], inZip.mimetype) : null;
			if (z) {
				kind = z;
			} else if (inZip) {
				kind = "unknown";
			} else if (["docx", "odt", "xlsx", "ods", "pptx", "odp"].indexOf(byName) < 0) {
				// Needs a look inside before it can be named.
				kind = "zip";
			}
		} else if (magic === "ole") {
			kind = byName === "doc" || byName === "ppt" ? byName : "xls";
		} else if (magic === "pdf") {
			kind = "pdf";
		} else if (magic === "rtf") {
			kind = "rtf";
		} else if (magic === "text" && /^fod[tsp]$/.test(extensionOf(inName))) {
			// Flat OpenDocument: one XML file.
			kind = byName;
		} else if (magic === "text") {
			if (["txt", "md", "csv"].indexOf(byName) < 0) {
				// An unknown or mislabelled file that is text: show it as text.
				kind = byName === "unknown" || isSupported(byName) ? "txt" : byName;
			}
		} else if (["txt", "md", "csv"].indexOf(byName) >= 0) {
			// A ".txt" with binary content is shown anyway (lossy), not refused.
			kind = byName;
		}
		var info = kinds[kind] || {family: null, label: kind === "pdf" ? "PDF document" : "Unknown file"};
		return {kind: kind, family: info.family, label: info.label, mismatch: byName !== kind && byName !== "unknown"};
	}

	//* The appinfo.json "mimeTypes" entries for every supported extension.
	function appinfoMimeTypes() {
		var out = [];
		var primary = {docx: 1, odt: 1, txt: 0, md: 1, xlsx: 1, xls: 1, ods: 1, csv: 1, pptx: 1, odp: 1};
		for (var ext in extensions) {
			if (extensions.hasOwnProperty(ext)) {
				var k = extensions[ext];
				if (isSupported(k) && ext === k && primary.hasOwnProperty(k)) {
					out.push({mime: kinds[k].mime, extension: ext});
				}
			}
		}
		return out;
	}

	var FileTypes = {
		kinds: kinds,
		extensions: extensions,
		extensionOf: extensionOf,
		kindFromName: kindFromName,
		isSupported: isSupported,
		isDocumentName: isDocumentName,
		sniff: sniff,
		looksLikeText: looksLikeText,
		kindFromZip: kindFromZip,
		resolve: resolve,
		appinfoMimeTypes: appinfoMimeTypes
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = FileTypes;
	} else {
		root.FileTypes = FileTypes;
	}
})(this);
