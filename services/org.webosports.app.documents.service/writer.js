/*
	The part of the service that touches files, kept apart from the bus so
	the unit tests run it on the build host (tools/tests/unit-tests.js).

	writeNew(root, path, base64) creates path and writes the bytes, and
	never replaces or changes a file that exists: the file is opened with
	O_CREAT|O_EXCL ("wx"), so an existing name fails with EEXIST. Only
	paths under root (/media/internal/ on a device) with the extension of a
	format the app writes are accepted, and no hidden file or folder but
	the app's cloud cache (.documents/cloud/), where the copy of a
	downloaded document is written before it is uploaded, and its print
	folder (.documents/print/).

	printCopy(root, name, base64) writes an edited document that is to be
	printed (LibreOffice makes the PDF from it) as root/.documents/print/name,
	after removing the copy printed before: the folder holds one at most.
*/
var fs = require("fs");

var MAX_BYTES = 64 * 1024 * 1024;

//* Below root: the hidden folders copies are written to (source/services/CloudService.js, printCopy).
var CLOUD_CACHE = ".documents/cloud/";
var PRINT_CACHE = ".documents/print/";

//* The path if it may be written, else null.
function checkPath(inRoot, inPath) {
	var p = String(inPath || "");
	var rest = p.slice(inRoot.length);
	// Below the cloud cache, the rest of the path is held to the same rules
	var below = rest.indexOf(CLOUD_CACHE) === 0 ? rest.slice(CLOUD_CACHE.length) :
		rest.indexOf(PRINT_CACHE) === 0 ? rest.slice(PRINT_CACHE.length) : rest;
	if (p.indexOf(inRoot) !== 0 || !below || p.indexOf("\u0000") >= 0 || /(^|\/)\.\.?(\/|$)/.test(rest) ||
			/(^|\/)\./.test(below) || !/\.(xlsx|ods|csv|odt|docx|txt|md|pptx|odp)$/i.test(p)) {
		return null;
	}
	return p;
}

function result(inOk, inCode, inText, inExtra) {
	var r = inOk ? {returnValue: true} : {returnValue: false, errorCode: inCode, errorText: inText};
	for (var k in inExtra || {}) {
		if (inExtra.hasOwnProperty(k)) {
			r[k] = inExtra[k];
		}
	}
	return r;
}

function writeNew(inRoot, inPath, inBase64) {
	var path = checkPath(inRoot, inPath);
	if (!path) {
		return result(false, "EINVAL", "Only new .xlsx, .ods, .csv, .odt, .docx, .pptx, .odp, .txt or .md files under " + inRoot + " can be written");
	}
	if (typeof inBase64 !== "string" || !inBase64.length || inBase64.length > Math.ceil(MAX_BYTES / 3) * 4) {
		return result(false, "EINVAL", "No data, or too much");
	}
	var bytes = Buffer.from(inBase64, "base64");
	var fd = null;
	try {
		// "wx": create, and fail if the name exists. Nothing is ever overwritten.
		fd = fs.openSync(path, "wx", 420);
		var off = 0;
		while (off < bytes.length) {
			off += fs.writeSync(fd, bytes, off, bytes.length - off, null);
		}
		fs.fsyncSync(fd);
		fs.closeSync(fd);
		fd = null;
		return result(true, null, null, {path: path, bytes: bytes.length});
	} catch (e) {
		if (fd !== null) {
			try {
				fs.closeSync(fd);
			} catch (e2) {}
			// A half-written file is no copy: remove the file this call created (never another one).
			try {
				fs.unlinkSync(path);
			} catch (e3) {}
		}
		return result(false, e.code || "EIO", String(e.message || e));
	}
}

function printCopy(inRoot, inName, inBase64) {
	var name = String(inName || "");
	var dir = inRoot + PRINT_CACHE;
	if (!name || name.indexOf("/") >= 0 || !checkPath(inRoot, dir + name)) {
		return result(false, "EINVAL", "Only a .xlsx, .ods, .csv, .odt, .docx, .pptx, .odp, .txt or .md file name, without a folder");
	}
	try {
		// The storage is the user's (USB mode): a folder on the way that is a link could point
		// anywhere, and what is in it would be removed. Each is made, or found, and looked at
		// before the next one is made in it.
		[inRoot + ".documents", dir.replace(/\/$/, "")].forEach(function(inDir) {
			try {
				fs.mkdirSync(inDir, 493);
			} catch (e) {
				if (e.code !== "EEXIST") {
					throw e;
				}
			}
			if (!fs.lstatSync(inDir).isDirectory()) {
				throw Object.assign(new Error(inDir + " is not a folder"), {code: "ENOTDIR"});
			}
		});
		// The copy printed before: only the files directly in the folder, never what a link points to
		fs.readdirSync(dir, {withFileTypes: true}).forEach(function(inEntry) {
			if (inEntry.isFile() || inEntry.isSymbolicLink()) {
				fs.unlinkSync(dir + inEntry.name);
			}
		});
	} catch (e) {
		return result(false, e.code || "EIO", String(e.message || e));
	}
	return writeNew(inRoot, dir + name, inBase64);
}

module.exports = {checkPath: checkPath, writeNew: writeNew, printCopy: printCopy, MAX_BYTES: MAX_BYTES};
