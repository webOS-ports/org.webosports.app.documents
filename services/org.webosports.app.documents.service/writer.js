/*
	The part of the service that touches files, kept apart from the bus so
	the unit tests run it on the build host (tools/tests/unit-tests.js).

	writeNew(root, path, base64) creates path and writes the bytes, and
	never replaces or changes a file that exists: the file is opened with
	O_CREAT|O_EXCL ("wx"), so an existing name fails with EEXIST. Only
	paths under root (/media/internal/ on a device) with the extension of a
	format the app writes are accepted.
*/
var fs = require("fs");

var MAX_BYTES = 64 * 1024 * 1024;

//* The path if it may be written, else null.
function checkPath(inRoot, inPath) {
	var p = String(inPath || "");
	var rest = p.slice(inRoot.length);
	if (p.indexOf(inRoot) !== 0 || !rest || p.indexOf("\u0000") >= 0 || /(^|\/)\.\.?(\/|$)/.test(rest) ||
			/(^|\/)\./.test(rest) || !/\.(xlsx|ods|csv|odt|docx|txt|md|pptx|odp)$/i.test(p)) {
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

module.exports = {checkPath: checkPath, writeNew: writeNew, MAX_BYTES: MAX_BYTES};
