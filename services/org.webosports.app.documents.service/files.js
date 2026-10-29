/*
	The library's file listing, in place of the File Manager's service
	(org.webosports.service.filemanager), which the app used for it before.
	Kept apart from the bus so the unit tests run it on the build host
	(tools/tests/unit-tests.js).

	list(root, dir)   {"returnValue": true, "dirs": [names], "files": [names]}
	size(root, path)  {"returnValue": true, "size": <bytes>}
	A failure is {"returnValue": false, "errorCode": "ENOENT" | "EINVAL" | ..., "errorText"}.

	Unlike the File Manager's service:
	- names come as arrays, so a name with a comma in it arrives whole (that
	  service joined them with commas);
	- only root (/media/internal on a device) and what is below it is read,
	  also when a symbolic link points elsewhere (that service read any path);
	- an entry that cannot be looked at (a dangling link) is left out, where
	  that service failed the whole listing.
*/
var fs = require("fs");
var nodePath = require("path");

function failure(inCode, inText) {
	return {returnValue: false, errorCode: inCode, errorText: inText};
}

/**
	The path if it is root or below it, both as written and once its links
	are followed; else null. Throws what realpath throws (ENOENT, ...).
*/
function inside(inRoot, inPath) {
	var root = inRoot.replace(/\/+$/, "");
	var p = typeof inPath === "string" ? inPath.replace(/\/+$/, "") : "";
	if (!p || p.indexOf("\u0000") >= 0 || /(^|\/)\.\.?(\/|$)/.test(p) || (p !== root && p.indexOf(root + "/") !== 0)) {
		return null;
	}
	var realRoot = fs.realpathSync(root);
	var real = fs.realpathSync(p);
	return real === realRoot || real.indexOf(realRoot + "/") === 0 ? p : null;
}

function refused(inRoot) {
	return failure("EINVAL", "Only " + inRoot.replace(/\/+$/, "") + " and what is below it can be read");
}

function list(inRoot, inDir) {
	try {
		var dir = inside(inRoot, inDir);
		if (!dir) {
			return refused(inRoot);
		}
		var dirs = [], files = [];
		fs.readdirSync(dir, {withFileTypes: true}).forEach(function(inEntry) {
			var isDir = inEntry.isDirectory();
			if (inEntry.isSymbolicLink()) {
				try {
					isDir = fs.statSync(nodePath.join(dir, inEntry.name)).isDirectory();
				} catch (e) {
					// A link to nothing: left out.
					return;
				}
			}
			(isDir ? dirs : files).push(inEntry.name);
		});
		return {returnValue: true, dirs: dirs, files: files};
	} catch (e) {
		return failure(e.code || "EIO", String(e.message || e));
	}
}

function size(inRoot, inPath) {
	try {
		var p = inside(inRoot, inPath);
		if (!p) {
			return refused(inRoot);
		}
		return {returnValue: true, size: fs.statSync(p).size};
	} catch (e) {
		return failure(e.code || "EIO", String(e.message || e));
	}
}

module.exports = {list: list, size: size};
