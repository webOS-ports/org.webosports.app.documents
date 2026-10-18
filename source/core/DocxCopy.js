/**
	Tidies a Word document written by the .docx editor (SuperDoc) before it
	is saved: the editor stamps every file it writes with properties of its
	own (its version, where the file came from, and a random id for the
	document). Documents does not tag people's files, so those go again;
	properties the document had of its own stay.

		DocDocxCopy.clean(JSZip, exportedBytes, originalBytes).then(function(bytes) {...});
*/
(function(root) {
	//* The custom properties the editor adds.
	var STAMPS = ["SuperdocVersion", "SuperdocDocumentOrigin", "DocumentGuid"];

	function removeStamps(inXml) {
		return inXml.replace(/<([\w]+:)?property\b[^>]*\bname="([^"]*)"[^>]*>[\s\S]*?<\/([\w]+:)?property>/g, function(m, p, name) {
			return STAMPS.indexOf(name) >= 0 ? "" : m;
		});
	}

	function hasProperties(inXml) {
		return /<([\w]+:)?property\b/.test(inXml);
	}

	/**
		inBytes (the editor's .docx) without the editor's stamps; inOriginal
		the document as it was opened, whose own custom properties are kept.
		Resolves with a Uint8Array.
	*/
	function clean(inJSZip, inBytes, inOriginal) {
		var path = "docProps/custom.xml";
		return Promise.all([inJSZip.loadAsync(inBytes), inJSZip.loadAsync(inOriginal)]).then(function(zips) {
			var zip = zips[0], original = zips[1];
			var part = zip.file(path);
			if (!part) {
				return zip;
			}
			return part.async("string").then(function(inXml) {
				var kept = removeStamps(inXml);
				if (hasProperties(kept) || original.file(path)) {
					zip.file(path, kept);
					return zip;
				}
				// The document had no custom properties: the part goes, with its relationship and content type.
				zip.remove(path);
				var rels = zip.file("_rels/.rels"), types = zip.file("[Content_Types].xml");
				return Promise.all([rels ? rels.async("string") : null, types ? types.async("string") : null]).then(function(texts) {
					if (texts[0] !== null) {
						zip.file("_rels/.rels", texts[0].replace(/<([\w]+:)?Relationship\b[^>]*Target="\/?docProps\/custom\.xml"[^>]*\/>/g, ""));
					}
					if (texts[1] !== null) {
						zip.file("[Content_Types].xml", texts[1].replace(/<([\w]+:)?Override\b[^>]*PartName="\/docProps\/custom\.xml"[^>]*\/>/g, ""));
					}
					return zip;
				});
			});
		}).then(function(inZip) {
			return inZip.generateAsync({type: "uint8array", compression: "DEFLATE", compressionOptions: {level: 6}});
		});
	}

	var DocDocxCopy = {
		clean: clean,
		removeStamps: removeStamps,
		STAMPS: STAMPS
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocDocxCopy;
	} else {
		root.DocDocxCopy = DocDocxCopy;
	}
})(this);
