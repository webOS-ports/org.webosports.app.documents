/**
	The Web Worker that opens documents (see Engine.js). Messages:

		page -> worker  {id, cmd: "open", name, buffer}   (buffer transferred)
		                {id, cmd: "edit", cell | para}       (see DocEngine.edit)
		                {id, cmd: "save", edits, format, sheet, text}
		worker -> page  {id, ok: true, result}            (picture bytes transferred)
		                {id, ok: false, error, code}

	The spreadsheet reader is loaded on first need (Engine's loadXLSX): it
	is the largest library and text documents never use it. The formula
	calculator likewise, on a workbook's first edit.
*/
/* global importScripts, self, DocEngine */
importScripts("../../lib/jszip/jszip.min.js", "Xml.js", "FileTypes.js", "Odf.js", "Pptx.js", "Text.js", "Sheets.js", "Docx.js",
	"SheetStream.js", "SheetPatch.js", "Formulas.js", "TextEdit.js", "Engine.js");

var engine = new DocEngine({JSZip: self.JSZip, loadXLSX: function() {
	importScripts("../../lib/sheetjs/xlsx.full.min.js");
	return self.XLSX;
}, loadHyperFormula: function() {
	importScripts("../../lib/hyperformula/hyperformula.full.min.js");
	return self.HyperFormula;
}});

self.onmessage = function(e) {
	var msg = e.data;
	var reply = function(inResult) {
		var transfer = [];
		var images = inResult && inResult.images;
		for (var k in images) {
			if (images.hasOwnProperty(k) && images[k].bytes) {
				transfer.push(images[k].bytes.buffer);
			}
		}
		if (inResult instanceof Uint8Array) {
			transfer.push(inResult.buffer);
		}
		self.postMessage({id: msg.id, ok: true, result: inResult}, transfer);
	};
	var fail = function(inError) {
		self.postMessage({id: msg.id, ok: false, error: String(inError && inError.message || inError),
			code: inError && inError.code || "", kind: inError && inError.kind || ""});
	};
	try {
		if (msg.cmd === "open") {
			engine.open({name: msg.name, buffer: msg.buffer, onProgress: function(inUpdate) {
				// Rows of a streamed workbook, after the first result.
				self.postMessage({id: msg.id, progress: inUpdate});
			}}).then(reply, fail);
		} else if (msg.cmd === "edit") {
			engine.edit(msg).then(reply, fail);
		} else if (msg.cmd === "save") {
			engine.save(msg).then(reply, fail);
		}
	} catch (err) {
		fail(err);
	}
};
