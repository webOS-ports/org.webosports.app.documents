/**
	Reads a file and has it opened by the engine (source/core/Engine.js),
	in a Web Worker when one starts, else in the page.

		DocLoader.read("/media/internal/a.docx", function(err, buffer) {...});
		DocLoader.open("a.docx", buffer, function(err, result) {...});
		DocLoader.edit({cell: {sheet, r, c, text}} | {para: {id, text}}, function(err, result) {...});
		DocLoader.save({edits, format}, function(err, bytes) {...});
		DocLoader.loadScript("lib/docx-preview/docx-preview.min.js", callback);
		DocLoader.loadStyle("lib/superdoc/style.css");

	Files are read with XMLHttpRequest on file: URLs, which WAM allows a
	"trusted" app (appinfo.json trustLevel) anywhere on the device.
*/
var DocLoader = {
	worker: null,
	workerFailed: false,
	pending: {},
	nextId: 1,
	scripts: {},
	engine: null,

	//* The app's own base URL, for the worker and the libraries.
	base: function() {
		return location.href.replace(/[?#].*$/, "").replace(/[^\/]*$/, "");
	},

	read: function(inPath, inCallback) {
		var url = /^[a-z]+:/i.test(inPath) ? inPath : DocLibrary.pathToUrl(inPath);
		var xhr = new XMLHttpRequest();
		var done = false;
		var finish = function(inErr, inBuffer) {
			if (!done) {
				done = true;
				inCallback(inErr, inBuffer);
			}
		};
		try {
			xhr.open("GET", url, true);
			xhr.responseType = "arraybuffer";
			xhr.onload = function() {
				// file: answers with status 0 on success in some engines.
				if ((xhr.status === 200 || xhr.status === 0) && xhr.response) {
					finish(null, xhr.response);
				} else {
					finish(new Error("The file could not be read (status " + xhr.status + ")."));
				}
			};
			xhr.onerror = function() {
				finish(new Error("The file could not be read. It may have been moved or deleted."));
			};
			xhr.send();
		} catch (e) {
			finish(new Error("The file could not be read: " + e.message));
		}
	},

	startWorker: function() {
		if (this.worker || this.workerFailed || window.__DOCS_NO_WORKER__) {
			return this.worker;
		}
		try {
			this.worker = new Worker(this.base() + "source/core/Worker.js");
			var self = this;
			this.worker.onmessage = function(e) {
				var msg = e.data;
				if (msg.progress) {
					self.deliverProgress(msg.id, msg.progress);
					return;
				}
				var cb = self.pending[msg.id];
				delete self.pending[msg.id];
				self.resultIn(msg.id);
				if (cb) {
					if (msg.ok) {
						cb(null, msg.result);
					} else {
						var err = new Error(msg.error);
						err.code = msg.code;
						err.kind = msg.kind;
						cb(err);
					}
				}
			};
			this.worker.onerror = function(e) {
				enyo.warn("document worker failed: " + (e && e.message));
				e.preventDefault && e.preventDefault();
				self.failWorker(new Error("The document reader stopped: " + (e && e.message || "unknown error")));
			};
		} catch (e) {
			enyo.warn("no document worker: " + e);
			this.workerFailed = true;
			this.worker = null;
		}
		return this.worker;
	},

	//* A worker that died: everything waiting fails, later calls run in the page.
	failWorker: function(inError) {
		var pending = this.pending;
		this.pending = {};
		if (this.worker) {
			this.worker.terminate();
		}
		this.worker = null;
		this.workerFailed = true;
		for (var k in pending) {
			if (pending.hasOwnProperty(k)) {
				pending[k](inError);
			}
		}
	},

	post: function(inMessage, inTransfer, inCallback, inOnProgress) {
		var id = this.nextId++;
		inMessage.id = id;
		this.pending[id] = inCallback;
		if (inOnProgress) {
			this.progress[id] = {fn: inOnProgress, ready: false, queue: []};
		}
		this.worker.postMessage(inMessage, inTransfer || []);
		return id;
	},

	progress: {},

	/**
		Progress (rows of a streamed workbook) may be sent before the result
		it belongs to has arrived: it waits until then, so the page always
		has the result first.
	*/
	deliverProgress: function(inId, inUpdate) {
		var p = this.progress[inId];
		if (!p) {
			return;
		}
		if (!p.ready) {
			p.queue.push(inUpdate);
			return;
		}
		p.fn(inUpdate);
		if (inUpdate.finished) {
			delete this.progress[inId];
		}
	},

	//* The result for inId is in: deliver what waited (after the result's own callback has run).
	resultIn: function(inId) {
		var self = this;
		var p = this.progress[inId];
		if (!p) {
			return;
		}
		setTimeout(function() {
			p.ready = true;
			var q = p.queue;
			p.queue = [];
			for (var i = 0; i < q.length; i++) {
				self.deliverProgress(inId, q[i]);
			}
		}, 0);
	},

	//* Stops delivering progress for every open request (a new document is being opened).
	dropProgress: function() {
		this.progress = {};
	},

	/**
		Opens a document. The buffer is handed to the worker (transferred),
		so pass a copy if the page still needs the bytes.
	*/
	open: function(inName, inBuffer, inCallback, inOnProgress) {
		this.dropProgress();
		if (this.startWorker()) {
			this.post({cmd: "open", name: inName, buffer: inBuffer}, [inBuffer], inCallback, inOnProgress);
			return;
		}
		var self = this;
		this.ensureLocalEngine(inName, function(inErr) {
			if (inErr) {
				inCallback(inErr);
				return;
			}
			var id = self.nextId++;
			if (inOnProgress) {
				self.progress[id] = {fn: inOnProgress, ready: false, queue: []};
			}
			self.engine.open({name: inName, buffer: inBuffer, onProgress: function(inUpdate) {
				self.deliverProgress(id, inUpdate);
			}}).then(function(r) {
				inCallback(null, r);
				self.resultIn(id);
			}, function(e) {
				inCallback(e);
			});
		});
	},

	//* An edit as it is made (see DocEngine.edit).
	edit: function(inRequest, inCallback) {
		if (this.worker) {
			this.post({cmd: "edit", cell: inRequest.cell, para: inRequest.para}, [], inCallback);
			return;
		}
		if (!this.engine) {
			inCallback(new Error("Nothing to edit."));
			return;
		}
		var self = this;
		var run = function() {
			self.engine.edit(inRequest).then(function(r) {
				inCallback(null, r);
			}, function(e) {
				inCallback(e);
			});
		};
		if (inRequest.cell) {
			// The calculator is loaded into the page on the first edit; without it formulas show as typed.
			this.loadScript("lib/hyperformula/hyperformula.full.min.js", run);
		} else {
			run();
		}
	},

	save: function(inRequest, inCallback) {
		if (this.worker) {
			this.post({cmd: "save", edits: inRequest.edits, format: inRequest.format, sheet: inRequest.sheet, text: inRequest.text},
				[], inCallback);
			return;
		}
		if (!this.engine) {
			inCallback(new Error("Nothing to save."));
			return;
		}
		var self = this;
		this.loadScript("lib/sheetjs/xlsx.full.min.js", function(inErr) {
			if (inErr) {
				inCallback(inErr);
				return;
			}
			self.engine.save(inRequest).then(function(b) {
				inCallback(null, b);
			}, function(e) {
				inCallback(e);
			});
		});
	},

	//* Without a worker: the libraries are loaded into the page on first use.
	ensureLocalEngine: function(inName, inCallback) {
		var self = this;
		var libs = ["lib/jszip/jszip.min.js"];
		var kind = FileTypes.kindFromName(inName);
		// Without a worker the page reads everything itself; nothing to stream into.
		if (FileTypes.kinds[kind] && FileTypes.kinds[kind].family === "sheet" || kind === "unknown") {
			libs.push("lib/sheetjs/xlsx.full.min.js");
		}
		this.loadScripts(libs, function(inErr) {
			if (!inErr && !self.engine) {
				self.engine = new DocEngine({JSZip: window.JSZip});
			}
			inCallback(inErr);
		});
	},

	loadScripts: function(inList, inCallback) {
		var self = this;
		var i = 0;
		var next = function(inErr) {
			if (inErr || i >= inList.length) {
				inCallback(inErr || null);
				return;
			}
			self.loadScript(inList[i++], next);
		};
		next();
	},

	//* Adds a style sheet once.
	loadStyle: function(inPath) {
		if (this.scripts[inPath]) {
			return;
		}
		this.scripts[inPath] = {loaded: true, waiting: []};
		var l = document.createElement("link");
		l.rel = "stylesheet";
		l.href = this.base() + inPath;
		document.head.appendChild(l);
	},

	//* Adds a script tag once; calls back when it has run.
	loadScript: function(inPath, inCallback) {
		var entry = this.scripts[inPath];
		if (entry && entry.loaded) {
			inCallback(null);
			return;
		}
		if (entry) {
			entry.waiting.push(inCallback);
			return;
		}
		entry = this.scripts[inPath] = {loaded: false, waiting: [inCallback]};
		var s = document.createElement("script");
		// The libraries hold non-ASCII characters (docx-preview's tab is U+2003).
		s.charset = "utf-8";
		s.src = this.base() + inPath;
		s.onload = function() {
			entry.loaded = true;
			var w = entry.waiting;
			entry.waiting = [];
			for (var i = 0; i < w.length; i++) {
				w[i](null);
			}
		};
		var self = this;
		s.onerror = function() {
			delete self.scripts[inPath];
			var w = entry.waiting;
			for (var i = 0; i < w.length; i++) {
				w[i](new Error("Could not load " + inPath));
			}
		};
		document.head.appendChild(s);
	}
};
