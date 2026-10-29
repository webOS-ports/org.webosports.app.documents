/**
	Finds the documents on the device: walks /media/internal through the
	app's own service (its list method, which gives names as arrays), or,
	in a build without it, the File Manager's service
	(org.webosports.service.filemanager/GetDirectories).

	The File Manager's service answers with the names joined by commas
	("dirs": "a,b"), so a name that itself has a comma arrives in pieces.
	splitNames() puts the pieces of document names back together, and the
	scanner checks a rejoined name exists before listing it. Names that
	come as an array are taken as they are.

		var scan = new DocLibrary.Scanner({
			list: function(dir, callback) {...},     // callback(err, {dirs: [..] or "..", files: [..] or ".."})
			exists: function(path, callback) {...},  // callback(true|false)
			onDone: function(docs) {...}             // [{path, name, dir, kind}]
		});
		scan.start("/media/internal");

	Pure logic with the I/O passed in; tested in node.
*/
(function(root) {
	var FT = root.FileTypes || (typeof require === "function" ? require("./FileTypes.js") : null);

	//* Directories never worth walking: hidden ones and the file system's own.
	var skipDirs = /^(\.|lost\+found$|LOST\.DIR$)/;

	/**
		Splits a comma-joined list of file names into candidate names. Pieces
		are joined back while the joined name ends in a document extension
		and the piece alone does not ("Report, final.docx" -> one name).
		Returns [{name, joined}] where joined says the name was rebuilt and
		should be checked. An array (the app's own service) is already whole
		names, none of them joined.
	*/
	function splitNames(inList) {
		if (!inList) {
			return [];
		}
		if (Array.isArray(inList)) {
			return inList.map(function(inName) {
				return {name: String(inName), joined: false};
			});
		}
		var parts = String(inList).split(",");
		var out = [];
		var pending = [];
		for (var i = 0; i < parts.length; i++) {
			pending.push(parts[i]);
			if (FT.isDocumentName(parts[i]) || i === parts.length - 1) {
				if (pending.length === 1) {
					out.push({name: parts[i], joined: false});
				} else {
					// Either the earlier pieces were names of their own (not documents),
					// or they belong to this name: offer both, the joined one first.
					out.push({name: pending.join(","), joined: true, alone: parts[i], before: pending.slice(0, -1)});
				}
				pending = [];
			}
		}
		return out;
	}

	function Scanner(inOptions) {
		this.list = inOptions.list;
		this.exists = inOptions.exists;
		this.onDone = inOptions.onDone;
		this.onProgress = inOptions.onProgress || function() {};
		this.maxDepth = inOptions.maxDepth || 8;
		this.maxDirs = inOptions.maxDirs || 2000;
		this.docs = [];
		this.queue = [];
		this.running = 0;
		this.dirsSeen = 0;
		this.cancelled = false;
	}

	Scanner.prototype.start = function(inRoot) {
		this.queue.push({path: inRoot.replace(/\/+$/, ""), depth: 0});
		this.pump();
	};

	Scanner.prototype.cancel = function() {
		this.cancelled = true;
	};

	// A few listings at a time: the service is a single node process.
	Scanner.prototype.pump = function() {
		if (this.cancelled) {
			return;
		}
		while (this.running < 3 && this.queue.length) {
			this.visit(this.queue.shift());
		}
		if (!this.running && !this.queue.length && !this.finished) {
			this.finished = true;
			this.docs.sort(function(a, b) {
				return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : a.name.toLowerCase() > b.name.toLowerCase() ? 1 : 0;
			});
			this.onDone(this.docs);
		}
	};

	Scanner.prototype.visit = function(inEntry) {
		var self = this;
		this.running++;
		this.dirsSeen++;
		this.list(inEntry.path, function(inErr, inResult) {
			if (self.cancelled) {
				return;
			}
			var checks = 0, listed = false, done = false;
			var finish = function() {
				if (listed && !done && checks === 0) {
					done = true;
					self.running--;
					self.onProgress(self.docs.length, self.dirsSeen);
					self.pump();
				}
			};
			if (!inErr && inResult) {
				if (inEntry.depth < self.maxDepth) {
					var dirs = Array.isArray(inResult.dirs) ? inResult.dirs : String(inResult.dirs || "").split(",");
					for (var d = 0; d < dirs.length; d++) {
						if (dirs[d] && !skipDirs.test(dirs[d]) && self.dirsSeen + self.queue.length < self.maxDirs) {
							self.queue.push({path: inEntry.path + "/" + dirs[d], depth: inEntry.depth + 1});
						}
					}
				}
				var names = splitNames(inResult.files);
				names.forEach(function(inCandidate) {
					if (!FT.isDocumentName(inCandidate.name)) {
						return;
					}
					if (!inCandidate.joined) {
						self.add(inEntry.path, inCandidate.name);
						return;
					}
					checks++;
					var full = inEntry.path + "/" + inCandidate.name;
					self.exists(full, function(inYes) {
						if (inYes) {
							self.add(inEntry.path, inCandidate.name);
						} else if (FT.isDocumentName(inCandidate.alone)) {
							self.add(inEntry.path, inCandidate.alone);
						}
						checks--;
						finish();
					});
				});
			}
			listed = true;
			finish();
		});
	};

	Scanner.prototype.add = function(inDir, inName) {
		this.docs.push({path: inDir + "/" + inName, name: inName, dir: inDir, kind: FT.kindFromName(inName)});
	};

	// --- Recent documents ------------------------------------------------

	/**
		The recently opened list, newest first, at most inMax entries:
		add(list, entry) returns the new list with entry moved to the top.
	*/
	function addRecent(inList, inEntry, inMax) {
		var max = inMax || 10;
		var out = [inEntry];
		for (var i = 0; i < (inList || []).length && out.length < max; i++) {
			if (inList[i] && inList[i].path !== inEntry.path) {
				out.push(inList[i]);
			}
		}
		return out;
	}

	//* A launch "target" (file URL or path) as a plain path; null when it is no local file.
	function targetToPath(inTarget) {
		var t = String(inTarget || "").trim();
		if (!t) {
			return null;
		}
		if (/^file:\/\//i.test(t)) {
			t = t.replace(/^file:\/\/(localhost)?/i, "");
			try {
				t = decodeURIComponent(t);
			} catch (e) {
				return null;
			}
		} else if (/^[a-z][a-z0-9+.-]*:/i.test(t)) {
			return null;
		}
		if (t.charAt(0) !== "/") {
			return null;
		}
		// No climbing out with "..": normalise the path.
		var parts = t.split("/");
		var out = [];
		for (var i = 0; i < parts.length; i++) {
			if (parts[i] === "..") {
				out.pop();
			} else if (parts[i] && parts[i] !== ".") {
				out.push(parts[i]);
			}
		}
		return "/" + out.join("/");
	}

	//* A local path as a file: URL (each segment percent-encoded).
	function pathToUrl(inPath) {
		return "file://" + String(inPath).split("/").map(function(s) {
			return encodeURIComponent(s);
		}).join("/");
	}

	function baseName(inPath) {
		var s = String(inPath || "");
		return s.slice(s.lastIndexOf("/") + 1);
	}

	function dirName(inPath) {
		var s = String(inPath || "");
		var i = s.lastIndexOf("/");
		return i <= 0 ? "/" : s.slice(0, i);
	}

	/**
		A name for a saved copy that does not clash with inExisting (names in
		the same folder): "Budget (edited).xlsx", then "(edited 2)", ...
	*/
	function copyName(inName, inExt, inExisting) {
		var dot = inName.lastIndexOf(".");
		var stem = dot > 0 ? inName.slice(0, dot) : inName;
		stem = stem.replace(/ \(edited( \d+)?\)$/, "");
		var taken = {};
		(inExisting || []).forEach(function(n) {
			taken[n.toLowerCase()] = true;
		});
		for (var i = 1; i < 1000; i++) {
			var candidate = stem + (i === 1 ? " (edited)" : " (edited " + i + ")") + "." + inExt;
			if (!taken[candidate.toLowerCase()]) {
				return candidate;
			}
		}
		return stem + " (edited " + Date.now() + ")." + inExt;
	}

	function formatSize(inBytes) {
		var b = +inBytes;
		if (!(b >= 0)) {
			return "";
		}
		if (b < 1024) {
			return b + " B";
		}
		if (b < 1024 * 1024) {
			return (Math.round(b / 102.4) / 10) + " KB";
		}
		return (Math.round(b / 104857.6) / 10) + " MB";
	}

	var DocLibrary = {
		splitNames: splitNames,
		Scanner: Scanner,
		addRecent: addRecent,
		targetToPath: targetToPath,
		pathToUrl: pathToUrl,
		baseName: baseName,
		dirName: dirName,
		copyName: copyName,
		formatSize: formatSize
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = DocLibrary;
	} else {
		root.DocLibrary = DocLibrary;
	}
})(this);
