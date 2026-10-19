/**
	Documents in cloud accounts: the accounts set up in Accounts with the
	DOCUMENTS capability (Dropbox, Box, Google Drive, OneDrive, ... from
	luneos-synergy-connectors), and their folders through each account's
	connector service.

		CloudService.accounts(function(err, [{id, provider, name, service}]) {...});
		CloudService.list(account, folderPath, function(err, [{id, type, name, size, modified, path, mimeType}]) {...});
		CloudService.download(account, entry, function(err, localPath) {...});
		CloudService.upload(account, folderPath, localPath, name, function(err, result) {...});
		CloudService.describeError(err)   // a short sentence for the user

	The account list is the Accounts service's
	(com.palm.service.accounts/listAccounts, group accountsservice.operation).
	Each account names its connector in its DOCUMENTS capability provider
	("implementation": "palm://com.palm.service.dropbox/"); the connector
	lets this app call it (its perm.json grants the group). A folder's
	"path" is the connector's own locator, a path at some providers and an
	id at others: it is only ever handed back, never taken apart.

	Downloads go to a hidden folder, which the library does not list:
	/media/internal/.documents/cloud/<account>/<entry>/<the file's own name>.
	A folder per entry keeps two files of the same name apart while the
	file keeps its name and extension (FileTypes goes by them).

	The bus calls go through FileService.call (CloudService.call, which the
	unit tests replace); the rest is pure logic, tested in node.
*/
(function(root) {
	var FT = root.FileTypes || (typeof require === "function" ? require("../core/FileTypes.js") : null);

	var cacheRoot = "/media/internal/.documents/cloud";

	//* "palm://com.palm.service.dropbox/" -> "com.palm.service.dropbox"; null for anything else.
	function serviceOf(inImplementation) {
		var m = /^palm:\/\/([^\/]+)\/?$/.exec(String(inImplementation || ""));
		return m && /^[A-Za-z0-9._-]+$/.test(m[1]) ? m[1] : null;
	}

	/**
		The accounts the app can browse, from listAccounts' results: those
		whose DOCUMENTS provider names a connector service. Sorted by
		provider, then by account name.
	*/
	function usableAccounts(inResults) {
		var out = [];
		(inResults || []).forEach(function(a) {
			if (!a || !a._id) {
				return;
			}
			var service = null;
			(a.capabilityProviders || []).forEach(function(p) {
				if (!service && p && p.capability === "DOCUMENTS") {
					service = serviceOf(p.implementation);
				}
			});
			if (service) {
				// An alias can come with spaces around it ("Gerardus "). icon: the account template's
				// picture of the provider, as the Accounts app shows it, or "".
				var icons = a.icon || {};
				out.push({id: String(a._id), provider: trim(a.loc_name) || String(a.templateId || service),
					name: trim(a.alias) || trim(a.username), service: service,
					icon: trim(icons.loc_48x48 || icons.loc_32x32)});
			}
		});
		out.sort(function(x, y) {
			var a = (x.provider + "\u0000" + x.name).toLowerCase(), b = (y.provider + "\u0000" + y.name).toLowerCase();
			return a < b ? -1 : a > b ? 1 : 0;
		});
		return out;
	}

	//* A folder's entries as they are shown: folders first, then files, each by name.
	function sortEntries(inEntries) {
		return (inEntries || []).filter(function(e) {
			return e && typeof e.name === "string" && e.name && (e.type === "folder" || e.type === "file");
		}).sort(function(x, y) {
			if (x.type !== y.type) {
				return x.type === "folder" ? -1 : 1;
			}
			var a = x.name.toLowerCase(), b = y.name.toLowerCase();
			return a < b ? -1 : a > b ? 1 : 0;
		});
	}

	//* Folders, and the files the library would list (ones the app opens, converts or hands on).
	function visibleEntries(inEntries) {
		return sortEntries(inEntries).filter(function(e) {
			return e.type === "folder" || FT.isDocumentName(e.name);
		});
	}

	//* Whether a folder's entries have a file or folder called inName (as most clouds do, ignoring case).
	function hasName(inEntries, inName) {
		var n = String(inName || "").toLowerCase();
		return (inEntries || []).some(function(e) {
			return e && String(e.name || "").toLowerCase() === n;
		});
	}

	//* FNV-1a of a string, as 8 hex digits: a short folder name per entry.
	function hash(inText) {
		var s = String(inText);
		var h = 0x811c9dc5;
		for (var i = 0; i < s.length; i++) {
			h ^= s.charCodeAt(i);
			h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
		}
		return ("0000000" + h.toString(16)).slice(-8);
	}

	//* An account id as one folder name: letters, digits, "-", "_" and "." only, never hidden.
	function folderName(inText) {
		var s = String(inText || "").replace(/[^A-Za-z0-9._-]/g, "_").replace(/^\.+/, "_").slice(0, 64);
		return s || "_";
	}

	//* A cloud file's name as a local one: no slashes or control characters, not hidden, its extension kept.
	function fileName(inName) {
		var s = String(inName || "").replace(/[\/\\\u0000-\u001f\u007f]/g, "_").replace(/^[\s.]+/, "");
		return s || "document";
	}

	//* Where entry is downloaded to: a folder per account, then one per entry, then the file's own name.
	function trim(inText) {
		return String(inText || "").replace(/^\s+|\s+$/g, "");
	}

	function cachePath(inAccount, inEntry) {
		return cacheRoot + "/" + folderName(inAccount.id) + "/" + hash((inEntry.id || "") + "\n" + (inEntry.path || "")) +
			"/" + fileName(inEntry.name);
	}

	/**
		Where a saved copy of a downloaded document is written before it is
		uploaded: next to the download, under a name no earlier copy has
		(the copy is uploaded under inName).
	*/
	function copyPath(inCachedPath, inName, inNow) {
		var s = String(inCachedPath);
		return s.slice(0, s.lastIndexOf("/")) + "/copy-" + (inNow || Date.now()).toString(36) + "-" + fileName(inName);
	}

	//* What a document's origin (kept with a recently opened one) must look like to be used again.
	function isOrigin(inOrigin) {
		var a = inOrigin && inOrigin.account;
		return !!(a && typeof a.id === "string" && typeof a.service === "string" && /^[A-Za-z0-9._-]+$/.test(a.service) &&
			(inOrigin.folderPath === null || inOrigin.folderPath === undefined || typeof inOrigin.folderPath === "string"));
	}

	/**
		A failed call as a short sentence for the user. inErr is what
		FileService.call passes: an Error with the answer's errorCode as
		code and its errorText as message.
	*/
	function describeError(inErr) {
		var code = String(inErr && inErr.code !== undefined && inErr.code !== null ? inErr.code : "");
		var text = String(inErr && inErr.message || "");
		if (code === "PERMISSION_DENIED") {
			return $L("This account's service does not let Documents use it.");
		}
		if (/^(NO_CREDENTIALS|NO_STORED_CREDENTIALS|OAUTH_TOKEN_FAILED|401)/.test(code) || /\b401\b|unauthori[sz]ed/i.test(text)) {
			return $L("Sign in to this account again in Accounts.");
		}
		if (code === "NOT_SUPPORTED") {
			return $L("This account cannot do that.");
		}
		if (/^Denied method call|does not exist|unknown service|no luna bus/i.test(text)) {
			return $L("The service for this account is not installed.");
		}
		if (/^(CURL_FAILED|NO_CHILD_PROCESS|ETIMEDOUT|ENOTFOUND|ECONNREFUSED)/.test(code) || /timed? ?out|network|resolve host/i.test(text)) {
			return $L("The cloud could not be reached. Check the connection and try again.");
		}
		return $L("The cloud service reported a problem.") + (code ? " (" + code + ")" : "");
	}

	var CloudService = {
		cacheRoot: cacheRoot,
		serviceOf: serviceOf,
		usableAccounts: usableAccounts,
		sortEntries: sortEntries,
		visibleEntries: visibleEntries,
		hasName: hasName,
		hash: hash,
		fileName: fileName,
		cachePath: cachePath,
		copyPath: copyPath,
		isOrigin: isOrigin,
		describeError: describeError,

		//* The bus: FileService.call on the device, a fake in the unit tests.
		call: function(inUri, inParams, inCallback) {
			return root.FileService.call(inUri, inParams, inCallback);
		},

		uri: function(inAccount, inMethod) {
			return "luna://" + inAccount.service + "/" + inMethod;
		},

		accounts: function(inCallback) {
			this.call("luna://com.palm.service.accounts/listAccounts", {capability: "DOCUMENTS"}, function(inErr, inResult) {
				inCallback(inErr, inErr ? null : usableAccounts(inResult && inResult.results));
			});
		},

		//* The entries of a folder (inFolderPath an entry's path; empty for the account's top folder).
		list: function(inAccount, inFolderPath, inCallback) {
			var params = {accountId: inAccount.id};
			if (inFolderPath) {
				params.path = inFolderPath;
			}
			this.call(this.uri(inAccount, "listFolder"), params, function(inErr, inResult) {
				inCallback(inErr, inErr ? null : sortEntries(inResult && inResult.entries));
			});
		},

		download: function(inAccount, inEntry, inCallback) {
			var localPath = cachePath(inAccount, inEntry);
			this.call(this.uri(inAccount, "downloadFile"), {accountId: inAccount.id, path: inEntry.path, localPath: localPath},
				function(inErr) {
					inCallback(inErr, inErr ? null : localPath);
				});
		},

		//* Uploads a local file as a new file called inName into a folder (empty inFolderPath: the top one).
		upload: function(inAccount, inFolderPath, inLocalPath, inName, inCallback) {
			var params = {accountId: inAccount.id, localPath: inLocalPath, name: inName};
			if (inFolderPath) {
				params.path = inFolderPath;
			}
			var kind = FT.kinds[FT.kindFromName(inName)];
			if (kind && kind.mime) {
				params.mimeType = kind.mime;
			}
			this.call(this.uri(inAccount, "uploadFile"), params, inCallback);
		}
	};

	if (typeof module !== "undefined" && module.exports) {
		module.exports = CloudService;
	} else {
		root.CloudService = CloudService;
	}
})(this);
