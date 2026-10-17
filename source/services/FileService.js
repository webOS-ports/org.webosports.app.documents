/**
	The luna services the app uses, behind plain callbacks:

		FileService.list("/media/internal", function(err, {dirs, files}) {...});
		FileService.size(path, function(err, bytes) {...});
		FileService.openUrl("https://...");            // in the browser (applicationManager/open)
		FileService.openWith(path, "application/pdf", function(err) {...}); // in the MIME type's app
		FileService.writeNew(path, bytes, function(err) {...});
		FileService.convert(path, "odt", function(err, {target}) {...});   // LibreOffice, when installed

	list and size are the File Manager's service
	(org.webosports.service.filemanager, group filemanager-service.operation).
	writeNew is this app's own service (services/org.webosports.app.documents.service),
	which creates a file and refuses to replace one that exists.
*/
var FileService = {
	call: function(inUri, inParams, inCallback) {
		if (typeof PalmServiceBridge === "undefined") {
			inCallback(new Error("No luna bus (not running on LuneOS)"));
			return null;
		}
		var bridge = new PalmServiceBridge();
		var done = false;
		bridge.onservicecallback = function(inText) {
			if (done) {
				return;
			}
			done = true;
			var r;
			try {
				r = JSON.parse(inText);
			} catch (e) {
				inCallback(new Error("Bad answer from " + inUri));
				return;
			}
			if (r.returnValue === false) {
				var err = new Error(r.errorText || ("Failed: " + inUri));
				err.code = r.errorCode;
				inCallback(err, r);
			} else {
				inCallback(null, r);
			}
			// Let the bridge go once answered.
			setTimeout(function() {
				try {
					bridge.cancel();
				} catch (e) {}
			}, 0);
		};
		bridge.call(inUri, JSON.stringify(inParams || {}));
		return bridge;
	},

	//* This app's id ("org.webosports.app.documents"), from WAM.
	appId: function() {
		var id = window.PalmSystem && PalmSystem.identifier ? String(PalmSystem.identifier).split(" ")[0] : "";
		return id || "org.webosports.app.documents";
	},

	list: function(inDir, inCallback) {
		this.call("luna://org.webosports.service.filemanager/GetDirectories", {dir: inDir}, function(inErr, inResult) {
			inCallback(inErr, inErr ? null : {dirs: inResult.dirs || "", files: inResult.files || ""});
		});
	},

	size: function(inPath, inCallback) {
		this.call("luna://org.webosports.service.filemanager/getFileSize", {path: inPath}, function(inErr, inResult) {
			inCallback(inErr, inErr ? -1 : inResult.size);
		});
	},

	exists: function(inPath, inCallback) {
		this.size(inPath, function(inErr) {
			inCallback(!inErr);
		});
	},

	openUrl: function(inUrl, inCallback) {
		this.call("luna://com.palm.applicationManager/open", {target: inUrl}, inCallback || function() {});
	},

	/**
		Opens a local file in the app registered for inMime: the application
		manager launches that app with {"target": "file://..."}. Fails when
		no installed app handles the type.
	*/
	openWith: function(inPath, inMime, inCallback) {
		var url = "file://" + encodeURI(inPath).replace(/[?#]/g, encodeURIComponent);
		var done = inCallback || function() {};
		this.call("luna://com.palm.applicationManager/open", {target: url, mime: inMime}, function(inErr, inResult) {
			// luna-appmanager without the fix for its launch replies passes the
			// launch on to SAM without waiting, and answers {"returnValue": false,
			// "errorText": ""} for an app that did start; with it, the answer is
			// SAM's. Either way "No handler for ..." is what a real failure says.
			done(inErr && !(inResult && inResult.returnValue === false && inResult.errorText === "") ? inErr : null);
		});
	},

	/**
		Whether this app's service (which writes saved copies) is installed:
		calls back true or false, asking the bus once per run.
	*/
	serviceAvailable: function(inCallback) {
		var self = this;
		if (this.serviceState !== undefined) {
			inCallback(this.serviceState);
			return;
		}
		(this.serviceWaiting = this.serviceWaiting || []).push(inCallback);
		if (this.serviceWaiting.length > 1) {
			return;
		}
		this.call("luna://" + this.appId() + ".service/info", {}, function(inErr) {
			self.serviceState = !inErr;
			var w = self.serviceWaiting;
			self.serviceWaiting = [];
			for (var i = 0; i < w.length; i++) {
				w[i](self.serviceState);
			}
		});
	},

	/**
		Writes inBytes (Uint8Array) to a new file at inPath. Fails, writing
		nothing, when the file exists.
	*/
	/**
		Has LibreOffice (org.webosports.service.office) convert a document
		into inFormat ("odt", "ods", "odp", ...). The result is a file in
		the service's cache: {target, format, cached}.
	*/
	convert: function(inPath, inFormat, inCallback) {
		this.call("luna://org.webosports.service.office/convert", {source: inPath, format: inFormat}, inCallback);
	},

	writeNew: function(inPath, inBytes, inCallback) {
		var bin = "";
		var chunk = 0x8000;
		for (var i = 0; i < inBytes.length; i += chunk) {
			bin += String.fromCharCode.apply(null, inBytes.subarray(i, i + chunk));
		}
		this.call("luna://" + this.appId() + ".service/writeNew", {path: inPath, data: btoa(bin)}, function(inErr, inResult) {
			inCallback(inErr, inResult);
		});
	}
};
