/*
	Documents' service, on the bus as org.webosports.app.documents.service:

		writeNew {"path": "/media/internal/.../Budget (edited).xlsx", "data": "<base64>"}
			-> {"returnValue": true, "path", "bytes"}
			-> {"returnValue": false, "errorCode": "EEXIST" | "EINVAL" | ..., "errorText"}
		list {"dir": "/media/internal/..."}
			-> {"returnValue": true, "dirs": [names], "files": [names]}
		size {"path": "/media/internal/..."}
			-> {"returnValue": true, "size": <bytes>}
			   the library's scan (files.js), only below /media/internal
		quit, info   for the installer (it asks a service to quit before an update)

	It talks to the bus through palmbus directly rather than through
	webos-service, which first creates an activity for every call: a call
	(the installer's "quit" too) would then wait for the activity manager,
	and never be answered when that is refused. The service exits after a
	short idle time; the bus starts it again on the next call.
*/
var palmbus = require("palmbus");
var path = require("path");
var writer = require("./writer.js");
var files = require("./files.js");

// The directory is named after the service (a test build may install it under another id).
var ID = path.basename(__dirname);
var ROOT = "/media/internal/";
var IDLE_MS = 30000;

process.title = ID;

var handle = new palmbus.Handle(ID);
var METHODS = ["writeNew", "list", "size", "quit", "info"];

METHODS.forEach(function(m) {
	handle.registerMethod("/", m);
});

var idle = null;
function touch() {
	if (idle) {
		clearTimeout(idle);
	}
	idle = setTimeout(function() {
		process.exit(0);
	}, IDLE_MS);
}

handle.addListener("request", function(inMessage) {
	touch();
	var payload = {};
	try {
		payload = JSON.parse(inMessage.payload() || "{}");
	} catch (e) {
		inMessage.respond(JSON.stringify({returnValue: false, errorCode: "EINVAL", errorText: "bad JSON"}));
		return;
	}
	switch (inMessage.method()) {
	case "writeNew":
		inMessage.respond(JSON.stringify(writer.writeNew(ROOT, payload.path, payload.data)));
		break;
	case "list":
		inMessage.respond(JSON.stringify(files.list(ROOT, payload.dir)));
		break;
	case "size":
		inMessage.respond(JSON.stringify(files.size(ROOT, payload.path)));
		break;
	case "info":
		inMessage.respond(JSON.stringify({returnValue: true, id: ID, methods: METHODS}));
		break;
	case "quit":
		inMessage.respond(JSON.stringify({returnValue: true, status: "quitting"}));
		setTimeout(function() {
			process.exit(0);
		}, 200);
		break;
	default:
		inMessage.respond(JSON.stringify({returnValue: false, errorCode: "ENOSYS", errorText: "Unknown method"}));
	}
});

touch();
