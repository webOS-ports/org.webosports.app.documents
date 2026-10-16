/*
	In-page mock of the webOS runtime and the network, for automated tests.

	The test runner (tools/tests/lib/browser.js) injects this before any page
	script runs, preceded by `window.__MOCK_CONFIG__ = {...}`. The app is
	loaded unchanged (debug.html); nothing in the product knows about it.

	What it fakes
	  - window.PalmSystem and window.PalmServiceBridge, so enyo.PalmService
	    (webos-lib) calls are answered from config.luna.
	  - XMLHttpRequest for absolute http(s) URLs, answered from config.http.
	    Relative URLs (the app's own files) go to the real XHR. A request that
	    matches no rule fails like a network error and is logged as unmocked,
	    so a test can assert that nothing escaped to the internet.
	  - localStorage, cleared or seeded per config.storage.
	  - window.open: with config.webos the call is logged as
	    {kind: "window", url, name, attributes} and a fake window is
	    returned (closing it is logged too); this is how an app asks WAM for
	    a dashboard or popup alert.

	Config (all optional):
	  webos: true            install PalmSystem/PalmServiceBridge (default true)
	  palmSystem: {...}      overrides for PalmSystem fields (locale, timeFormat, ...)
	  luna: {"luna://svc/method": rule}
	  http: [rule, ...]      first rule whose `match` (RegExp source) matches the URL wins
	  storage: {clear: bool, set: {key: value}}
	  acg: {groups: {name: ["service/method", ...]}, granted: [name, ...], exempt: [RegExp source]}
	                         when given, a luna call is denied the way ls-hubd does unless one
	                         of its groups is granted (the app's requiredPermissions); groups
	                         come from the device's api-permissions.d files, see
	                         tools/tests/fixtures/acg/. A method no group lists is denied and
	                         logged as acg "unknown", so a missing table shows up too.

	A rule answers with one of
	  {json: {...}}                    success
	  {error: {errorCode, errorText}}  luna failure
	  {status: 500, json: {...}}       HTTP failure (status 0 = offline)
	  {responses: [rule, ...]}         subscription: each entry is sent in turn
	  plus {delay: ms} to answer late, or {never: true} to never answer.

	window.__MOCK__ exposes
	  calls     [{kind: "http"|"luna", url, params, unmocked}] in call order
	  relaunch(params), activate(), deactivate(): fire the webOS lifecycle
	  set(path, rule): change a luna/http rule while the page runs
*/
(function() {
	var config = window.__MOCK_CONFIG__ || {};
	var calls = [];
	var mock = window.__MOCK__ = {calls: calls, config: config};

	function clone(inValue) {
		return inValue === undefined ? undefined : JSON.parse(JSON.stringify(inValue));
	}

	function later(inRule, inFn) {
		if (inRule && inRule.never) {
			return;
		}
		setTimeout(inFn, (inRule && inRule.delay) || 0);
	}

	// --- localStorage ----------------------------------------------------

	// Once per tab, not on every load: a test that reloads the page is
	// checking that storage survives.
	try {
		var storage = config.storage || {};
		var first = !window.sessionStorage.getItem("__webosMockStorageDone");
		window.sessionStorage.setItem("__webosMockStorageDone", "1");
		if (first && storage.clear) {
			window.localStorage.clear();
		}
		for (var key in (first && storage.set) || {}) {
			if (storage.set.hasOwnProperty(key)) {
				window.localStorage.setItem(key, typeof storage.set[key] === "string" ? storage.set[key] : JSON.stringify(storage.set[key]));
			}
		}
	} catch (e) {
		console.error("webos-mock: storage setup failed: " + e);
	}

	// --- webOS runtime ---------------------------------------------------

	if (config.webos !== false) {
		var palm = {
			identifier: "org.webosports.app.weather 0",
			launchParams: "{}",
			deviceInfo: JSON.stringify({modelName: "Mock", platformVersion: "LuneOS mock", screenWidth: window.innerWidth,
				screenHeight: window.innerHeight}),
			locale: "en_us",
			localeRegion: "us",
			phoneRegion: "us",
			timeFormat: "HH24",
			isActivated: true,
			screenOrientation: "up",
			stageReady: function() { calls.push({kind: "palm", url: "stageReady"}); },
			activate: function() {},
			deactivate: function() {},
			setWindowProperties: function() {},
			setWindowOrientation: function() {},
			enableFullScreenMode: function() {},
			addBannerMessage: function(inText) { calls.push({kind: "palm", url: "addBannerMessage", params: inText}); return 1; },
			removeBannerMessage: function() {},
			addNewContentIndicator: function() { return 1; },
			removeNewContentIndicator: function() {},
			paste: function() {},
			// WAM synthesises a click at a point; Enyo 1 uses it to focus an input under a finger.
			simulateMouseClick: function(inX, inY, inDown) {
				if (!inDown) {
					return;
				}
				var el = document.elementFromPoint(inX, inY);
				if (el && el.focus && /^(INPUT|TEXTAREA)$/.test(el.nodeName)) {
					el.focus();
				}
			}
		};
		var overrides = config.palmSystem || {};
		for (var p in overrides) {
			if (overrides.hasOwnProperty(p)) {
				palm[p] = overrides[p];
			}
		}
		window.PalmSystem = palm;

		// WAM's synchronous resource read, which Enyo 1 uses for appinfo.json and
		// framework_config.json: "const json" parses the file, otherwise text.
		window.palmGetResource = function(inUrl, inMode) {
			var x = new RealXHR();
			x.open("GET", inUrl, false);
			x.send();
			if (x.status !== 200 && x.status !== 0) {
				throw new Error("palmGetResource: " + inUrl + " (" + x.status + ")");
			}
			return /json/.test(inMode || "") ? JSON.parse(x.responseText) : x.responseText;
		};

		// "luna://svc/method/" -> "svc/method"; matched against api-permissions entries.
		var checkAcg = function(inUrl) {
			var acg = config.acg;
			if (!acg) {
				return {allowed: true, result: "unchecked"};
			}
			var key = inUrl.replace(/^luna:\/\//, "").replace(/\/+$/, "");
			var exempt = (acg.exempt || []).some(function(inRe) { return new RegExp(inRe).test(inUrl); });
			if (exempt) {
				return {allowed: true, result: "exempt"};
			}
			var groups = [];
			for (var g in acg.groups) {
				if (acg.groups.hasOwnProperty(g) && acg.groups[g].some(function(inEntry) {
					return inEntry === key || (inEntry.slice(-1) === "*" && key.indexOf(inEntry.slice(0, -1)) === 0);
				})) {
					groups.push(g);
				}
			}
			if (!groups.length) {
				return {allowed: false, result: "unknown"};
			}
			var ok = groups.some(function(inGroup) { return (acg.granted || []).indexOf(inGroup) >= 0; });
			return {allowed: ok, result: ok ? "granted" : "denied: needs one of " + groups.join(", ")};
		};

		window.PalmServiceBridge = function() {
			this.cancelled = false;
		};
		window.PalmServiceBridge.prototype.call = function(inUrl, inParams) {
			var self = this;
			var params = {};
			try {
				params = JSON.parse(inParams || "{}");
			} catch (e) {}
			var url = inUrl.replace(/^palm:\/\//, "luna://");
			// Accept both "luna://svc/method" and "luna://svc/method/" spellings.
			var rule = (config.luna || {})[url] || (config.luna || {})[url.replace(/\/+$/, "")];
			// A rule with cases answers by the call's parameters: the first case whose
			// "when" fields all equal the parameters wins, else the rule's own answer.
			if (rule && rule.cases) {
				var picked = null;
				for (var ci = 0; ci < rule.cases.length && !picked; ci++) {
					var when = rule.cases[ci].when || {};
					var all = true;
					for (var wk in when) {
						if (when.hasOwnProperty(wk) && params[wk] !== when[wk]) {
							all = false;
						}
					}
					if (all) {
						picked = rule.cases[ci];
					}
				}
				rule = picked || (rule.json || rule.error ? rule : null);
			}
			var acg = checkAcg(url);
			calls.push({kind: "luna", url: url, params: params, unmocked: !rule && acg.allowed, acg: acg.result});
			if (!acg.allowed) {
				// What ls-hubd answers for a method the app has no permission group for.
				rule = {error: {errorCode: -1, errorText: "Denied method call \"" + url.replace(/^.*\//, "") + "\" for category \"/\""}};
			}
			var send = function(inRule) {
				later(inRule, function() {
					if (self.cancelled || !self.onservicecallback) {
						return;
					}
					var body = inRule.error ? Object.assign({returnValue: false}, inRule.error) :
						Object.assign({returnValue: true}, clone(inRule.json) || {});
					self.onservicecallback(JSON.stringify(body));
				});
			};
			if (!rule) {
				send({error: {errorCode: -1, errorText: "Unknown service (mock has no rule for " + url + ")"}});
			} else if (rule.responses) {
				var at = 0;
				rule.responses.forEach(function(inRule) {
					at += inRule.delay || 0;
					send(Object.assign({}, inRule, {delay: at}));
				});
			} else {
				send(rule);
			}
		};
		window.PalmServiceBridge.prototype.cancel = function() {
			this.cancelled = true;
		};

		// window.open with attributes is WAM's way to open a dashboard or a
		// popup alert; record what was asked for instead of opening a popup.
		window.open = function(inUrl, inName, inFeatures) {
			var attributes = null;
			var m = /attributes=(\{.*\})/.exec(inFeatures || "");
			if (m) {
				try {
					attributes = JSON.parse(m[1]);
				} catch (e) {
					attributes = "unparseable: " + m[1];
				}
			}
			var entry = {kind: "window", url: String(inUrl), name: inName, attributes: attributes};
			calls.push(entry);
			var fake = {
				closed: false,
				close: function() {
					if (!fake.closed) {
						fake.closed = true;
						calls.push({kind: "window", url: "close:" + entry.url, name: inName});
					}
				}
			};
			return fake;
		};

		// The lifecycle hooks webos-lib's LunaBindings installs on Mojo.
		mock.relaunch = function(inParams) {
			window.PalmSystem.launchParams = JSON.stringify(inParams || {});
			window.Mojo && window.Mojo.relaunch && window.Mojo.relaunch();
		};
		mock.activate = function() {
			window.Mojo && window.Mojo.stageActivated && window.Mojo.stageActivated();
		};
		mock.deactivate = function() {
			window.Mojo && window.Mojo.stageDeactivated && window.Mojo.stageDeactivated();
		};
	}

	// --- Network ---------------------------------------------------------

	var RealXHR = window.XMLHttpRequest;

	function findHttpRule(inUrl) {
		var rules = config.http || [];
		for (var i = 0; i < rules.length; i++) {
			if (new RegExp(rules[i].match).test(inUrl)) {
				return rules[i];
			}
		}
		return null;
	}

	function MockXHR() {
		this.readyState = 0;
		this.status = 0;
		this.responseText = "";
		this.listeners = {};
		this.aborted = false;
	}
	MockXHR.prototype = {
		open: function(inMethod, inUrl) {
			this.method = inMethod;
			this.url = inUrl;
			this.readyState = 1;
		},
		setRequestHeader: function() {},
		getResponseHeader: function(inName) {
			return /content-type/i.test(inName) ? "application/json" : null;
		},
		getAllResponseHeaders: function() {
			return "content-type: application/json\r\n";
		},
		addEventListener: function(inType, inFn) {
			(this.listeners[inType] = this.listeners[inType] || []).push(inFn);
		},
		removeEventListener: function() {},
		abort: function() {
			this.aborted = true;
		},
		fire: function(inType) {
			var fn = this["on" + inType];
			if (fn) {
				fn.call(this, {type: inType, target: this});
			}
			(this.listeners[inType] || []).forEach(function(inFn) {
				inFn.call(this, {type: inType, target: this});
			}, this);
		},
		send: function() {
			var self = this;
			var rule = findHttpRule(this.url);
			calls.push({kind: "http", url: this.url, unmocked: !rule});
			if (!rule) {
				console.warn("webos-mock: unmocked request " + this.url);
				rule = {status: 0};
			}
			later(rule, function() {
				if (self.aborted) {
					return;
				}
				self.readyState = 4;
				self.status = rule.status === undefined ? 200 : rule.status;
				self.responseText = rule.text !== undefined ? rule.text :
					rule.json !== undefined ? JSON.stringify(rule.json) : "";
				self.fire("readystatechange");
				self.fire(self.status ? "load" : "error");
				self.fire("loadend");
			});
		}
	};

	window.XMLHttpRequest = function() {
		var mockXhr = new MockXHR();
		var realXhr = null;
		// Decide on open(): the app's own files still load for real.
		var proxy = {
			open: function(inMethod, inUrl) {
				if (/^https?:\/\//i.test(inUrl) && inUrl.indexOf(location.origin) !== 0) {
					mockXhr.open(inMethod, inUrl);
					proxy.target = mockXhr;
				} else {
					// config.rewrite: [{from: RegExp source, to: "replacement"}], e.g. device file:
					// URLs onto the test server, so the app's own reads work unchanged.
					var rewritten = inUrl;
					(config.rewrite || []).forEach(function(inRule) {
						rewritten = rewritten.replace(new RegExp(inRule.from), inRule.to);
					});
					if (rewritten !== inUrl) {
						calls.push({kind: "file", url: inUrl, to: rewritten});
						arguments[1] = rewritten;
					}
					realXhr = new RealXHR();
					proxy.target = realXhr;
					realXhr.onreadystatechange = function() {
						proxy.readyState = realXhr.readyState;
						proxy.status = realXhr.status;
						if (!realXhr.responseType || realXhr.responseType === "text") {
							proxy.responseText = realXhr.responseText;
						}
						if (realXhr.readyState === 4) {
							proxy.response = realXhr.response;
						}
						proxy.onreadystatechange && proxy.onreadystatechange();
					};
					realXhr.onload = function(e) { proxy.onload && proxy.onload(e); };
					realXhr.onerror = function(e) { proxy.onerror && proxy.onerror(e); };
					realXhr.open.apply(realXhr, arguments);
				}
			},
			send: function(inBody) {
				if (proxy.responseType && proxy.target !== mockXhr) {
					proxy.target.responseType = proxy.responseType;
				}
				if (proxy.target === mockXhr) {
					["onreadystatechange", "onload", "onerror", "onloadend"].forEach(function(inName) {
						mockXhr[inName] = function(e) {
							proxy.readyState = mockXhr.readyState;
							proxy.status = mockXhr.status;
							proxy.responseText = mockXhr.responseText;
							proxy[inName] && proxy[inName].call(proxy, e);
						};
					});
				}
				proxy.target.send(inBody);
			},
			abort: function() { proxy.target && proxy.target.abort(); },
			setRequestHeader: function(a, b) { proxy.target.setRequestHeader(a, b); },
			getResponseHeader: function(a) { return proxy.target.getResponseHeader(a); },
			getAllResponseHeaders: function() { return proxy.target.getAllResponseHeaders(); },
			addEventListener: function(a, b) { proxy.target.addEventListener(a, b); },
			removeEventListener: function() {},
			readyState: 0,
			status: 0,
			responseText: "",
			responseType: "",
			response: null
		};
		return proxy;
	};

	mock.set = function(inKind, inKey, inRule) {
		if (inKind === "luna") {
			(config.luna = config.luna || {})[inKey] = inRule;
		} else {
			(config.http = config.http || []).unshift(Object.assign({match: inKey}, inRule));
		}
	};

	// Network APIs the app does not use must not reach the internet either.
	if (window.fetch) {
		var realFetch = window.fetch;
		window.fetch = function(inUrl) {
			var url = String(inUrl && inUrl.url || inUrl);
			if (/^https?:\/\//i.test(url) && url.indexOf(location.origin) !== 0) {
				calls.push({kind: "http", url: url, unmocked: true});
				return Promise.reject(new TypeError("webos-mock: fetch is not mocked"));
			}
			return realFetch.apply(window, arguments);
		};
	}
})();
