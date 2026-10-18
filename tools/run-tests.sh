#!/bin/sh
# Runs the whole test suite:
#   1. syntax + ES5 check of source/ (the app's own code is ES5, like the
#      other Enyo 1 apps; lib/ is third-party and left as it comes)
#   2. unit tests of the readers, helpers and the service's writer
#      (tools/tests/unit-tests.js), and of editing (tools/tests/edit-tests.js)
#   3. end-to-end tests of the app in headless Chromium with mocks, offline
#      (tools/tests/e2e-tests.js)
#
#   tools/run-tests.sh [--screenshots]
#   TEST_FILTER=<regexp> tools/run-tests.sh      only matching tests
#   CHROME=/path/to/chromium tools/run-tests.sh
#   ENYO=/path/to/enyo-1.0-ports/framework tools/run-tests.sh
set -e
cd "$(dirname "$0")/.."
E2E_ARGS=""
for a in "$@"; do
	case "$a" in
		--screenshots) E2E_ARGS="--screenshots" ;;
		*) echo "unknown option: $a" >&2; exit 2 ;;
	esac
done

echo "== syntax / ES5"
for f in source/*.js source/*/*.js services/*/*.js; do node --check "$f"; done
# Arrow functions, let/const, template strings, classes: not in the app's code.
if grep -nE '=>|^\s*(let|const) |`|^\s*class ' source/*.js source/*/*.js; then
	echo "ES2015+ syntax in source/ (see above)" >&2
	exit 1
fi
for f in appinfo.json services/*/services.json services/*/package.json; do node -e 'JSON.parse(require("fs").readFileSync(process.argv[1], "utf8"))' "$f"; done
echo "ok"

echo "== unit tests"
node tools/tests/unit-tests.js

echo "== editing unit tests"
node tools/tests/edit-tests.js

echo "== end-to-end tests (mocked, offline)"
node tools/tests/e2e-tests.js $E2E_ARGS
