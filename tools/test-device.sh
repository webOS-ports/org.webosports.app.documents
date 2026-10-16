#!/bin/bash
# Installs the app on a LuneOS device over adb, copies the test documents to
# /media/internal/documents-test/, launches the app and runs
# tools/tests/device-tests.js against it in the real WAM.
#
#   tools/test-device.sh <adb-serial> [name] [--no-build] [--no-large]
#   TEST_ID=<id> tools/test-device.sh ...      install under another app id
#   WITH_SERVICE=1 tools/test-device.sh ...    with the app's service (see tools/build-ipk.sh)
#
# The large documents (test-output/large/, made by
# `python3 -I tools/tests/make-fixtures.py --large`) are copied too unless
# --no-large; the test then times them.
#
# Install goes through com.webos.appInstallService (opkg with its own root,
# /media/cryptofs/apps), like Preware, so system packages are not touched.
# Screenshots and results: test-output/device/<name>/.
set -e
cd "$(dirname "$0")/.."
SERIAL="$1"
NAME="${2:-$SERIAL}"
[ -n "$SERIAL" ] || { echo "usage: $0 <adb-serial> [name] [--no-build] [--no-large]" >&2; exit 2; }
ID=${TEST_ID:-$(node -p 'require("./appinfo.json").id')}
VERSION=$(node -p 'require("./appinfo.json").version')
IPK="build/${ID}_${VERSION}_all.ipk"
ADB="adb -s $SERIAL"
PORT=$((19000 + RANDOM % 900))
DIR=/media/internal/documents-test

case " $* " in
	*" --no-build "*) ;;
	*) tools/build-ipk.sh >/dev/null ;;
esac

echo "== $NAME: $($ADB shell getprop ro.product.device) / $($ADB shell uname -m)"
$ADB shell "mkdir -p $DIR"
for f in tools/tests/fixtures/*.*; do
	$ADB push "$f" "$DIR/" >/dev/null
done
case " $* " in
	*" --no-large "*) ;;
	*)
		for f in test-output/large/large.docx test-output/large/large.odt test-output/large/large.xlsx \
				test-output/large/large.ods test-output/large/large.csv; do
			[ -f "$f" ] && $ADB push "$f" "$DIR/" >/dev/null
		done
		;;
esac

$ADB push "$IPK" /tmp/ >/dev/null
# Close it first, and give the card time to go: an install that has to close an
# app still closing can wait on it for good.
$ADB shell "luna-send -n 1 -w 8000 luna://com.webos.service.applicationmanager/closeByAppId '{\"id\":\"$ID\"}' > /dev/null 2>&1; sleep 4"
# luna-send prints nothing on an adb pipe; results go through a file.
$ADB shell "luna-send -w 180000 -i luna://com.webos.appInstallService/install '{\"id\":\"$ID\",\"ipkUrl\":\"/tmp/$(basename "$IPK")\",\"subscribe\":true}' > /tmp/documents-inst.out 2>&1 & P=\$!; \
	for i in \$(seq 1 180); do sleep 1; grep -qE '\"installed\"|failed|errorCode' /tmp/documents-inst.out && break; done; kill \$P 2>/dev/null; \
	rm -f /tmp/$(basename "$IPK"); grep -oE '\"state\":\"installed\"|errorText\":\"[^\"]*\"|reason\":\"[^\"]*\"' /tmp/documents-inst.out | tail -n 1"
$ADB shell "luna-send -w 20000 -n 1 luna://com.webos.service.applicationmanager/launch '{\"id\":\"$ID\"}' > /tmp/documents-l.out 2>&1; cat /tmp/documents-l.out"
echo
$ADB forward tcp:$PORT tcp:9998 >/dev/null
trap '$ADB forward --remove tcp:$PORT >/dev/null 2>&1 || true' EXIT

RC=0
DEVICE_SERIAL="$SERIAL" APP_ID="$ID" DEVICE_DEBUG=http://127.0.0.1:$PORT DEVICE_NAME="$NAME" node tools/tests/device-tests.js || RC=$?

# The whole screen as the compositor draws it.
mkdir -p "test-output/device/$NAME"
$ADB shell "luna-send -w 10000 -n 1 luna://com.webos.surfacemanager/captureCompositorOutput '{\"output\":\"/tmp/documents-shot.png\",\"format\":\"PNG\"}' > /tmp/documents-c.out 2>&1" || true
$ADB pull /tmp/documents-shot.png "test-output/device/$NAME/compositor.png" >/dev/null 2>&1 && \
	echo "screen: test-output/device/$NAME/compositor.png" || echo "(no compositor screenshot)"
$ADB shell rm -f /tmp/documents-shot.png
exit $RC
