#!/bin/bash
# Builds build/<id>_<version>_all.ipk: a plain ar archive with debian-binary,
# control.tar.gz and data.tar.gz. No palm-package or SDK needed.
#
#   usr/palm/applications/<id>/           the app (no tests, tools or desktop entry point)
#   usr/palm/services/<id>.service/       its service (writes saved copies to new files), with WITH_SERVICE=1
#   usr/palm/packages/<id>/packageinfo.json  names both, so appinstalld registers the service
#
# The service is left out by default: on images whose appinstalld lacks the
# applicationinstall.interface group (appinstalld2 before 2e06f07) a package
# with a JS service installs once but can then be neither updated nor removed
# (the installer's /quit to the service is denied). Without the service the
# app works the same, except that "Save a Copy" says it is not available.
#
#   tools/build-ipk.sh                       app only
#   WITH_SERVICE=1 tools/build-ipk.sh        app and service
#   TEST_ID=<id> tools/build-ipk.sh          the same under another id
set -e
cd "$(dirname "$0")/.."

ID=$(node -p 'require("./appinfo.json").id')
# TEST_ID=<id> builds the same app under another id (and title), e.g. to test
# beside an installed copy; the service is renamed with it.
APPID=$ID
if [ -n "$TEST_ID" ]; then
	ID=$TEST_ID
fi
VERSION=$(node -p 'require("./appinfo.json").version')
TITLE=$(node -p 'require("./appinfo.json").title')
MAINTAINER=$(node -p 'require("./appinfo.json").maintainer')
SERVICE="$ID.service"

WORK=build/ipk
rm -rf "$WORK"
APP="$WORK/data/usr/palm/applications/$ID"
mkdir -p "$WORK/control" "$APP" "$WORK/data/usr/palm/services" "$WORK/data/usr/palm/packages/$ID"

cp -r framework_config.json depends.js index.html icon.png icon-256x256.png \
	source css images lib LICENSE NOTICE "$APP/"
node -e 'var a = require("./appinfo.json"); if (process.argv[1] !== a.id) { a.id = process.argv[1]; a.title += " (test)"; }
	require("fs").writeFileSync(process.argv[2], JSON.stringify(a, null, "\t") + "\n");' "$ID" "$APP/appinfo.json"
if [ "$WITH_SERVICE" = 1 ]; then
cp -r "services/$APPID.service" "$WORK/data/usr/palm/services/$SERVICE"
node -e 'var f = process.argv[1], s = require(f); s.id = process.argv[2]; s.services[0].name = process.argv[2];
	require("fs").writeFileSync(f, JSON.stringify(s, null, "\t") + "\n");' "$PWD/$WORK/data/usr/palm/services/$SERVICE/services.json" "$SERVICE"
cat > "$WORK/data/usr/palm/packages/$ID/packageinfo.json" <<JSON
{
	"id": "$ID",
	"version": "$VERSION",
	"app": "$ID",
	"services": ["$SERVICE"]
}
JSON
else
	rmdir "$WORK/data/usr/palm/services" "$WORK/data/usr/palm/packages/$ID" "$WORK/data/usr/palm/packages"
fi

SIZE=$(du -sk "$WORK/data" | cut -f1)
cat > "$WORK/control/control" <<CONTROL
Package: $ID
Version: $VERSION
Section: misc
Priority: optional
Architecture: all
Installed-Size: $SIZE
Maintainer: $MAINTAINER
Description: $TITLE
 Reads Word, OpenDocument, Excel, PowerPoint, CSV, text and Markdown files.
webOS-Package-Format-Version: 2
webOS-Packager-Version: x.y.x
CONTROL

echo "2.0" > "$WORK/debian-binary"
TAROPTS="--owner=0 --group=0 --numeric-owner --sort=name"
tar $TAROPTS -czf "$WORK/control.tar.gz" -C "$WORK/control" ./control
tar $TAROPTS -czf "$WORK/data.tar.gz" -C "$WORK/data" ./usr

OUT="build/${ID}_${VERSION}_all.ipk"
rm -f "$OUT"
(cd "$WORK" && ar rc "../$(basename "$OUT")" debian-binary control.tar.gz data.tar.gz)
echo "$OUT"
