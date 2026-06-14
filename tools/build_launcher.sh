#!/bin/zsh
set -euo pipefail

ROOT="/Users/bala/Documents/Codex/Diet Plan"
ICON_HTML="$ROOT/assets/icon-art.html"
ICON_PNG="$ROOT/assets/diet-dashboard-icon-1024.png"
ICONSET_DIR="$ROOT/assets/DietDashboard.iconset"
ICNS_FILE="$ROOT/assets/diet-dashboard.icns"
APP_NAME="Diet Dashboard.app"
PROJECT_APP="$ROOT/$APP_NAME"
DESKTOP_APP="/Users/bala/Desktop/$APP_NAME"
URL="https://balamm27.github.io/Diet-Plan/"
CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

mkdir -p "$ROOT/assets"
rm -f "$ICON_PNG" "$ICNS_FILE"
rm -rf "$ICONSET_DIR"

"$CHROME" \
  --headless=new \
  --disable-gpu \
  --allow-file-access-from-files \
  --window-size=1024,1024 \
  --screenshot="$ICON_PNG" \
  "file://$ICON_HTML" >/dev/null 2>&1

mkdir -p "$ICONSET_DIR"
for size in 16 32 64 128 256 512; do
  sips -z "$size" "$size" "$ICON_PNG" --out "$ICONSET_DIR/icon_${size}x${size}.png" >/dev/null
done
sips -z 32 32 "$ICON_PNG" --out "$ICONSET_DIR/icon_16x16@2x.png" >/dev/null
sips -z 64 64 "$ICON_PNG" --out "$ICONSET_DIR/icon_32x32@2x.png" >/dev/null
sips -z 256 256 "$ICON_PNG" --out "$ICONSET_DIR/icon_128x128@2x.png" >/dev/null
sips -z 512 512 "$ICON_PNG" --out "$ICONSET_DIR/icon_256x256@2x.png" >/dev/null
cp "$ICON_PNG" "$ICONSET_DIR/icon_512x512@2x.png"

iconutil -c icns "$ICONSET_DIR" -o "$ICNS_FILE"

for target in "$PROJECT_APP" "$DESKTOP_APP"; do
  rm -rf "$target"
  osacompile -o "$target" <<APPLESCRIPT
on run
  open location "$URL"
end run
APPLESCRIPT
  cp "$ICNS_FILE" "$target/Contents/Resources/applet.icns"
  /usr/bin/touch "$target"
done

echo "Created launcher apps:"
echo " - $PROJECT_APP"
echo " - $DESKTOP_APP"
