#!/usr/bin/env bash
set -euo pipefail

# =========================
# Edirep — Build macOS (.app + .dmg + sha256)
# =========================

APP_NAME="Edirep"

# -------------------------
# Sécurité : venv obligatoire
# -------------------------
if [[ -z "${VIRTUAL_ENV:-}" ]]; then
  echo "❌ Aucun environnement virtuel actif."
  echo "👉 Activez le venv : source venv/bin/activate"
  exit 1
fi

# -------------------------
# Version (obligatoire)
# -------------------------
if [[ $# -lt 1 ]]; then
  echo "❌ Version manquante."
  echo "👉 Usage : ./build-macos.sh 1.2.3"
  exit 1
fi
VERSION="$1"

# -------------------------
# Chemins
# -------------------------
ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
DIST_DIR="$ROOT_DIR/dist"
BUILD_DIR="$ROOT_DIR/build"
RELEASES_DIR="$ROOT_DIR/releases"
ASSETS_DIR="$ROOT_DIR/assets"

mkdir -p "$RELEASES_DIR"

# -------------------------
# Architecture macOS
# -------------------------
ARCH="$(uname -m)"   # x86_64 ou arm64

# -------------------------
# Noms de sortie
# -------------------------
APP_BUNDLE_NAME="${APP_NAME}-v${VERSION} - macOS - ${ARCH}.app"
DMG_NAME="${APP_NAME}-v${VERSION} - macOS - ${ARCH}.dmg"

APP_BUNDLE_PATH="$RELEASES_DIR/$APP_BUNDLE_NAME"
DMG_PATH="$RELEASES_DIR/$DMG_NAME"

ICON_ICNS="$ASSETS_DIR/logo.icns"

echo "▶ Build ${APP_NAME} macOS"
echo "   Version      : v${VERSION}"
echo "   Architecture : ${ARCH}"
echo "   APP          : releases/$APP_BUNDLE_NAME"
echo "   DMG          : releases/$DMG_NAME"

# -------------------------
# Nettoyage PyInstaller
# -------------------------
rm -rf "$DIST_DIR" "$BUILD_DIR"

# -------------------------
# Build PyInstaller (binaire temporaire)
# -------------------------
PYI_ICON_ARGS=()
if [[ -f "$ICON_ICNS" ]]; then
  PYI_ICON_ARGS+=(--icon "$ICON_ICNS")
else
  echo "⚠️  Icône absente : assets/logo.icns (l’app sera sans icône)"
fi

pyinstaller \
  --name "$APP_NAME" \
  --windowed \
  --onefile \
  --clean \
  --noconfirm \
  --add-data "assets:assets" \
  "${PYI_ICON_ARGS[@]}" \
  edirep.py

# =========================
# Création du bundle .app
# =========================
MACOS_DIR="$APP_BUNDLE_PATH/Contents/MacOS"
RESOURCES_DIR="$APP_BUNDLE_PATH/Contents/Resources"

rm -rf "$APP_BUNDLE_PATH"
mkdir -p "$MACOS_DIR" "$RESOURCES_DIR"

# Exécutable
mv "$DIST_DIR/$APP_NAME" "$MACOS_DIR/$APP_NAME"
chmod +x "$MACOS_DIR/$APP_NAME"

# Icône
ICON_PLIST_LINE=""
if [[ -f "$ICON_ICNS" ]]; then
  cp "$ICON_ICNS" "$RESOURCES_DIR/logo.icns"
  ICON_PLIST_LINE=$'  <key>CFBundleIconFile</key>\n  <string>logo</string>\n'
fi

# Info.plist
cat > "$APP_BUNDLE_PATH/Contents/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
 "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleName</key><string>${APP_NAME}</string>
  <key>CFBundleDisplayName</key><string>${APP_NAME}</string>
  <key>CFBundleIdentifier</key><string>com.mrklm.edirep</string>
  <key>CFBundleShortVersionString</key><string>${VERSION}</string>
  <key>CFBundleVersion</key><string>${VERSION}</string>
  <key>CFBundleExecutable</key><string>${APP_NAME}</string>
  <key>CFBundlePackageType</key><string>APPL</string>
${ICON_PLIST_LINE}</dict>
</plist>
EOF

echo "✅ App bundle créé"

# =========================
# Création du DMG
# =========================
DMG_STAGING="$RELEASES_DIR/.dmg-staging"
TMP_DMG="$RELEASES_DIR/.tmp-${APP_NAME}.dmg"

rm -rf "$DMG_STAGING" "$TMP_DMG" "$DMG_PATH"
mkdir -p "$DMG_STAGING"

cp -R "$APP_BUNDLE_PATH" "$DMG_STAGING/"
ln -s /Applications "$DMG_STAGING/Applications"

hdiutil create "$TMP_DMG" \
  -volname "${APP_NAME}" \
  -srcfolder "$DMG_STAGING" \
  -fs HFS+ \
  -ov >/dev/null

hdiutil convert "$TMP_DMG" \
  -format UDZO \
  -o "$DMG_PATH" >/dev/null

rm -rf "$DMG_STAGING" "$TMP_DMG"

echo "✅ DMG créé"

# -------------------------
# SHA256 du DMG
# -------------------------
(
  cd "$RELEASES_DIR"
  shasum -a 256 "$DMG_NAME" > "${DMG_NAME}.sha256"
)

echo "🔐 SHA256 généré : ${DMG_NAME}.sha256"
echo "🎉 Build macOS terminé."
