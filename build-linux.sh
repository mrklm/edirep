#!/usr/bin/env bash
set -euo pipefail

# =========================
# Edirep — Build Linux + AppImage + tar.gz + SHA256
# Sorties: ./releases/
# Usage : ./build-linux.sh 3.10
# =========================

APP_NAME="Edirep"
ENTRYPOINT="edirep.py"

# -------------------------
# venv obligatoire
# -------------------------
if [[ -z "${VIRTUAL_ENV:-}" ]]; then
  echo "❌ Aucun environnement virtuel actif."
  echo "👉 Active : source .venv/bin/activate  (ou venv/bin/activate)"
  exit 1
fi

# -------------------------
# Version obligatoire
# -------------------------
if [[ $# -lt 1 ]]; then
  echo "❌ Version manquante."
  echo "👉 Usage : ./build-linux.sh 3.10"
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
mkdir -p "$RELEASES_DIR"

# -------------------------
# Outils
# -------------------------
command -v pyinstaller >/dev/null 2>&1 || {
  echo "❌ pyinstaller introuvable dans ce venv."
  echo "👉 pip install pyinstaller"
  exit 1
}

if [[ ! -f "$ROOT_DIR/$ENTRYPOINT" ]]; then
  echo "❌ Entrypoint introuvable: $ENTRYPOINT"
  exit 1
fi

if [[ ! -d "$ROOT_DIR/assets" ]]; then
  echo "❌ Dossier ./assets introuvable."
  echo "👉 Ton code attend au moins assets/logo.png et assets/AIDE.md."
  exit 1
fi

# -------------------------
# Arch
# -------------------------
ARCH="$(uname -m)"
case "$ARCH" in
  x86_64|aarch64) : ;;
  *)
    echo "❌ Architecture non supportée (AppImage): $ARCH"
    echo "👉 Supportées: x86_64, aarch64"
    exit 1
    ;;
esac

# -------------------------
# Noms de sortie (SANS espaces)
# -------------------------
OUTPUT_BIN_NAME="${APP_NAME}-v${VERSION}-linux-${ARCH}"
OUTPUT_APPIMAGE_NAME="${APP_NAME}-v${VERSION}-linux-${ARCH}.AppImage"
OUTPUT_TAR_NAME="${APP_NAME}-v${VERSION}-linux-${ARCH}.tar.gz"

echo "▶ Build ${APP_NAME} Linux"
echo "   Version      : v${VERSION}"
echo "   Architecture : ${ARCH}"
echo "   Output BIN   : ${OUTPUT_BIN_NAME}"
echo "   Output AI    : ${OUTPUT_APPIMAGE_NAME}"
echo "   Output TAR   : ${OUTPUT_TAR_NAME}"
echo

# -------------------------
# Nettoyage
# -------------------------
rm -rf "$DIST_DIR" "$BUILD_DIR"

# -------------------------
# PyInstaller (onefile)
# assets embarqués → ton resource_path() les trouvera via sys._MEIPASS
# -------------------------
pyinstaller \
  --name "$APP_NAME" \
  --windowed \
  --onefile \
  --clean \
  --noconfirm \
  --add-data "assets:assets" \
  "$ENTRYPOINT"

if [[ ! -f "$DIST_DIR/$APP_NAME" ]]; then
  echo "❌ PyInstaller n'a pas produit dist/$APP_NAME"
  exit 1
fi

mv "$DIST_DIR/$APP_NAME" "$RELEASES_DIR/$OUTPUT_BIN_NAME"
chmod +x "$RELEASES_DIR/$OUTPUT_BIN_NAME"
echo "✅ Binaire créé : releases/$OUTPUT_BIN_NAME"

# SHA256 binaire
(
  cd "$RELEASES_DIR"
  sha256sum "$OUTPUT_BIN_NAME" > "${OUTPUT_BIN_NAME}.sha256"
)
echo "🔐 SHA256 binaire : releases/${OUTPUT_BIN_NAME}.sha256"

# -------------------------
# tar.gz (binaire)
# -------------------------
(
  cd "$RELEASES_DIR"
  tar -czf "$OUTPUT_TAR_NAME" "$OUTPUT_BIN_NAME"
)
echo "📦 Tar.gz créé : releases/$OUTPUT_TAR_NAME"

(
  cd "$RELEASES_DIR"
  sha256sum "$OUTPUT_TAR_NAME" > "${OUTPUT_TAR_NAME}.sha256"
)
echo "🔐 SHA256 tar.gz : releases/${OUTPUT_TAR_NAME}.sha256"

echo

# =========================
# AppImage
# =========================

# Trouver appimagetool
APPIMAGETOOL=""
if command -v appimagetool >/dev/null 2>&1; then
  APPIMAGETOOL="appimagetool"
elif [[ -x "$ROOT_DIR/tools/appimagetool" ]]; then
  APPIMAGETOOL="$ROOT_DIR/tools/appimagetool"
fi

if [[ -z "$APPIMAGETOOL" ]]; then
  echo "⚠️ appimagetool introuvable."
  echo "👉 Installe appimagetool (ou place-le dans ./tools/appimagetool)."
  echo "✅ On s'arrête ici après binaire + tar.gz + sha256."
  exit 0
fi

APPDIR="$ROOT_DIR/AppDir"
APPDIR_USR_BIN="$APPDIR/usr/bin"
APPDIR_USR_SHARE="$APPDIR/usr/share"
APPDIR_USR_SHARE_APPS="$APPDIR_USR_SHARE/applications"
APPDIR_USR_SHARE_ICONS="$APPDIR_USR_SHARE/icons/hicolor/256x256/apps"

rm -rf "$APPDIR"
mkdir -p "$APPDIR_USR_BIN" "$APPDIR_USR_SHARE_APPS" "$APPDIR_USR_SHARE_ICONS"

# 1) binaire dans AppDir (nom interne stable: Edirep)
cp "$RELEASES_DIR/$OUTPUT_BIN_NAME" "$APPDIR_USR_BIN/$APP_NAME"
chmod +x "$APPDIR_USR_BIN/$APP_NAME"

# 2) desktop file
cat > "$APPDIR_USR_SHARE_APPS/${APP_NAME}.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Edirep
Exec=Edirep
Icon=edirep
Categories=Utility;
Terminal=false
EOF

# 3) icône (prend edirep.png si dispo sinon logo.png)
if [[ -f "$ROOT_DIR/assets/edirep.png" ]]; then
  cp "$ROOT_DIR/assets/edirep.png" "$APPDIR_USR_SHARE_ICONS/edirep.png"
elif [[ -f "$ROOT_DIR/assets/logo.png" ]]; then
  cp "$ROOT_DIR/assets/logo.png" "$APPDIR_USR_SHARE_ICONS/edirep.png"
else
  echo "⚠️ Aucune icône PNG trouvée (assets/edirep.png ou assets/logo.png)."
fi

# 4) AppRun
cat > "$APPDIR/AppRun" <<'EOF'
#!/bin/sh
HERE="$(dirname "$(readlink -f "$0")")"
export PATH="$HERE/usr/bin:$PATH"
exec "$HERE/usr/bin/Edirep" "$@"
EOF
chmod +x "$APPDIR/AppRun"

# 5) conventions utiles : desktop + icône à la racine
cp "$APPDIR_USR_SHARE_APPS/${APP_NAME}.desktop" "$APPDIR/${APP_NAME}.desktop" || true
if [[ -f "$APPDIR_USR_SHARE_ICONS/edirep.png" ]]; then
  cp "$APPDIR_USR_SHARE_ICONS/edirep.png" "$APPDIR/edirep.png" || true
fi

# 6) Build AppImage
export ARCH="$ARCH"
"$APPIMAGETOOL" "$APPDIR" "$RELEASES_DIR/$OUTPUT_APPIMAGE_NAME"

chmod +x "$RELEASES_DIR/$OUTPUT_APPIMAGE_NAME"
echo "✅ AppImage créée : releases/$OUTPUT_APPIMAGE_NAME"

# SHA256 AppImage
(
  cd "$RELEASES_DIR"
  sha256sum "$OUTPUT_APPIMAGE_NAME" > "${OUTPUT_APPIMAGE_NAME}.sha256"
)
echo "🔐 SHA256 AppImage : releases/${OUTPUT_APPIMAGE_NAME}.sha256"

# Nettoyage AppDir
rm -rf "$APPDIR"

echo
echo "🎉 Terminé."
echo "📁 Dossier final : releases/"
