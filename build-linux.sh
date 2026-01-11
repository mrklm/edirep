#!/usr/bin/env bash
set -euo pipefail

# =========================
# Edirep — Build Linux + AppImage + SHA256
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
  echo "👉 Usage : ./build-linux.sh 1.2.3"
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
# Architecture
# -------------------------
ARCH="$(uname -m)"  # x86_64, aarch64, etc.

# Nom de sortie (comme Garage)
OUTPUT_BIN_NAME="${APP_NAME}-v${VERSION} - Linux - ${ARCH}"
OUTPUT_APPIMAGE_NAME="${APP_NAME}-v${VERSION} - Linux - ${ARCH}.AppImage"

echo "▶ Build ${APP_NAME} Linux"
echo "   Version      : v${VERSION}"
echo "   Architecture : ${ARCH}"
echo "   Output BIN   : ${OUTPUT_BIN_NAME}"
echo "   Output AI    : ${OUTPUT_APPIMAGE_NAME}"

# -------------------------
# Nettoyage PyInstaller
# -------------------------
rm -rf "$DIST_DIR" "$BUILD_DIR"

# -------------------------
# Build PyInstaller (binaire)
# -------------------------
pyinstaller \
  --name "$APP_NAME" \
  --windowed \
  --onefile \
  --clean \
  --noconfirm \
  --add-data "assets:assets" \
  edirep.py

# Déplacement / renommage du binaire
mv "$DIST_DIR/$APP_NAME" "$RELEASES_DIR/$OUTPUT_BIN_NAME"
chmod +x "$RELEASES_DIR/$OUTPUT_BIN_NAME"

echo "✅ Binaire créé : releases/$OUTPUT_BIN_NAME"

# -------------------------
# SHA256 du binaire
# -------------------------
(
  cd "$RELEASES_DIR"
  sha256sum "$OUTPUT_BIN_NAME" > "${OUTPUT_BIN_NAME}.sha256"
)
echo "🔐 SHA256 binaire : releases/${OUTPUT_BIN_NAME}.sha256"

# =========================
# AppImage
# =========================

# On cherche appimagetool
APPIMAGETOOL=""
if command -v appimagetool >/dev/null 2>&1; then
  APPIMAGETOOL="appimagetool"
elif [[ -x "$ROOT_DIR/tools/appimagetool" ]]; then
  APPIMAGETOOL="$ROOT_DIR/tools/appimagetool"
fi

if [[ -z "$APPIMAGETOOL" ]]; then
  echo "❌ appimagetool introuvable."
  echo "👉 Installe appimagetool (ou place-le dans ./tools/appimagetool), puis relance."
  echo "   (On s'arrête ici après avoir produit le binaire + sha256.)"
  exit 0
fi

# Répertoires AppDir (jetables)
APPDIR="$ROOT_DIR/AppDir"
APPDIR_USR_BIN="$APPDIR/usr/bin"
APPDIR_USR_SHARE="$APPDIR/usr/share"
APPDIR_USR_SHARE_APPS="$APPDIR_USR_SHARE/applications"
APPDIR_USR_SHARE_ICONS="$APPDIR_USR_SHARE/icons/hicolor/256x256/apps"

rm -rf "$APPDIR"
mkdir -p "$APPDIR_USR_BIN" "$APPDIR_USR_SHARE_APPS" "$APPDIR_USR_SHARE_ICONS"

# 1) binaire dans AppDir
cp "$RELEASES_DIR/$OUTPUT_BIN_NAME" "$APPDIR_USR_BIN/$APP_NAME"
chmod +x "$APPDIR_USR_BIN/$APP_NAME"

# 2) assets dans AppDir (si ton app en a besoin au runtime)
mkdir -p "$APPDIR/usr/assets"
cp -R "$ROOT_DIR/assets" "$APPDIR/usr/assets/assets"

# 3) desktop file
cat > "$APPDIR_USR_SHARE_APPS/${APP_NAME}.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Edirep
Exec=Edirep
Icon=edirep
Categories=Utility;
Terminal=false
EOF

# 4) icône (fallback : logo.png)
# (Idéalement tu fournis un assets/edirep.png 256x256 ; sinon on réutilise assets/logo.png)
if [[ -f "$ROOT_DIR/assets/edirep.png" ]]; then
  cp "$ROOT_DIR/assets/edirep.png" "$APPDIR_USR_SHARE_ICONS/edirep.png"
elif [[ -f "$ROOT_DIR/assets/logo.png" ]]; then
  cp "$ROOT_DIR/assets/logo.png" "$APPDIR_USR_SHARE_ICONS/edirep.png"
else
  echo "⚠️ Aucune icône PNG trouvée (assets/edirep.png ou assets/logo.png). AppImage OK mais icône manquante."
fi

# 5) AppRun
cat > "$APPDIR/AppRun" <<'EOF'
#!/bin/sh
HERE="$(dirname "$(readlink -f "$0")")"
export PATH="$HERE/usr/bin:$PATH"
exec "$HERE/usr/bin/Edirep" "$@"
EOF
chmod +x "$APPDIR/AppRun"

# 6) lien “sympa” attendu par certains outils (optionnel mais utile)
# Copie desktop + icône à la racine d'AppDir (convention fréquente)
cp "$APPDIR_USR_SHARE_APPS/${APP_NAME}.desktop" "$APPDIR/${APP_NAME}.desktop" || true
if [[ -f "$APPDIR_USR_SHARE_ICONS/edirep.png" ]]; then
  cp "$APPDIR_USR_SHARE_ICONS/edirep.png" "$APPDIR/edirep.png" || true
fi

# 7) Build AppImage
# Certains environnements ont besoin de cette variable pour éviter des warnings
export ARCH="$ARCH"

"$APPIMAGETOOL" "$APPDIR" "$RELEASES_DIR/$OUTPUT_APPIMAGE_NAME"

chmod +x "$RELEASES_DIR/$OUTPUT_APPIMAGE_NAME"
echo "✅ AppImage créée : releases/$OUTPUT_APPIMAGE_NAME"

# -------------------------
# SHA256 AppImage
# -------------------------
(
  cd "$RELEASES_DIR"
  sha256sum "$OUTPUT_APPIMAGE_NAME" > "${OUTPUT_APPIMAGE_NAME}.sha256"
)
echo "🔐 SHA256 AppImage : releases/${OUTPUT_APPIMAGE_NAME}.sha256"

# Nettoyage AppDir (optionnel)
rm -rf "$APPDIR"

echo "🎉 Terminé."
