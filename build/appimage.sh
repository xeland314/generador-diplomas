#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
BUILD_DIR="$PROJECT_DIR/build/build-linux"
DIST_DIR="$PROJECT_DIR/dist"
APPDIR="$DIST_DIR/GeneradorDiplomas.AppDir"
APPIMAGE_NAME="GeneradorDiplomas-x86_64.AppImage"

echo "=== Creating AppImage ==="

# Paso 1: ejecutar build-linux si no existe el binario
if [ ! -f "$DIST_DIR/generador-diplomas/generador-diplomas" ]; then
    echo "--- Binary not found, running build-linux.sh first ---"
    bash "$SCRIPT_DIR/build-linux.sh"
fi

# Paso 2: crear la estructura AppDir
echo "--- Setting up AppDir structure ---"
rm -rf "$APPDIR"
mkdir -p "$APPDIR/usr/bin"
mkdir -p "$APPDIR/usr/share/applications"
mkdir -p "$APPDIR/usr/share/icons/hicolor/256x256/apps"

# Paso 3: copiar el binario
cp -r "$DIST_DIR/generador-diplomas/"* "$APPDIR/usr/bin/"

# Paso 4: crear el desktop entry
cat > "$APPDIR/generador-diplomas.desktop" << 'EOF'
[Desktop Entry]
Type=Application
Name=Generador de Diplomas
Exec=generador-diplomas
Icon=generador-diplomas
Categories=Office;
Comment=Generador de certificados a partir de plantillas PowerPoint
Terminal=false
EOF

cp "$APPDIR/generador-diplomas.desktop" "$APPDIR/usr/share/applications/"

# Paso 5: buscar un icono o crear uno basico
if [ -f "$PROJECT_DIR/assets/diploma_escuela_v7.jpg" ]; then
    cp "$PROJECT_DIR/assets/diploma_escuela_v7.jpg" "$APPDIR/usr/share/icons/hicolor/256x256/apps/generador-diplomas.jpg"
    cp "$APPDIR/usr/share/icons/hicolor/256x256/apps/generador-diplomas.jpg" "$APPDIR/generador-diplomas.jpg"
fi

# Paso 6: crear el symlink AppRun
cat > "$APPDIR/AppRun" << 'APPRUN'
#!/bin/bash
SELF=$(readlink -f "$0")
HERE=${SELF%/*}
export PATH="${HERE}/usr/bin:$PATH"
exec "${HERE}/usr/bin/generador-diplomas" "$@"
APPRUN
chmod +x "$APPDIR/AppRun"

# Paso 7: descargar appimagetool si no existe
APPIMAGETOOL="$DIST_DIR/appimagetool"
if [ ! -f "$APPIMAGETOOL" ]; then
    echo "--- Downloading appimagetool ---"
    curl -L -o "$APPIMAGETOOL" \
        "https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage"
    chmod +x "$APPIMAGETOOL"
fi

# Paso 8: generar la AppImage
echo "--- Building AppImage ---"
"$APPIMAGETOOL" --no-appstream "$APPDIR" "$DIST_DIR/$APPIMAGE_NAME"

rm -rf "$APPDIR"

echo ""
echo "=== AppImage created ==="
echo "File: $DIST_DIR/$APPIMAGE_NAME"
