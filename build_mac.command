#!/bin/bash
#
# Costruisce TimbroEasyPDF.app su un Mac (Apple Silicon o Intel).
#
# COME SI USA (sul Mac):
#   1. Installa Python 3 da https://www.python.org/downloads/macos/  (una volta sola)
#   2. Copia questo file sul Mac, apri il Terminale e lancia:   bash build_mac.command
#      (oppure: chmod +x build_mac.command  e poi doppio clic sul file)
#
# Lo script scarica da solo il codice dal repository, compila l'app e mette
# TimbroEasyPDF.app (e un .zip da inviare) sulla Scrivania. Non installa nulla
# nel sistema: usa un ambiente Python isolato in una cartella temporanea.

set -e
echo "================================================"
echo "  Costruzione di TimbroEasyPDF per Mac"
echo "================================================"

# 1) Python 3 presente?
if ! command -v python3 >/dev/null 2>&1; then
  echo ""
  echo "!! Python 3 non è installato su questo Mac."
  echo "   Scaricalo e installalo da:  https://www.python.org/downloads/macos/"
  echo "   Poi rilancia questo script."
  open "https://www.python.org/downloads/macos/" 2>/dev/null || true
  exit 1
fi
echo "Python trovato: $(python3 --version)"

# 2) Cartella di lavoro temporanea + download del codice (ramo 'mac')
WORK="$(mktemp -d)"
echo ""
echo "Scarico il codice sorgente..."
cd "$WORK"
curl -fL -o src.tar.gz "https://github.com/ingcandidoalex-wq/TimbroEasyPDF/archive/refs/heads/mac.tar.gz"
tar xzf src.tar.gz
cd TimbroEasyPDF-mac

# 3) Ambiente Python isolato + dipendenze
echo ""
echo "Preparo l'ambiente e installo pymupdf + pyinstaller..."
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install pymupdf pyinstaller

# 4) Icona dell'app (facoltativa: se fallisce, si compila senza)
python -c "import timbro_pdf as tp; tp.icon_pixmap(1024).save('icon.png')" 2>/dev/null \
  && sips -s format icns icon.png --out icon.icns >/dev/null 2>&1 || true
ICON=""
[ -f icon.icns ] && ICON="--icon icon.icns"

# 5) Compilazione dell'app
echo ""
echo "Compilo l'app (circa un minuto)..."
pyinstaller --noconfirm --windowed --name TimbroEasyPDF $ICON \
  --osx-bundle-identifier com.cbcengineering.timbroeasypdf \
  --add-data "tessdata:tessdata" \
  --add-data "Guida TimbroEasyPDF.pdf:." \
  timbro_pdf.py

# 6) Copia sulla Scrivania + .zip da inviare (ditto conserva i permessi)
DEST="$HOME/Desktop"
rm -rf "$DEST/TimbroEasyPDF.app"
cp -R "dist/TimbroEasyPDF.app" "$DEST/"
cd "$DEST"
rm -f "TimbroEasyPDF-mac.zip"
ditto -c -k --keepParent "TimbroEasyPDF.app" "TimbroEasyPDF-mac.zip"

echo ""
echo "================================================"
echo "  FATTO!"
echo "  Sulla Scrivania trovi:"
echo "    - TimbroEasyPDF.app        (l'app da usare)"
echo "    - TimbroEasyPDF-mac.zip    (da mandare al collega)"
echo ""
echo "  Primo avvio: tasto destro sull'app -> Apri -> Apri"
echo "  (serve solo la prima volta, perché l'app non è firmata Apple)."
echo "================================================"
