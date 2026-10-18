#!/bin/sh
# Lanzador del sistema - Maternidad (macOS / Linux)
# Doble clic desde Terminal, o: sh Iniciar.sh
cd "$(dirname "$0")"

if command -v python3 >/dev/null 2>&1; then
    PY=python3
elif command -v python >/dev/null 2>&1; then
    PY=python
else
    echo "No se encontro Python en este equipo."
    echo "Instala desde python.org (macOS ya trae Tk; en Linux: sudo apt install python3-tk)."
    exit 1
fi

echo "Abriendo el sistema..."
"$PY" app.py
if [ $? -ne 0 ]; then
    echo ""
    echo "El sistema termino con errores. Revisa _ERRORES.log en esta carpeta."
    read -r dummy
fi
