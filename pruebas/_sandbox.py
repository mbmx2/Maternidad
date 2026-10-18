# Sandbox de pruebas: copia app.py a un tempdir y lo importa desde ahi.
# Asi las suites prueban el app.py actual sin ensuciar el proyecto con datos.
import shutil
import sys
import tempfile
from pathlib import Path


def preparar_app():
    raiz = Path(__file__).resolve().parent.parent
    destino = Path(tempfile.mkdtemp(prefix="maternidad_test_"))
    shutil.copy2(raiz / "app.py", destino / "app.py")
    # activos de producto que la app lee al arrancar (pedido 1.14: icono)
    if (raiz / "iconos").is_dir():
        shutil.copytree(raiz / "iconos", destino / "iconos")
    sys.path.insert(0, str(destino))
    import app
    return app, destino
