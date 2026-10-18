# Maternidad — Expedientes clínicos offline

Sistema de expedientes para consultorio de maternidad: registro de pacientes,
episodios de embarazo (FUM/FPP/EG), historia clínica NOM-004, controles
prenatales, gráficos de progreso, calculadoras médicas y alertas. **Funciona
100 % sin internet**: todos los datos viven en carpetas locales de la PC del
consultorio.

## Requisitos

- Python 3.10 o superior **con tkinter** (viene incluido en la instalación
  estándar de Windows; en Linux: `sudo apt install python3-tk`).
- Opcional: `pip install cryptography` para respaldos cifrados (sin él, los
  respaldos quedan con integridad SHA-256 pero sin cifrado).

## Cómo usarlo

- **Windows:** doble clic en `Iniciar.bat` (o leer `LEEME_WINDOWS.txt`).
- **macOS / Linux:** `sh Iniciar.sh` (o `python3 app.py`).

Funciona en **Windows, macOS y Linux** con el mismo archivo (las decisiones
de sistema operativo viven en una sola sección del código; `app.py` sección
00c). Al primer arranque crea sus carpetas de trabajo (`EXPEDIENTES/`,
`_CONFIG.json`, `_BITACORA.log`, etc.). Esas carpetas contienen los datos del
consultorio y **nunca** deben copiarse a ningún repositorio (el `.gitignore`
ya las excluye).

## Cómo verificar que funciona

```bash
python3 -m py_compile app.py
python3 pruebas/test_prototipo.py
python3 pruebas/test_calendario.py
python3 pruebas/test_ondas.py
python3 pruebas/test_ui_smoke.py   # requiere pantalla (display)
```

## Estructura

| Ruta | Contenido |
|------|-----------|
| `app.py` | El sistema completo (un solo archivo, biblioteca estándar de Python) |
| `esquemas/` | Esquemas canónicos: historia clínica NOM-004 y catálogos clínicos |
| `pruebas/` | Suites de verificación automáticas |
| `iconos/` | Ícono del producto (PNG maestro; la app lo muestra en ventana y barra de tareas) |
| `Iniciar.bat` · `Iniciar.sh` · `LEEME_WINDOWS.txt` | Arranque y guía para la PC del consultorio |

En **Linux**, para que el ícono aparezca en los menús: copia `maternidad.desktop`
(plantilla en el repo privado o genera uno con `Exec`/`Icon` apuntando a tus
rutas absolutas) a `~/.local/share/applications/` y ejecuta
`xdg-icon-resource install --novendor --size 256 iconos/icono_maternidad_256.png maternidad`.

## Aviso

Software de uso personal del consultorio. No es un dispositivo médico
certificado ni sustituye el juicio clínico: las escalas y calculadoras
estratifican riesgo, no diagnostican. Los umbrales clínicos son cifras del
médico dueño del sistema.
