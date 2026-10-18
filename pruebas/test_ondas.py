# Pruebas de las ondas T2: fecha futura (C2) y parseo de filtros (A1)
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _sandbox import preparar_app

app, _ = preparar_app()
FALLAS = []


def check(nombre, cond, detalle=""):
    print(f"[{'OK ' if cond else 'FAIL'}] {nombre}" + (f"  <- {detalle}" if (detalle and not cond) else ""))
    if not cond:
        FALLAS.append(nombre)


app.inicializar_estructura()

# --- C2: fecha de nacimiento futura ---
llamadas = []


def cb_no(mensaje):
    llamadas.append(mensaje)
    return False


ok, msg, ruta, _ = app.crear_paciente("FECHA", "FUTURA", "ANA", "2099-01-01", "",
                                      ventana_callback=cb_no)
check("C2 futura sin confirmar -> rechazada", not ok and "posterior a hoy" in msg, msg)
check("C2 muestra advertencia al usuario", len(llamadas) == 1, str(llamadas))
check("C2 no crea carpeta al rechazar", ruta is None)

ok2, msg2, ruta2, _ = app.crear_paciente("FECHA", "FUTURA", "ANA", "2099-01-01", "",
                                         ventana_callback=lambda m: True)
check("C2 futura confirmada -> aceptada", ok2 and ruta2 is not None, msg2)

llamadas.clear()
ok3, msg3, _, _ = app.crear_paciente("HOY", "TEST", "MARIA", app.fecha_hoy(), "",
                                     ventana_callback=lambda m: (llamadas.append(m), True)[1])
check("C2 fecha de hoy no dispara advertencia", ok3 and not llamadas,
      f"{msg3} llamadas={llamadas}")
check("C2 pasada no dispara advertencia", app.crear_paciente(
    "PASADA", "TEST", "LUIS", "1985-05-05", "",
    ventana_callback=lambda m: (llamadas.append(m), True)[1])[0] and not llamadas)

# --- A1: parseo de fechas para filtros ---
check("Filtro parsea dd/mm/aaaa a ISO", app.parsear_fecha("15/03/1985") == "1985-03-15",
      app.parsear_fecha("15/03/1985"))
check("Filtro parsea dd-MMM-aaaa a ISO", app.parsear_fecha("15-JUN-1990") == "1990-06-15")

app.crear_paciente("RANGO", "TEST", "PEDRO", "15/06/1990", "")
todos = app.listar_pacientes()

dentro = app.filtrar(todos, desde="1980-01-01", hasta="1995-12-31")
fechas = [app.cargar_paciente(r).get("fecha_nacimiento", "") for _, r in dentro]
check("Filtro rango incluye solo fechas dentro",
      fechas and all("1980-01-01" <= f <= "1995-12-31" for f in fechas if f), str(fechas))

fuera = app.filtrar(todos, hasta="1970-01-01")
check("Filtro excluye fechas posteriores al 'hasta'",
      all(app.cargar_paciente(r).get("fecha_nacimiento", "") > "1970-01-01" for _, r in fuera))

sin_filtro = app.filtrar(todos)
check("Sin filtros la lista no cambia", len(sin_filtro) == len(todos))

print()
print("=" * 60)
if FALLAS:
    print(f"RESULTADO: {len(FALLAS)} FALLAS")
    for f in FALLAS:
        print(f"  - {f}")
    sys.exit(1)
print("RESULTADO: TODAS LAS PRUEBAS DE ONDAS PASARON")
