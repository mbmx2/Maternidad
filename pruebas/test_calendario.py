# Pruebas del sistema de calendario (utilidades de fecha y episodios) — headless
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

# Aislamiento: esta suite usa MEDINA; se elimina si existe de corridas previas
import shutil
_ruta_base = app.EXPEDIENTES_DIR / "M" / "MEDINA_ROJAS_CLAUDIA_19900720"
if _ruta_base.exists():
    shutil.rmtree(_ruta_base)

# --- A3: calcular_fpp (Naegele) ---
check("FPP: FUM 2026-06-10 -> 2027-03-17",
      app.calcular_fpp("2026-06-10") == "2027-03-17", app.calcular_fpp("2026-06-10"))
check("FPP: FUM 2024-01-05 -> 2024-10-12",
      app.calcular_fpp("2024-01-05") == "2024-10-12", app.calcular_fpp("2024-01-05"))
check("FPP: trampa Feb-29 (FUM 2024-05-22 -> 2025-02-28)",
      app.calcular_fpp("2024-05-22") == "2025-02-28", app.calcular_fpp("2024-05-22"))
check("FPP: cruzamiento de anio (FUM 2025-12-30 -> 2026-10-06)",
      app.calcular_fpp("2025-12-30") == "2026-10-06", app.calcular_fpp("2025-12-30"))
check("FPP: invalida -> None", app.calcular_fpp("no-fecha") is None)
check("FPP: vacia -> None", app.calcular_fpp("") is None)

# --- A3: calcular_eg ---
hoy = app.fecha_hoy()
fum_hoy = hoy
check("EG: FUM hoy -> '0 s 0 d'", app.calcular_eg(fum_hoy) == "0 s 0 d",
      app.calcular_eg(fum_hoy))
import datetime as _dt
fum_19 = (_dt.date.fromisoformat(hoy) - _dt.timedelta(days=19)).isoformat()
check("EG: FUM hace 19 dias -> '2 s 5 d'", app.calcular_eg(fum_19) == "2 s 5 d",
      app.calcular_eg(fum_19))
fum_futura = (_dt.date.fromisoformat(hoy) + _dt.timedelta(days=10)).isoformat()
check("EG: FUM futura -> None (auditoria M2)", app.calcular_eg(fum_futura) is None)
check("EG: invalida -> None", app.calcular_eg("xx") is None)

# --- D1: round-trip legible <-> parsear_fecha ---
check("Round-trip: fecha_legible(1985-03-15) re-parsea al ISO",
      app.parsear_fecha(app.fecha_legible("1985-03-15")) == "1985-03-15",
      app.fecha_legible("1985-03-15"))

# --- B06: dias_en_mes con bisiestos ---
check("B06 febrero 2024 (bisiesto) -> 29", app.dias_en_mes(2024, 2) == 29,
      app.dias_en_mes(2024, 2))
check("B06 febrero 2023 -> 28", app.dias_en_mes(2023, 2) == 28,
      app.dias_en_mes(2023, 2))
check("B06 febrero 1900 (no bisiesto) -> 28", app.dias_en_mes(1900, 2) == 28,
      app.dias_en_mes(1900, 2))
check("B06 diciembre -> 31", app.dias_en_mes(2026, 12) == 31,
      app.dias_en_mes(2026, 12))

# --- Tier A (v2.0): edad_anios ---
check("edad 2000-01-01 @ 2026-10-16 -> 26", app.edad_anios("2000-01-01", "2026-10-16") == 26,
      app.edad_anios("2000-01-01", "2026-10-16"))
check("edad cumple manana -> aun no cumple", app.edad_anios("2000-10-17", "2026-10-16") == 25,
      app.edad_anios("2000-10-17", "2026-10-16"))
check("edad cumple hoy -> cumple", app.edad_anios("2000-10-16", "2026-10-16") == 26,
      app.edad_anios("2000-10-16", "2026-10-16"))
check("edad de fecha futura -> None", app.edad_anios("2030-01-01", "2026-10-16") is None)
check("edad de fecha invalida -> None", app.edad_anios("no-fecha") is None)
check("edad sin hoy usa el dia actual (nac 1-ene -> anio actual - 1990)",
      app.edad_anios("1990-01-01") == app.datetime.date.today().year - 1990,
      app.edad_anios("1990-01-01"))

# --- ADR-0027 (v2.5): cuentas regresivas y stats ---
check("dias_para_fpp positivo (futura)", app.dias_para_fpp("2027-03-17", "2026-10-16") == 152,
      app.dias_para_fpp("2027-03-17", "2026-10-16"))
check("dias_para_fpp negativa (pasada)", app.dias_para_fpp("2026-10-01", "2026-10-16") == -15,
      app.dias_para_fpp("2026-10-01", "2026-10-16"))
check("dias_para_fpp invalida -> None", app.dias_para_fpp("x") is None)
check("dias_desde hoy mismo -> 0", app.dias_desde("2026-10-16", "2026-10-16") == 0)
check("trimestre I (<14 s)", app.trimestre_desde_eg("13 s 6 d") == "I")
check("trimestre II (14 s)", app.trimestre_desde_eg("14 s 0 d") == "II")
check("trimestre II (27 s)", app.trimestre_desde_eg("27 s 6 d") == "II")
check("trimestre III (>=28 s)", app.trimestre_desde_eg("28 s 0 d") == "III")
check("trimestre sin EG/vacio -> ''",
      app.trimestre_desde_eg(None) == "" and app.trimestre_desde_eg("") == ""
      and app.trimestre_desde_eg("malo") == "")
st = app.stats_serie([("2026-06-01", 70.5), ("2026-06-15", 71.2)])
check("stats_serie: ultimo, delta y dias_hace",
      st is not None and st[0] == 71.2 and abs(st[3] - 0.7) < 1e-9
      and isinstance(st[2], int), str(st))
check("stats_serie sin fecha -> dias_hace None",
      app.stats_serie([("n/x", 5.0)])[2] is None)
check("stats_serie vacia -> None", app.stats_serie([]) is None)

# --- B96: stats_serie extendido (ADR-0045/v6.0) min/max/media ---
st2 = app.stats_serie([("2026-06-01", 70.0), ("2026-06-10", 72.5),
                       ("2026-06-20", 71.0)])
check("B96 retrocompatible: 4 primeros intactos",
      st2 is not None and st2[0] == 71.0 and st2[1] == "2026-06-20"
      and isinstance(st2[2], int) and abs(st2[3] - (-1.5)) < 1e-9, str(st2))
check("B96 min/max/media correctos",
      st2 is not None and st2[4] == 70.0 and st2[5] == 72.5
      and abs(st2[6] - (70.0 + 72.5 + 71.0) / 3) < 1e-9, str(st2))
check("B96 vacia -> None",
      app.stats_serie([]) is None)
st_no_num = app.stats_serie([("2026-06-01", "texto")])
check("B96 sin numeros -> min/max/media None",
      st_no_num is not None and st_no_num[4] is None
      and st_no_num[5] is None and st_no_num[6] is None, str(st_no_num))

# --- B95: segmentos_temporales (ADR-0045/v6.0) ---
# serie regular (hueco < 45 d) -> 1 segmento temporal con dias relativos
t_ok, segs = app.segmentos_temporales([("2026-06-01", 70.0),
                                       ("2026-06-15", 71.0),
                                       ("2026-06-29", 72.0)])
check("B95 serie regular -> 1 segmento temporal, dias 0/14/28",
      t_ok and len(segs) == 1
      and [x for x, _v, _e in segs[0]] == [0.0, 14.0, 28.0], f"{t_ok} {segs}")
# hueco > 45 d corta el segmento (fidelidad, no juicio clinico)
t_h, segs_h = app.segmentos_temporales([("2026-06-01", 70.0),
                                        ("2026-06-20", 71.0),
                                        ("2026-09-15", 72.0)])  # hueco 87 d
check("B95 hueco 87 d -> 2 segmentos",
      t_h and len(segs_h) == 2
      and [x for x, _v, _e in segs_h[0]] == [0.0, 19.0]
      and [x for x, _v, _e in segs_h[1]] == [106.0], f"{t_h} {segs_h}")
# hueco justo en el umbral (45) NO corta; 46 sí
t_u, segs_u = app.segmentos_temporales([("2026-06-01", 1.0),
                                        ("2026-07-16", 2.0)])  # 45 d
t_u2, segs_u2 = app.segmentos_temporales([("2026-06-01", 1.0),
                                           ("2026-07-17", 2.0)])  # 46 d
check("B95 umbral exacto: 45 une, 46 corta",
      t_u and len(segs_u) == 1 and t_u2 and len(segs_u2) == 2,
      f"{len(segs_u)} vs {len(segs_u2)}")
# etiqueta no parseable -> modo indice para TODA la serie (v1.x)
t_i, segs_i = app.segmentos_temporales([("Semana 12", 70.0),
                                        ("2026-06-15", 71.0)])
check("B95 etiqueta no parseable -> modo indice, 1 segmento",
      (not t_i) and len(segs_i) == 1
      and [x for x, _v, _e in segs_i[0]] == [0.0, 1.0], f"{t_i} {segs_i}")
# dias relativos al MINIMO de la serie (aunque vaya desordenada)
t_d, segs_d = app.segmentos_temporales([("2026-06-20", 71.0),
                                        ("2026-06-01", 70.0)])
check("B95 dias relativos al minimo (orden no importa)",
      t_d and [x for x, _v, _e in sorted(segs_d[0])] == [0.0, 19.0],
      str(segs_d))
check("B95 vacia -> (True, [])",
      app.segmentos_temporales([]) == (True, []))
check("B95 umbral configurable se respeta",
      app.segmentos_temporales([("2026-06-01", 1.0),
                                ("2026-06-20", 2.0)], umbral_dias=10)[1].__len__() == 2)

# --- B51: seriales de Excel (lector del Progreso Quimico, C-17) ---
check("serial Excel 46000 -> 2025-12-09", app._serial_excel_a_iso(46000) == "2025-12-09",
      app._serial_excel_a_iso(46000))
check("serial Excel 1 -> 1899-12-31 (epoch; exacto desde serial 61)",
      app._serial_excel_a_iso(1) == "1899-12-31", app._serial_excel_a_iso(1))
check("serial Excel 61 -> 1900-03-01 (epoch exacto post-bisiesto falso)",
      app._serial_excel_a_iso(61) == "1900-03-01", app._serial_excel_a_iso(61))
check("serial fuera de rango -> None", app._serial_excel_a_iso(0) is None)
check("no numerico -> None", app._serial_excel_a_iso("abc") is None)

# --- F0/M2: episodio_contexto (fallback del panel de Inicio y tira de Docs) ---
check("episodio_contexto usa el activo presente",
      app.episodio_contexto({"episodio_activo": "E01_2020",
                             "episodios": [{"id": "E01_2020"}, {"id": "E02_2021"}]}
                            )["id"] == "E01_2020")
check("episodio_contexto cae al ultimo si el activo falta (F0/M2)",
      app.episodio_contexto({"episodio_activo": "E09_2099",
                             "episodios": [{"id": "E01_2020"}, {"id": "E02_2021"}]}
                            )["id"] == "E02_2021")
check("episodio_contexto sin episodios -> None",
      app.episodio_contexto({"episodios": []}) is None)

# --- B3/E4: crear_episodio ---
datos = {"episodios": []}
eid = app.crear_episodio(datos)
anio_actual = _dt.date.today().year
check("crear_episodio retrocompatible (sin args -> anio actual)",
      eid == f"E01_{anio_actual}", eid)
check("crear_episodio sin args NO agrega campos nuevos",
      "fecha_referencia" not in datos["episodios"][0] and "fum" not in datos["episodios"][0])

datos2 = {"episodios": []}
eid2 = app.crear_episodio(datos2, anio=1999, fecha_ref="2024-05-15", fum="2023-08-08")
check("D4: fecha_ref tiene precedencia sobre anio", eid2 == "E01_2024", eid2)
check("B3: campos nuevos persistidos en el dict",
      datos2["episodios"][0].get("fecha_referencia") == "2024-05-15"
      and datos2["episodios"][0].get("fum") == "2023-08-08")
eid3 = app.crear_episodio(datos2, fecha_ref="no-valida")
check("B3: fecha_ref invalida cae a year actual (no truenca)",
      eid3 == f"E02_{anio_actual}", eid3)

# --- E6: actualizar_episodio (headless sobre un paciente real) ---
ok, msg, ruta, d = app.crear_paciente("MEDINA", "ROJAS", "CLAUDIA", "1990-07-20", "")
check("Paciente base para episodios", ok, msg)
eid = d["episodio_activo"]
ok1 = app.actualizar_episodio(ruta, eid, {"fum": "2026-06-10",
                                          "fecha_referencia": "2027-03-17"})
d2 = app.cargar_paciente(ruta)
epi = [e for e in d2["episodios"] if e["id"] == eid][0]
check("actualizar_episodio guarda fum y fecha_referencia",
      ok1 and epi.get("fum") == "2026-06-10"
      and epi.get("fecha_referencia") == "2027-03-17")
ok2 = app.actualizar_episodio(ruta, eid, {"fum": None})
d3 = app.cargar_paciente(ruta)
epi3 = [e for e in d3["episodios"] if e["id"] == eid][0]
check("actualizar_episodio(None) limpia la clave (Limpiar FUM)",
      ok2 and "fum" not in epi3 and epi3.get("fecha_referencia") == "2027-03-17")
check("actualizar_episodio eid inexistente -> False",
      app.actualizar_episodio(ruta, "E99_NOPE", {"fum": "2020-01-01"}) is False)

# --- B3: nuevo_episodio con fecha_ref (carpeta y id) ---
eido = app.nuevo_episodio(ruta, fecha_ref="2024-05-15", fum="2023-08-08")
check("nuevo_episodio con fecha_ref crea E02_2024",
      eido == "E02_2024", eido)
check("carpeta del episodio creada", (Path(ruta) / "E02_2024").is_dir())

# --- E7: mapa_marcadores con y sin FUM ---
app.fijar_episodio(ruta, eid)  # episodio inicial sin FUM
d4 = app.cargar_paciente(ruta)
m = app.mapa_marcadores(d4)
check("marcadores vacios sin FUM",
      m["{{FUM}}"] == "" and m["{{FPP}}"] == "" and m["{{EG}}"] == "")
app.fijar_episodio(ruta, eido)  # episodio con FUM 2023-08-08
d5 = app.cargar_paciente(ruta)
m2 = app.mapa_marcadores(d5)
check("{{FUM}} legible", m2["{{FUM}}"] == "08-AGO-2023", m2["{{FUM}}"])
check("{{FPP}} calculada", m2["{{FPP}}"] == "15-MAY-2024", m2["{{FPP}}"])
check("{{EG}} con FUM pasada", m2["{{EG}}"] and m2["{{EG}}"].endswith("d"),
      m2["{{EG}}"])

# --- Fijar episodio y progreso siguen en la carpeta correcta tras cambios ---
check("episodio activo coincide tras actualizar",
      d5.get("episodio_activo") == eido, d5.get("episodio_activo"))

print()
print("=" * 60)
if FALLAS:
    print(f"RESULTADO: {len(FALLAS)} FALLAS")
    for f in FALLAS:
        print(f"  - {f}")
    sys.exit(1)
print("RESULTADO: TODAS LAS PRUEBAS DE CALENDARIO PASARON")
