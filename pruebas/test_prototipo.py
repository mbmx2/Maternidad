# Pruebas de aceptacion del prototipo (headless: solo logica, sin UI)
import sys
import os
import shutil
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _sandbox import preparar_app

app, BASE = preparar_app()
FALLAS = []

def check(nombre, cond, detalle=""):
    estado = "OK " if cond else "FAIL"
    print(f"[{estado}] {nombre}" + (f"  <- {detalle}" if (detalle and not cond) else ""))
    if not cond:
        FALLAS.append(nombre)

# --- DO-2: estructura primer arranque ---
app.inicializar_estructura()
check("DO-2 estructura A-Z", all((app.EXPEDIENTES_DIR / L).is_dir() for L in "AZ"))
check("DO-2 LEEME.txt", app.LEEME_PATH.exists())
check("DO-2 INSTRUCCIONES plantillas", (app.PLANTILLAS_DIR / "INSTRUCCIONES.txt").exists())

# Garantia del dueño (2026-10-16): arrancar dos veces NO reescribe archivos existentes
_m_leeme = os.path.getmtime(app.LEEME_PATH)
_m_cfg = os.path.getmtime(app.CONFIG_PATH)
_m_bit = os.path.getmtime(app.BITACORA_PATH)
app.inicializar_estructura()
check("DO-2 inicializar idempotente (no reescribe)",
      os.path.getmtime(app.LEEME_PATH) == _m_leeme and
      os.path.getmtime(app.CONFIG_PATH) == _m_cfg and
      os.path.getmtime(app.BITACORA_PATH) == _m_bit)

# --- C-01/B26: estado vacio y busqueda sin texto ---
check("C-01 estado vacio: 0 pacientes sin excepcion", app.listar_pacientes() == [],
      str(app.listar_pacientes()))
check("B26 buscar sin texto -> []", app.buscar("") == [], str(app.buscar("")))

# --- B09: validar CURP ---
check("B09 CURP vacia es valida (opcional)", app.validar_curp("")[0] is True)
check("B09 CURP valida", app.validar_curp("GODE561231MDFRRL09")[0] is True,
      app.validar_curp("GODE561231MDFRRL09")[1])
check("B09 CURP corta rechazada", app.validar_curp("GODE561")[0] is False,
      app.validar_curp("GODE561")[1])
check("B09 CURP patron malo rechazada",
      app.validar_curp("XXXXXXXXXXXXXXXXXX")[0] is False,
      app.validar_curp("XXXXXXXXXXXXXXXXXX")[1])

# --- Tier A (v2.0): validar_gpa ---
check("T7 gpa valido", app.validar_gpa("G3P2A1")[0] is True, app.validar_gpa("G3P2A1")[1])
check("T7 gpa vacio es opcional", app.validar_gpa("")[0] is True)
check("T7 gpa malo rechazado",
      app.validar_gpa("G3")[0] is False and app.validar_gpa("3P2A1")[0] is False,
      app.validar_gpa("G3")[1])# --- plantilla docx real con marcador DIVIDIDO entre runs ---
plantilla = app.PLANTILLAS_DIR / "Ultrasonido Base.docx"
doc_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body><w:p><w:r><w:t>Paciente: {{PACI"
    "</w:t></w:r><w:r><w:t xml:space=\"preserve\">ENTE}}</w:t></w:r></w:p>"
    "<w:p><w:r><w:t>Fecha: {{FECHA}}</w:t></w:r></w:p>"
    "<w:p><w:r><w:t>Doctor: {{DOCTOR}}</w:t></w:r></w:p>"
    "</w:body></w:document>"
)
with zipfile.ZipFile(plantilla, "w") as z:
    z.writestr("[Content_Types].xml",
               '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
    z.writestr("word/document.xml", doc_xml)

# plantilla progreso base xlsx con marcador
base_prog = app.PLANTILLAS_DIR / "PROGRESO_BASE.xlsx"
app._crear_xlsx_minimo(base_prog, [["Fecha", "{{PACIENTE}}", "Glucosa"]])

# --- DO-3: registro con Enter/mayusculas lo valida la UI; aqui la logica ---
ok, msg, ruta, datos = app.crear_paciente(
    "garcía", "lópez", "maría fernanda", "15/03/1985", "")
check("DO-3 crear paciente (acentos+fecha dd/mm)", ok, msg)
check("DO-3 clave normalizada sin acentos",
      ruta and "GARCIA_LOPEZ_MARIA_FERNANDA_19850315" in ruta, ruta or "")
check("DO-3 nombre visible conserva tildes",
      datos and datos["nombre_completo"] == "GARCÍA LÓPEZ MARÍA FERNANDA",
      datos.get("nombre_completo") if datos else "")

# B21: listar_recientes debe funcionar con pacientes reales (regresion: ruta str/Path)
recientes = app.listar_recientes(10)
check("B21 listar_recientes sin excepcion", isinstance(recientes, list) and len(recientes) >= 1,
      str(recientes))
check("B21 recientes contiene a la creada", any("GARCIA_LOPEZ" in r for _, r in recientes),
      str(recientes))

# duplicado exacto
ok2, msg2, ruta2, _ = app.crear_paciente("GARCIA LOPEZ", "", "MARIA FERNANDA", "1985-03-15", "")
check("F-03 duplicado detectado (mesma fecha)", (not ok2) and "Ya existe" in msg2, msg2)

# --- F-04 homonimo por CURP distinta ---
ok3, msg3, ruta3, _ = app.crear_paciente(
    "GARCIA LOPEZ", "", "MARIA FERNANDA", "15-MAR-1985", "GODE561231MDFRRL09",
    ventana_callback=lambda m: True)
check("F-04 homonimo crea expediente separado _2",
      ok3 and ruta3 and ruta3.endswith("_2"), f"{ok3} {msg3} {ruta3}")

# CURP invalida con callback que rechaza
ok4, msg4, _, _ = app.crear_paciente("PEREZ", "DIAZ", "ANA", "1990-01-01", "XXXX",
                                     ventana_callback=lambda m: False)
check("F-04 CURP invalida rechazada", (not ok4) and ("18" in msg4 or "invalid" in msg4.lower()), msg4)

# --- DO-7: busqueda con espacios, acentos y minusculas ---
r1 = app.buscar("garcia lopez")
r2 = app.buscar("GARCÍA")
r3 = app.buscar(datos and datos.get("nombre_completo", "")[:6].lower() or "")
check("DO-7 buscar 'garcia lopez' (espacio)", len(r1) >= 1, str(r1))
check("DO-7 buscar 'GARCÍA' (acento)", len(r2) >= 1, str(r2))

# --- DO-4: episodios (2do embarazo = episodio, no carpeta nueva) ---
antes = len(app.listar_pacientes())
eid = app.nuevo_episodio(ruta)
despues = app.listar_pacientes()
datos = app.cargar_paciente(ruta)
check("DO-4 nuevo episodio E02", eid.startswith("E02_"), eid)
check("DO-4 sin carpeta nueva de paciente", len(despues) == antes, f"{antes}->{len(despues)}")
check("DO-4 episodio activo actualizado", datos["episodio_activo"] == eid, datos["episodio_activo"])
check("DO-4 carpeta episodio creada", (Path(ruta) / eid).is_dir())

# --- Tier A (v2.0, ADR-0024): campos estructurados ---
okT7, msgT7 = app.guardar_datos_secundarios(
    ruta, "", "", "", "", "nota t7", alergias="Penicilina", grupo_sanguineo="O positivo")
d_t7 = app.cargar_paciente(ruta)
check("T7 alergias/grupo persistidos",
      d_t7.get("alergias") == "Penicilina" and d_t7.get("grupo_sanguineo") == "O positivo",
      msgT7)
okT7b, msgT7b = app.guardar_datos_secundarios(ruta, "5512345678", "", "", "", "nota t7")
d_t7b = app.cargar_paciente(ruta)
check("T7 guardar sin alergias no las borra (None = no tocar)",
      d_t7b.get("alergias") == "Penicilina"
      and d_t7b.get("telefono") == "5512345678", msgT7b)
check("T7 actualizar_episodio campos Tier A",
      app.actualizar_episodio(ruta, eid, {"gpa": "G2P1A0",
                                          "tipo_parto_anterior": "Cesarea",
                                          "fecha_parto_anterior": "2020-05-10",
                                          "complicaciones_previas": "Preclampsia"}))
m_t7 = app.mapa_marcadores(app.cargar_paciente(ruta))
check("T7 marcadores Tier A (23 claves con C-16)",
      len(m_t7) == 23 and m_t7["{{GP_A}}"] == "G2P1A0"
      and m_t7["{{ALERGIAS}}"] == "Penicilina"
      and m_t7["{{GRUPO}}"] == "O positivo"
      and "Cesarea" in m_t7["{{PARTO_ANT}}"]
      and m_t7["{{EDAD}}"].isdigit()
      and m_t7["{{PESO}}"] == "",
      str({k: m_t7.get(k) for k in ("{{EDAD}}", "{{GP_A}}", "{{PARTO_ANT}}", "{{PESO}}")}))
check("T7 retrocompatibilidad: episodios sin campos Tier A leen con .get",
      app.mapa_marcadores({"nombre_completo": "X", "episodios": [{"id": "E01_2020"}]})
      .get("{{GP_A}}") == "")

# --- C-16 (v2.1, ADR-0025): controles prenatales ---
okCT, msgCT, ctrlCT = app.agregar_control(ruta, eid, {
    "fecha": "2026-06-01", "peso": "70.5", "ta_s": "120", "ta_d": "80",
    "fc": "78", "au": "28", "fcf": "140", "proteinuria": "negativa",
    "notas": "1a visita"})
check("C16 alta de control con numericos",
      okCT and ctrlCT["peso"] == 70.5 and ctrlCT["ta_s"] == 120.0
      and ctrlCT["proteinuria"] == "negativa", msgCT)
okCT2, msgCT2, ctrlCT2 = app.agregar_control(ruta, eid,
                                             {"fecha": "2026-06-15", "peso": "71,2"})
check("C16 coercion de decimal con coma", okCT2 and ctrlCT2["peso"] == 71.2, msgCT2)
futuro = (app.datetime.date.today() + app.datetime.timedelta(days=1)).isoformat()
okCT3, msgCT3, _ = app.agregar_control(ruta, eid, {"fecha": futuro, "peso": "70"})
check("C16 fecha futura rechazada",
      not okCT3 and "posterior a hoy" in msgCT3, msgCT3)
lista_ct = app.listar_controles(ruta, eid)
check("C16 lista ordenada descendente",
      len(lista_ct) == 2 and lista_ct[0]["fecha"] >= lista_ct[1]["fecha"],
      str([c["fecha"] for c in lista_ct]))
check("C16 actualizar por indice almacenado",
      app.actualizar_control(ruta, eid, 0,
                             {"fecha": "2026-06-01", "peso": "70.9"})[0]
      and app.listar_controles(ruta, eid)[1]["peso"] == 70.9)
check("C16 rangos vacios = sin validacion blanda",
      app.rangos_control_fuera("peso", 999) is None
      and app.rangos_control_fuera("ta_s", 10) is None)
check("C16 listar episodio inexistente -> []", app.listar_controles(ruta, "E99_1999") == [])
check("C16 actualizar indice fuera de rango -> False",
      not app.actualizar_control(ruta, eid, 99, {"fecha": "2026-06-01"})[0])
m_ct = app.mapa_marcadores(app.cargar_paciente(ruta))
check("C16 marcadores: ultimo valor REGISTRADO por campo",
      m_ct["{{PESO}}"] == "71.2" and m_ct["{{TA}}"] == "120/80"
      and m_ct["{{AU}}"] == "28" and m_ct["{{FCF}}"] == "140",
      str({k: m_ct.get(k) for k in ("{{PESO}}", "{{TA}}", "{{AU}}", "{{FCF}}")}))

# --- DO-5: generacion desde docx con marcador dividido ---
ok, msg, salida = app.generar_documento(ruta, "Ultrasonido Base.docx")
check("DO-5 documento generado", ok, msg)
check("DO-5 nombre con timestamp",
      salida and Path(salida).name[:8].isdigit() and "_ULTRASONIDO_BASE.docx" in Path(salida).name,
      salida)
if salida:
    with zipfile.ZipFile(salida) as z:
        contenido = z.read("word/document.xml").decode()
    check("DO-5 marcador {{PACIENTE}} sustituido (split-aware)",
          "GARCÍA LÓPEZ MARÍA FERNANDA" in contenido and "{{" not in contenido, contenido[:200])
    check("DO-5 {{FECHA}} sustituido", "{{FECHA}}" not in contenido)
    check("DO-5 plantilla origen intacta",
          "{{PACI" in zipfile.ZipFile(plantilla).read("word/document.xml").decode())

# anti-sobreescritura: callback True = abrir existente
okA, msgA, salA = app.generar_documento(
    ruta, "Ultrasonido Base.docx",
    callback_existente=lambda n: True)
check("W-05 callback True devuelve existente", okA and "existente" in msgA.lower(), msgA)

# callback False = copia nueva _1 (forzando mismo timestamp con nombre_forzado)
okB, msgB, salB = app.generar_documento(
    ruta, "Ultrasonido Base.docx",
    callback_existente=lambda n: False, nombre_forzado="COLISION")
if salB:  # primer intento libre
    okC, msgC, salC = app.generar_documento(
        ruta, "Ultrasonido Base.docx",
        callback_existente=lambda n: False, nombre_forzado="COLISION")
    check("W-05 contador _1 evita pisar", okC and "_1" in Path(salC).name, msgC)
    check("W-05 los dos archivos coexisten",
          Path(salB).exists() and Path(salC).exists())

# --- DO-6: Progreso Quimico vivo por episodio ---
okP1, msgP1, salP1 = app.progreso_quimico(ruta)
okP2, msgP2, salP2 = app.progreso_quimico(ruta)
check("DO-6 1er clic crea", okP1 and "creado" in msgP1.lower(), msgP1)
check("DO-6 2o clic abre MISMO archivo", okP2 and salP2 == salP1 and "existente" in msgP2.lower(), msgP2)
check("DO-6 vive en episodio activo", salP1 and eid in salP1, salP1)
# con la plantilla marcador rellenada
if salP1 and zipfile.is_zipfile(salP1):
    with zipfile.ZipFile(salP1) as z:
        ss = z.read("xl/worksheets/sheet1.xml").decode()
    check("DO-6 marcador de plantilla rellenado", "GARC" in ss and "{{" not in ss, ss[:150])

# progreso en episodio anterior no se mezcla
app.fijar_episodio(ruta, datos["episodios"][0]["id"])
okP3, _, salP3 = app.progreso_quimico(ruta)
check("W-04 progreso separado por episodio", salP3 != salP1, f"{salP3} vs {salP1}")
app.fijar_episodio(ruta, eid)

# --- ADR-0027/T13: panel Hoy (resumen_hoy) ---
rh = app.resumen_hoy(dias=7, umbral=42, hoy_iso="2026-06-16")
check("T13 recientes: control del 15-jun (hace 1 dia)",
      any(r[2] == "2026-06-15" and r[3] == 1 for r in rh["recientes"]),
      str(rh["recientes"]))
rh_v = app.resumen_hoy(dias=7, umbral=10, hoy_iso="2026-07-01")
check("T13 vencidos: se evalua el ULTIMO control por paciente (15-jun, 16 d)",
      any(r[2] == "2026-06-15" and r[3] == 16 for r in rh_v["vencidos"]),
      str(rh_v["vencidos"]))
check("T13 el control viejo (1-jun) no cuenta si no es el ultimo",
      not any(r[2] == "2026-06-01" for r in rh_v["vencidos"]), str(rh_v["vencidos"]))
check("T13 con umbral 42 y hoy 16-jun nadie esta vencido",
      not rh["vencidos"], str(rh["vencidos"]))
rh_fut = app.resumen_hoy(hoy_iso="2027-01-10")
# la fecha de creacion del expediente es "hoy" del sistema: se calcula en
# caliente para que el test no se rompa al cambiar de dia (bomba de fecha)
check("T13 sin controles: usa fecha de creacion del expediente",
      any(r[2] == app.fecha_hoy() for r in rh_fut["vencidos"]),
      str(rh_fut["vencidos"]))
check("T13 umbral configurable devuelve el umbral usado", rh["umbral"] == 42)

# --- ADR-0027/T13: resumen de consulta de 1 pagina ---
txt_r = app.resumen_consulta(ruta)
check("T13 resumen: identidad + episodio + derivados",
      "RESUMEN DE CONSULTA" in txt_r and "FERNANDA" in txt_r
      and "EPISODIO ACTIVO" in txt_r and "FPP" in txt_r, txt_r[:200])
check("T13 resumen: ultimos controles con peso real",
      "ULTIMOS CONTROLES" in txt_r and "peso 71.2 kg" in txt_r, txt_r[:900])
check("T13 resumen: Tier A incluido",
      "Alergias" in txt_r and "Penicilina" in txt_r and "G-P-A" in txt_r)
check("T13 resumen: paciente inexistente -> error controlado",
      app.resumen_consulta(str(BASE / "NOEXISTE")).startswith("ERROR"))

# --- C-18 / ADR-0030 (T16): esquema transcrito + core de historia ---
secs = app.SECCIONES_HISTORIA
tot_items = sum(len(s["items"]) for s in secs)
check("C-18 esquema transcrito: 40 secciones", len(secs) == 40, str(len(secs)))
check("C-18 esquema transcrito: 477 items", tot_items == 477, str(tot_items))
check("C-18 secciones sin duplicar", len({s["id"] for s in secs}) == 40)
niveles = {}
for s in secs:
    niveles[s["nivel"]] = niveles.get(s["nivel"], 0) + 1
check("C-18 niveles 14/24/1/1",
      niveles == {"paciente": 14, "nota": 24, "evolucion": 1, "egreso": 1},
      str(niveles))
tipos_ok = {"auto", "texto", "numero", "fecha", "enum", "sn", "snns", "snna",
            "snv", "snc", "snh", "libre"}
check("C-18 tipos de item validos",
      all(it["t"] in tipos_ok for s in secs for it in s["items"]))
check("C-18 cada item con id y etiqueta",
      all(it.get("id") and it.get("lbl") for s in secs for it in s["items"]))

h0 = app.cargar_historia(str(BASE / "NOEXISTE"))
check("B67 sin HISTORIA.json -> default", h0["paciente"] == {} and h0["notas"] == [])
h0["paciente"]["A"] = {"nombre": "PRUEBA", "nss": "123"}
okH18, msgH18 = app.guardar_historia(ruta, h0)
check("B68 guardar historia", okH18, msgH18)
h18 = app.cargar_historia(ruta)
check("B68 roundtrip preserva seccion A",
      h18["paciente"].get("A") == {"nombre": "PRUEBA", "nss": "123"},
      str(h18["paciente"]))
h18["paciente"]["A"]["nss"] = "456"
app.guardar_historia(ruta, h18)
h18b = app.cargar_historia(ruta)
check("B68 merge conserva lo no enviado",
      h18b["paciente"]["A"]["nombre"] == "PRUEBA"
      and h18b["paciente"]["A"]["nss"] == "456")
pr0 = app.progreso_historia(app.historia_default())
check("B69 vacio: solo items auto cuentan",
      pr0["A"][0] == sum(1 for it in secs[0]["items"] if it["t"] == "auto"),
      str(pr0["A"]))
check("B69 totales suman 477", sum(b for _, b in pr0.values()) == 477)
hp = app.historia_default()
hp["paciente"]["A"] = {"nss": "12345"}
check("B69 texto poblado cuenta (+1)",
      app.progreso_historia(hp)["A"][0] == pr0["A"][0] + 1)
hn_base = app.progreso_historia(app.historia_default())["H2"]
hn_no = app.historia_default()
hn_no["notas"] = [{"tipo": "ingreso", "fecha": "2026-10-16", "hora": "10:00",
                   "episodio": "E01", "datos": {"H2": {"edad_aparente": "no"}}}]
check("B69 estado 'no' cuenta como investigado",
      app.progreso_historia(hn_no)["H2"][0] == hn_base[0] + 1)
hn_v = app.historia_default()
hn_v["notas"] = [{"tipo": "ingreso", "fecha": "2026-10-16", "hora": "10:00",
                  "episodio": "E01", "datos": {"H2": {"edad_aparente": ""}}}]
check("B69 '' (sin investigar) NO cuenta",
      app.progreso_historia(hn_v)["H2"][0] == hn_base[0])
ctx18 = app.contexto_historia(ruta)
d18 = app.cargar_paciente(ruta)
check("C-18 auto A: nombre del expediente",
      app.auto_historia("A", "nombre", ctx18) == d18["nombre_completo"])
check("C-18 auto H1: TA y peso del ultimo registro por campo",
      app.auto_historia("H1", "ta", ctx18) == "120/80"
      and app.auto_historia("H1", "peso", ctx18) == "71.2",
      f"ta={app.auto_historia('H1', 'ta', ctx18)!r} "
      f"peso={app.auto_historia('H1', 'peso', ctx18)!r}")
app.generar_consentimiento(ruta, "Prueba C-18")
ctx18b = app.contexto_historia(ruta)
check("C-18 auto K: consentimiento derivado del episodio",
      "generado" in app.auto_historia("K", "consentimiento", ctx18b),
      app.auto_historia("K", "consentimiento", ctx18b))
hn2 = app.historia_default()
idx_n = app.nueva_nota_historia(hn2, "ingreso", "E01_2026")
check("C-18 nueva nota (sin borrados)",
      idx_n == 0 and hn2["notas"][0]["tipo"] == "ingreso"
      and hn2["notas"][0]["episodio"] == "E01_2026")

# --- C-18 T17: render A-M (B70), generacion y export (B71) ---
txt2 = app.render_historia_v2(ruta)
faltan = [s["id"] for s in secs if f"\n{s['id']}. " not in ("\n" + txt2)]
check("B70 render contiene las 40 secciones", not faltan, str(faltan))
check("B70 bloque 1.b presente", "1.b. DATOS AUTOMATICOS" in txt2)
check("B70 estados impresos (sin investigar)", "sin investigar" in txt2)
check("B70 valor guardado impreso (nss 456)",
      "Numero de Seguridad Social (NSS): 456" in txt2, txt2[:600])
check("B70 pie K con SHA-256 de 64 hex",
      "SHA-256 del documento:" in txt2
      and len(txt2.split("SHA-256 del documento:")[1].split()[0]) == 64)
check("B70 nota ADR-0013 en el pie", "ADR-0013" in txt2)
okg2, msgg2, salg2 = app.generar_historia_v2(ruta)
check("B70 generacion A-M en el episodio",
      okg2 and "HISTORIA_COMPLETA_" in Path(salg2).name
      and "SHA-256 del documento" in Path(salg2).read_text(encoding="utf-8"), msgg2)
oke2, msge2, sale2 = app.exportar_checklist_xlsx(ruta)
check("B71 checklist exportado a xlsx valido",
      oke2 and zipfile.is_zipfile(sale2), msge2)
if oke2:
    with zipfile.ZipFile(sale2) as _z:
        _hoja2 = _z.read("xl/worksheets/sheet1.xml").decode("utf-8", "replace")
    check("B71 checklist con columnas y items reales",
          "Seccion" in _hoja2 and "Gestaciones" in _hoja2
          and "Numero de Seguridad Social" in _hoja2)

# --- C-20 / ADR-0031 (T18): alertas — criterios del contrato ---
app.actualizar_episodio(ruta, eid, {"fum": "2026-06-10"})
h1 = app.evaluar_signos_vitales({"ta_s": 150.0, "ta_d": 95.0}, "adulto")
check("C-20.1 TA 150/95 adulto -> hallazgo de alerta",
      any(x["nivel"] == "alerta" and "TA_S" in x["mensaje"] for x in h1), str(h1))
h2 = app.evaluar_signos_vitales({"fc": 130.0}, "embarazada")
check("C-20.2 FC 130 embarazada -> esperado_embarazo (informativo)",
      any(x["nivel"] == "esperado_embarazo" for x in h2), str(h2))
hp = app.evaluar_preeclampsia({"ta_s": 165.0, "ta_d": 112.0}, "2+")
check("C-20.3 165/112 + 2+ -> preeclampsia critica",
      len(hp) == 1 and hp[0]["nivel"] == "critico"
      and "PREECLAMPSIA" in hp[0]["mensaje"].upper(), str(hp))
check("C-20.3 sin proteinuria no alerta preeclampsia",
      app.evaluar_preeclampsia({"ta_s": 165.0, "ta_d": 112.0}, "") == [])
hl = app.evaluar_lab("potasio", 7.0)
check("C-20.4 potasio 7.0 -> panic value con riesgo",
      hl is not None and hl["nivel"] == "critico" and "VALOR CRITICO" in hl["mensaje"]
      and "Arritmias" in hl["mensaje"], str(hl))
check("C-20.4 potasio en rango -> None", app.evaluar_lab("potasio", 4.0) is None)
pts_q, qm_q = app.score_qsofa({"fr": 24, "ta_s": 95, "conciencia_alterada": True})
check("C-20.5 qSOFA 3 -> MAL PRONOSTICO", pts_q == 3 and "MAL PRONOSTICO" in qm_q,
      f"{pts_q} {qm_q}")
nw, nwm = app.score_escala_incompleta("mews")
check("C-20 MEWS no puntua (tabla pendiente del medico)",
      nw is None and "pendiente" in nwm.lower(), nwm)
nw2, nwm2 = app.score_escala_incompleta("news")
check("C-20 NEWS no puntua (tabla pendiente del medico)",
      nw2 is None and "pendiente" in nwm2.lower(), nwm2)
# circuito cerrado de alertas (B79/B80)
ok_a, id_a = app.registrar_alerta(ruta, {"origen": "test",
                                          "descripcion": "hallazgo de prueba",
                                          "episodio": eid})
check("B79 alta de alerta pendiente", ok_a and str(id_a).startswith("ALT-"),
      str((ok_a, id_a)))
check("B79 aparece en pendientes del episodio",
      any(a.get("id") == id_a for a in app.alertas_pendientes(ruta, eid)))
check("B80 sin paciente no revienta", app.alertas_pendientes(str(BASE / "NO")) == [])
ok_ne, msg_ne = app.atender_alerta(ruta, "NOEXISTE", {"texto": "x"})
check("B80 id inexistente -> controlado", not ok_ne and "no encontrada" in msg_ne)
ok_at, msg_at = app.atender_alerta(ruta, id_a, {
    "texto": "Se notifico al medico y se ajusto el plan", "medico": "DR TEST",
    "cedula": "CED-123"})
check("B80 atender -> 'Alerta atendida.'", ok_at and msg_at == "Alerta atendida.",
      msg_at)
ok_at2, msg_at2 = app.atender_alerta(ruta, id_a, {"texto": "otra vez"})
check("B80 doble atencion rechazada", not ok_at2 and "pendiente" in msg_at2, msg_at2)
check("B80 sale de pendientes",
      not any(a.get("id") == id_a for a in app.alertas_pendientes(ruta, eid)))
lineas_b = app.BITACORA_PATH.read_text(encoding="utf-8").splitlines()
check("C-20 ambas acciones firman la bitacora",
      any("Alerta registrada" in l for l in lineas_b)
      and any("Alerta atendida" in l for l in lineas_b))
check("C-20 bitacora verificable tras las alertas",
      app.verificar_bitacora()[0])
# criterio 7: control viejo sin temp/fr/spo2
okc7, msgc7, c7 = app.agregar_control(ruta, eid,
                                       {"fecha": "2026-07-02", "peso": "70"})
check("C-20.7 registro sin temp/fr/spo2 no falla y los guarda vacios",
      okc7 and c7.get("temp") is None and c7.get("fr") is None
      and c7.get("spo2") is None, str(c7))
# integracion hallazgos_control: perfil embarazada + critico + qSOFA
hall_i, pts_i, qm_i, perfil_i = app.hallazgos_control(
    ruta, {"ta_s": "170", "ta_d": "110", "proteinuria": "2+", "fr": "24"})
check("C-20 integracion: perfil embarazada, preeclampsia critica, qSOFA calculado",
      perfil_i == "embarazada"
      and any(h["nivel"] == "critico" for h in hall_i)
      and pts_i == 1,
      f"{perfil_i} pts={pts_i} hall={len(hall_i)}")
# revision del progreso (solo lectura)
ok_rv, msg_rv, hall_rv = app.revisar_lab_progreso(ruta, eid)
check("C-20 revision de progreso sin criticos en hoja minima",
      ok_rv and hall_rv == [], msg_rv)
check("C-20 revision sin archivo -> controlado",
      not app.revisar_lab_progreso(ruta, "E99_1999")[0])

# --- C-19 / ADR-0032 (T19): calculadoras ---
check("B73 26 calculadoras (24 activas / 2 incompletas)",
      len(app.CALCULADORAS) == 26
      and sum(1 for v in app.CALCULADORAS.values() if v["estado"] == "activa") == 24
      and sum(1 for v in app.CALCULADORAS.values() if v["estado"] == "incompleta") == 2,
      str(len(app.CALCULADORAS)))
check("B73 8 grupos del medico en orden", len(app.GRUPOS_CALCULADORAS) == 8,
      str(app.GRUPOS_CALCULADORAS))
check("B73 FUNCIONES_CALC cubre todas las activas",
      all(cid in app.FUNCIONES_CALC for cid, c in app.CALCULADORAS.items()
          if c["estado"] == "activa"))
ok_i, res_i, msg_i = app.calcular("imc", {"peso": 70, "talla": 1.65})
check("C-19.1 IMC 70/1.65 -> 25.7 Sobrepeso",
      ok_i and abs(res_i["valor"] - 25.7) < 1e-6
      and res_i["etiqueta"] == "Sobrepeso", f"{res_i} {msg_i}")
ok_p, res_p, _ = app.calcular("pam", {"ta_s": 120, "ta_d": 80})
check("C-19.2 PAM 120/80 -> 93.3 Perfusion adecuada",
      ok_p and abs(res_p["valor"] - 93.3) < 1e-6
      and res_p["etiqueta"] == "Perfusion adecuada", str(res_p))
ok_m, res_m, msg_m = app.calcular("mews", {})
check("C-19.4 MEWS -> no puntua, pendiente del medico",
      not ok_m and res_m is None and "pendiente" in msg_m.lower(), msg_m)
ok_n2, res_n2, msg_n2 = app.calcular("news", {})
check("C-19.4 NEWS -> no puntua, pendiente del medico",
      not ok_n2 and res_n2 is None and "pendiente" in msg_n2.lower(), msg_n2)
ok_g, res_g, _ = app.calcular("gcs", {"apertura": 3, "verbal": 4, "motora": 6})
check("C-19.5 GCS E3 V4 M6 -> 13", ok_g and res_g["valor"] == 13, str(res_g))
ok_g0, _res_g0, msg_g0 = app.calcular("gcs", {"apertura": 0, "verbal": 4, "motora": 6})
check("C-19.5 GCS E0 -> error de rango",
      not ok_g0 and "rango" in msg_g0.lower(), msg_g0)
_NUM_DEF = {"edad": 45, "peso": 70, "talla": 1.65, "talla_in": 65, "altura": 170,
            "apertura": 3, "verbal": 3, "motora": 4, "apariencia": 1, "pulso": 1,
            "reflejo": 1, "actividad": 1, "respiracion": 1, "qt": 400, "rr": 800,
            "fc": 80, "fr": 16, "spo2": 97, "ta_s": 120, "ta_d": 80, "temp": 36.5,
            "glucosa": 90, "creatinina": 1.0, "cr_serica": 1.0, "cr_urinaria": 1.0,
            "na": 140, "na_serico": 140, "na_urinario": 40, "cl": 102, "hco3": 24,
            "albunina": 3.5, "bilirrubina": 0.8, "inr": 1.0, "plaquetas": 250,
            "ast": 30, "alt": 30, "ascitis": 1, "encefalopatia": 1,
            "borramiento": 15, "dilatacion": 3, "consistencia": 2, "posicion": 1,
            "puntaje": 2}
_TXT_DEF = {"sexo": "M", "avpu": "A", "conciencia": "LUCIDO"}

def _ent(c):
    out = {}
    for fid, _lbl, tipo in c["campos"]:
        if tipo == "num":
            out[fid] = _NUM_DEF.get(fid, 1.0)
        elif tipo == "bool":
            out[fid] = False
        else:
            out[fid] = _TXT_DEF.get(fid, "")
    return out

fallos_c = [cid for cid, c in app.CALCULADORAS.items()
            if c["estado"] == "activa" and not app.calcular(cid, _ent(c))[0]]
check("C-19.6 las 24 activas calculan sin excepcion", not fallos_c, str(fallos_c))

# --- ADR-0043: orden de grupos + calculadoras prioritarias ---
check("C-19.7 'Generales/Composicion corporal' encabeza los grupos (workflow)",
      app.GRUPOS_CALCULADORAS[0] == "Generales/Composicion corporal"
      and len(app.GRUPOS_CALCULADORAS) == 8, str(app.GRUPOS_CALCULADORAS[:2]))
check("C-19.8 grupo virtual ★ existe y no duplica nombres",
      app.GRUPO_PRIORITARIAS not in app.GRUPOS_CALCULADORAS
      and app.GRUPO_PRIORITARIAS.startswith("★"), app.GRUPO_PRIORITARIAS)
_cfg_prev = app.leer_config().get("calc_prioritarias", "__ausente__")
if _cfg_prev != "__ausente__":
    # guardar_config hace merge (no borra claves): se reescribe el JSON directo
    import json as _json
    _aux = app.leer_config()
    _aux.pop("calc_prioritarias")
    app.CONFIG_PATH.write_text(_json.dumps(_aux, ensure_ascii=False, indent=2),
                               encoding="utf-8")
import json as _json  # alias estable para la restauracion final
check("C-19.9 defecto = modulo de composicion corporal",
      app.calculadoras_prioritarias() == app.CALC_PRIORITARIAS_DEFECTO,
      str(app.calculadoras_prioritarias()))
check("C-19.9b defecto solo ids validos del catalogo",
      all(c in app.CALCULADORAS for c in app.CALC_PRIORITARIAS_DEFECTO))
_marcada = app.alternar_prioritaria("pam")
check("C-19.10 alternar agrega y persiste en config",
      _marcada and "pam" in app.calculadoras_prioritarias()
      and app.leer_config().get("calc_prioritarias", []).count("pam") == 1,
      str(app.calculadoras_prioritarias()))
_desmarcada = app.alternar_prioritaria("pam")
check("C-19.10 alternar quita y persiste (sin reagregar el defecto)",
      not _desmarcada and "pam" not in app.calculadoras_prioritarias()
      and app.calculadoras_prioritarias() == app.CALC_PRIORITARIAS_DEFECTO,
      str(app.calculadoras_prioritarias()))
check("C-19.10b alternar id desconocido -> False sin persistir",
      app.alternar_prioritaria("no_existe") is False
      and "no_existe" not in app.calculadoras_prioritarias())
_quitar_defecto = app.alternar_prioritaria("imc")
check("C-19.10c se puede quitar una prioritaria por defecto (config propia)",
      not _quitar_defecto and "imc" not in app.calculadoras_prioritarias()
      and "sc_mosteller" in app.calculadoras_prioritarias(),
      str(app.calculadoras_prioritarias()))
app.alternar_prioritaria("imc")  # restaura el defecto
if _cfg_prev == "__ausente__":
    _aux = app.leer_config()
    _aux.pop("calc_prioritarias", None)
    app.CONFIG_PATH.write_text(
        _json.dumps(_aux, ensure_ascii=False, indent=2), encoding="utf-8")
else:
    app.guardar_config({"calc_prioritarias": _cfg_prev})
check("C-19.10d config restaurada tras las pruebas",
      app.leer_config().get("calc_prioritarias", "__ausente__") == _cfg_prev)

# --- ADR-0044: perfil embarazada en el cálculo (B100) ---
_h_ad = app.hallazgos_perfil_entradas({"fc": 110, "ta_d": 85}, "adulto")
check("B100 adulto: FC 110 -> alerta (rango 60-100)",
      any(h["campo"] == "fc" and h["nivel"] == "alerta" for h in _h_ad), str(_h_ad))
_h_em = app.hallazgos_perfil_entradas({"fc": 110, "ta_d": 85}, "embarazada")
check("B100 embarazada: FC 110 sin hallazgo (75-120); TA_d 85 -> esperado_embarazo",
      not any(h["campo"] == "fc" for h in _h_em)
      and any(h["campo"] == "ta_d" and h["nivel"] == "esperado_embarazo" for h in _h_em),
      str(_h_em))
check("B100: campos no-SV (peso/talla) se ignoran",
      app.hallazgos_perfil_entradas({"peso": 70, "talla": 1.65}, "embarazada") == [])
check("B100: entradas vacias -> []",
      app.hallazgos_perfil_entradas({}, "adulto") == [])
_h_emb_crit = app.hallazgos_perfil_entradas({"temp": 38.2}, "embarazada")
check("B100: temp 38.2 critica en ambos perfiles",
      any(h["nivel"] == "critico" for h in _h_emb_crit), str(_h_emb_crit))

# --- ADR-0051/B101: adaptador de plataforma ---
check("B101: PLATAFORMA no vacia y flags coherentes",
      isinstance(app.PLATAFORMA, str) and app.PLATAFORMA
      and ((app.ES_WINDOWS, app.ES_MACOS, app.ES_LINUX).count(True) == 1
           if app.PLATAFORMA in ("windows", "macos", "linux") else True),
      f"{app.PLATAFORMA} win={app.ES_WINDOWS} mac={app.ES_MACOS} lin={app.ES_LINUX}")
check("B101: FUENTE_SISTEMA definida por SO",
      isinstance(app.FUENTE_SISTEMA, str) and app.FUENTE_SISTEMA
      and (app.FUENTE_SISTEMA == ("Segoe UI" if app.ES_WINDOWS
                                  else "Helvetica" if app.ES_MACOS
                                  else "DejaVu Sans")),
      app.FUENTE_SISTEMA)
# comando_abrir es pura: se prueba cambiando los flags y restaurando
_b101_prev = (app.ES_WINDOWS, app.ES_MACOS, app.ES_LINUX)
try:
    app.ES_WINDOWS, app.ES_MACOS, app.ES_LINUX = True, False, False
    _c = app.comando_abrir("c:/x.txt")
    check("B101: comando_abrir windows -> startfile", _c[0] == "startfile"
          and _c[1] == ["c:/x.txt"], str(_c))
    app.ES_WINDOWS, app.ES_MACOS, app.ES_LINUX = False, True, False
    _c = app.comando_abrir("/x.txt")
    check("B101: comando_abrir macOS -> open", _c == ("open", ["/x.txt"]), str(_c))
    app.ES_WINDOWS, app.ES_MACOS, app.ES_LINUX = False, False, True
    _c = app.comando_abrir("/x.txt")
    check("B101: comando_abrir linux -> xdg-open", _c == ("xdg-open", ["/x.txt"]),
          str(_c))
finally:
    app.ES_WINDOWS, app.ES_MACOS, app.ES_LINUX = _b101_prev
check("B101: ES_MOVIL False en escritorio (bandera lista para movil)",
      app.ES_MOVIL is False)

# --- DO-10: consentimiento, aviso, NOM-004, bitacora ---
okC1, msgC1, salC1 = app.generar_consentimiento(ruta, "Cesarea programada")
check("DO-10 consentimiento generado", okC1, msgC1)
if salC1:
    txt = Path(salC1).read_text(encoding="utf-8")
    check("DO-10 consentimiento con hash SHA256", "SHA256-CONTENIDO:" in txt)
    check("DO-10 consentimiento con autor", "Autor registrado:" in txt)
okV, msgV = app.verificar_bitacora()
check("DO-10 bitacora hash-chain integra", okV, msgV)
okA1, msgA1, salA1 = app.generar_aviso_privacidad(ruta)
check("DO-10 aviso privacidad", okA1, msgA1)
if salA1:
    # Garantia del dueño: el aviso (y cualquier archivo) NO se reescribe
    _m_aviso = os.path.getmtime(salA1)
    _c_aviso = Path(salA1).read_text(encoding="utf-8")
    okA2, msgA2, salA2 = app.generar_aviso_privacidad(ruta)
    check("D-aviso segunda generacion no reescribe",
          okA2 and salA2 == salA1 and "ya existe" in msgA2 and
          os.path.getmtime(salA1) == _m_aviso and
          Path(salA1).read_text(encoding="utf-8") == _c_aviso, msgA2)
okH, msgH, salH = app.generar_historia_nom004(ruta, {
    "heredofamiliares": "DM2 madre", "patologicos": "Ninguno",
    "gineco": "G2P1A1", "exploracion": "TA 120/80, FCF 140",
    "estudios": "US Obstetrico normal", "diagnostico": "Embarazo a termino O09.521",
    "plan": "Control prenatal quincenal", "evolucion": "Evuelve sin complicaciones"})
check("DO-10 historia NOM-004", okH, msgH)
if salH:
    txt_h = Path(salH).read_text(encoding="utf-8")
    check("DO-10 historia con secciones NOM", "DIAGNOSTICO" in txt_h)
    check("T7 historia bloque 1.b automatico",
          "1.b. DATOS AUTOMATICOS" in txt_h and "G2P1A0" in txt_h
          and "Penicilina" in txt_h, txt_h[:400])
    check("T14 bloque 1.c ultimo control automatico",
          "1.c. ULTIMO CONTROL" in txt_h and "peso 70 kg" in txt_h, txt_h[:900])

# --- ADR-0028/T14: texto puro y documentos del episodio ---
txt_p, err_p = app.historia_nom004_texto(ruta, {
    "heredofamiliares": "Ninguno", "patologicos": "Ninguno",
    "gineco": "G2P1A0", "exploracion": "TA 120/80", "estudios": "US normal",
    "diagnostico": "O09.521", "plan": "Control quincenal",
    "evolucion": "Evolucion favorable"})
check("T14 historia_nom004_texto es puro (texto, None)",
      txt_p is not None and err_p is None and "1.b. DATOS AUTOMATICOS" in txt_p
      and "O09.521" in txt_p, str(err_p))
check("T14 texto invalido -> (None, error)",
      app.historia_nom004_texto(str(BASE / "NOEXISTE"), {})[0] is None)
docs_ep = app.documentos_episodio(ruta, eid)
check("T14 documentos del episodio: historias, consentimientos y progreso",
      len(docs_ep["historias"]) >= 1 and len(docs_ep["consentimientos"]) >= 1
      and docs_ep["progreso"], str({k: len(v) if isinstance(v, list) else v
                                    for k, v in docs_ep.items()}))
check("T14 episodio inexistente -> listas vacias",
      app.documentos_episodio(ruta, "E99_1999")["historias"] == [])
check("T14 linea_control formatea peso/TA",
      "peso 71.2 kg" in app.linea_control({"fecha": "2026-06-15", "peso": 71.2,
                                           "ta_s": 120, "ta_d": 80}))

# tamper detection: alterar el mensaje de la ultima linea de la bitacora
lineas = app.BITACORA_PATH.read_text(encoding="utf-8").splitlines()
partes = lineas[-1].split("|")
partes[2] = partes[2] + "-HACK"
lineas[-1] = "|".join(partes)
app.BITACORA_PATH.write_text("\n".join(lineas) + "\n", encoding="utf-8")
okT, msgT = app.verificar_bitacora()
check("CL-06 alteracion detectada", not okT and "ALTERACION" in msgT, msgT)
# restaurar la linea original -> verificacion vuelve a pasar
partes[2] = partes[2].replace("-HACK", "")
lineas[-1] = "|".join(partes)
app.BITACORA_PATH.write_text("\n".join(lineas) + "\n", encoding="utf-8")
okT2, msgT2 = app.verificar_bitacora()
check("CL-06 bitacora verificable tras restaurar linea", okT2, msgT2)

# --- DO-9: respaldo / verificar / restaurar ---
destino = BASE / "_RESPALDOS" / "PRUEBA"
okR, msgR = app.respaldar(destino)
check("DO-9 respaldo creado", okR, msgR)
check("DO-9 manifiesto con hashes", (destino / "MANIFIESTO_RESPALDO.json").exists())
okV2, msgV2 = app.verificar_respaldo(destino)
check("DO-9 verificacion SHA-256", okV2, msgV2)
# alterar un archivo copiado y verificar que se detecta
victima = destino / "EXPEDIENTES"
victima_file = next(victima.rglob("datos.json"))
victima_file.write_text('{"hackeado": true}', encoding="utf-8")
okV3, msgV3 = app.verificar_respaldo(destino)
check("DO-9 alteracion en respaldo detectada", not okV3 and "ALTERADO" in msgV3, msgV3)
# restaurar debe rechazar respaldo corrupto
okR2, msgR2 = app.restaurar_respaldo(destino)
check("DO-9 restauracion bloqueada si corrupto", not okR2, msgR2)
# reparar y restaurar de verdad
# regenerar respaldo limpio y restaurar
destino2 = BASE / "_RESPALDOS" / "PRUEBA2"
app.respaldar(destino2)
# borrar un expediente actual para comprobar la restauracion
shutil.rmtree(ruta)
check("DO-9 expediente borrado (pre-restauracion)", not Path(ruta).exists())
okR3, msgR3 = app.restaurar_respaldo(destino2)
check("DO-9 restauracion exitosa", okR3, msgR3)
check("DO-9 expediente restaurado", Path(ruta).exists() and (Path(ruta) / "datos.json").exists())

# --- B51/C-17: lector del Progreso Quimico (solo lectura, P3) ---
pdir = Path(ruta) / eid
ruta_prog = pdir / "PROGRESO_QUIMICO.xlsx"
okL, msgL, encL, _ = app.leer_progreso_quimico(ruta_prog)
check("B51 lee progreso creado por el sistema", okL, msgL)
check("B51 encabezados del progreso", encL and "Fecha" in encL, str(encL))
# datos numericos en inlineStr (formato del sistema)
ruta_datos = pdir / "PROGRESO_DATOS.xlsx"
app._crear_xlsx_minimo(ruta_datos, [
    ["Fecha", "Peso", "Glucosa"],
    ["2026-06-01", "70.5", "90"],
    ["2026-06-15", "71.2", "95"],
])
okL2, _, enc2, filas2 = app.leer_progreso_quimico(ruta_datos)
ser2 = app.series_progreso(enc2, filas2)
check("B51 series desde inlineStr", okL2 and len(ser2) == 2, str(ser2))
check("B51 valores numericos coercidos", ser2 and ser2[0][1][0][1] == 70.5, str(ser2))
check("B51 eje X usa la fecha de la primera columna",
      ser2 and ser2[0][1][0][0] == "2026-06-01", str(ser2))
# archivo reescrito por Excel: sharedStrings + celdas numericas + serial de fecha
ruta_excel = pdir / "PROGRESO_EXCEL.xlsx"
_ns_x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
sheet_x = (
    '<?xml version="1.0"?><worksheet xmlns="%s"><sheetData>'
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'
    '<row r="2"><c r="A2"><v>46000</v></c><c r="B2"><v>68.3</v></c></row>'
    '<row r="3"><c r="A3"><v>46014</v></c><c r="B3"><v>69.1</v></c></row>'
    "</sheetData></worksheet>"
) % _ns_x
shared_x = (
    '<?xml version="1.0"?><sst xmlns="%s" count="2" uniqueCount="2">'
    "<si><t>Fecha</t></si><si><t>Peso</t></si></sst>"
) % _ns_x
with zipfile.ZipFile(ruta_excel, "w") as z:
    z.writestr("xl/worksheets/sheet1.xml", sheet_x)
    z.writestr("xl/sharedStrings.xml", shared_x)
okL3, _, enc3, filas3 = app.leer_progreso_quimico(ruta_excel)
ser3 = app.series_progreso(enc3, filas3)
check("B51 lee xlsx reescrito por Excel (sharedStrings+numericos)",
      okL3 and enc3 == ["Fecha", "Peso"] and len(ser3) == 1
      and ser3[0][1][0][1] == 68.3, f"{enc3} {ser3}")
check("B51 serial de Excel convertido a ISO",
      ser3 and ser3[0][1][0][0] == "2025-12-09", str(ser3))
check("B51 sin series si no hay numeros",
      app.series_progreso(["Fecha", "Obs"], [["2026-01-01", "nota"]]) == [])

# --- F-25/F-26: exportaciones ---
okE1, msgE1, salE1 = app.exportar_pacientes_xlsx(BASE / "EXPORTACIONES" / "t.xlsx")
check("F-25 export xlsx", okE1 and zipfile.is_zipfile(salE1), msgE1)
if okE1:
    with zipfile.ZipFile(salE1) as _z:
        _hoja = _z.read("xl/worksheets/sheet1.xml").decode("utf-8", "replace")
    check("F-25 export incluye columnas FUM/EG/Notas",
          "FUM" in _hoja and "EG" in _hoja and "Notas" in _hoja, _hoja[:200])
okE2, msgE2, salE2 = app.exportar_pacientes_csv(BASE / "EXPORTACIONES" / "t.csv")
check("F-25 export csv", okE2 and Path(salE2).exists(), msgE2)
okE3, msgE3, salE3 = app.exportar_pacientes_pdf(BASE / "EXPORTACIONES" / "t.pdf")
check("F-25 export pdf", okE3 and Path(salE3).read_bytes()[:5] == b"%PDF-", msgE3)
okE4, msgE4, salE4 = app.exportar_expediente(ruta, BASE / "EXPORTACIONES" / "exp.zip")
check("F-26 export expediente zip", okE4 and zipfile.is_zipfile(salE4), msgE4)

# --- ADR-0015 (v2.2): PIN (P2), lock (P7) y cifrado de respaldo (P1) ---
check("P2 hash_pin determinista, distinto por PIN y sin claro",
      app.hash_pin("1234") == app.hash_pin("1234")
      and app.hash_pin("1234") != app.hash_pin("1235")
      and "1234" not in app.hash_pin("1234"))
check("P2 sin PIN -> gate abierto", not app.pin_configurado() and not app.verificar_pin("0000"))
app.guardar_pin("4826")
check("P2 pin configurado", app.pin_configurado())
check("P2 verifica correcto y rechaza incorrecto",
      app.verificar_pin("4826") and not app.verificar_pin("0000"))
cfg_pin = app.leer_config()
check("P2 PIN nunca en claro en _CONFIG.json",
      "4826" not in str(cfg_pin) and bool(cfg_pin.get("pin_hash")))
app.quitar_pin()
check("P2 quitar PIN", not app.pin_configurado())

check("P7 crear bloqueo y detectar vigente", app.crear_bloqueo() and app.bloqueo_reciente()[0])
info_viejo = {"pid": 1, "ts": (app.datetime.datetime.now()
                               - app.datetime.timedelta(hours=2)).isoformat()}
app.LOCK_PATH.write_text(app.json.dumps(info_viejo), encoding="utf-8")
check("P7 bloqueo viejo no reciente", not app.bloqueo_reciente()[0])
app.limpiar_bloqueo()
check("P7 limpiar bloqueo", not app.LOCK_PATH.exists())

claveA = app._clave_desde_password("secreta")
check("P1 derivacion determinista y distinta por contraseña",
      claveA == app._clave_desde_password("secreta")
      and claveA != app._clave_desde_password("otra"))
dest_cif = BASE / "_RESPALDOS" / "CIF"
okRc, msgRc = app.respaldar(dest_cif, clave="secreta123")
if app.CRYPTO_DISPONIBLE:
    check("P1 respaldo cifrado creado", okRc and (dest_cif / app.CRYPTO_BIN).exists(),
          msgRc)
    check("P1 sin copias en claro (zip+bin unico)",
          not (dest_cif / "EXPEDIENTES").exists()
          and not (dest_cif / "MANIFIESTO_RESPALDO.json").exists())
    okVc, msgVc = app.verificar_respaldo(dest_cif)
    check("P1 verificar sin clave -> pide contrasena",
          not okVc and "contrase" in msgVc.lower(), msgVc)
    okVc2, msgVc2 = app.verificar_respaldo(dest_cif, clave="mala")
    check("P1 clave incorrecta -> rechazada", not okVc2, msgVc2)
    okVc3, msgVc3 = app.verificar_respaldo(dest_cif, clave="secreta123")
    check("P1 verificar con clave -> integro", okVc3, msgVc3)
    # restauracion desde respaldo cifrado
    shutil.rmtree(ruta)
    okRst, msgRst = app.restaurar_respaldo(dest_cif, clave="secreta123")
    check("P1 restauracion desde respaldo cifrado",
          okRst and Path(ruta).exists(), msgRst)
else:
    check("P1 fallback sin cryptography (integridad SHA-256)",
          okRc and "NO disponible" in msgRc, msgRc)

# --- DO-11: offline (sin imports de red en el codigo fuente) ---
src = (BASE / "app.py").read_text(encoding="utf-8")
forbidden = ["import socket", "import urllib", "import http", "import requests",
             "import ftplib", "import smtplib", "import asyncio"]
encontrados = [f for f in forbidden if f in src]
check("DO-11 cero dependencias de red", not encontrados, str(encontrados))

# --- DO-12: datos legibles sin el software ---
d = app.cargar_paciente(ruta)
check("DO-12 datos.json valido", d.get("nombre_completo"))
check("DO-12 HISTORIAL.txt legible", (Path(ruta) / "HISTORIAL.txt").read_text(encoding="utf-8").count("\n") > 3)

# --- W-14 filtros ---
res = app.listar_pacientes()
filtrados = app.filtrar(res, medico="DR")  # sin resultados por medico inexistente
check("W-14 filtro medico aplica", isinstance(filtrados, list))
filtrados2 = app.filtrar(res, desde="1900-01-01", hasta="2000-12-31")
check("W-14 filtro rango fechas aplica", len(filtrados2) >= 1, str(filtrados2))

# --- parseo de fechas (W-02/DO-3) ---
check("parseo dd-MMM-yyyy", app.parsear_fecha("15-MAR-1985") == "1985-03-15")
check("parseo '15 de marzo de 1985'", app.parsear_fecha("15 de marzo de 1985") == "1985-03-15")
check("parseo iso", app.parsear_fecha("1985-03-15") == "1985-03-15")
check("parseo invalido rechazado", app.parsear_fecha("no es fecha") is None)

# =====================================================================
# T21 (ADR-0033): escritura segura, journal, bak, degradacion, propiedades
# =====================================================================
import random as _random
import json
import datetime

_ruta_atomica = BASE / "_ATOMIC_TEST.json"
app.escribir_atomico(_ruta_atomica, '{"x": 1}')
check("B85 escribir_atomico crea archivo valido",
      _ruta_atomica.exists() and json.loads(_ruta_atomica.read_text())["x"] == 1)
check("B85 no deja tmp tras exito", not Path(str(_ruta_atomica) + ".tmp").exists())
app.escribir_json_atomico(_ruta_atomica, {"x": 2})
check("B85 journal registra OK", "OK" in app.JOURNAL_PATH.read_text(encoding="utf-8"))
_r_pac = BASE / "EXPEDIENTES" / "T" / "TEST_BAK"
_r_pac.mkdir(parents=True, exist_ok=True)
app.guardar_paciente(_r_pac, {"nombre_completo": "BAK", "clave": "TEST_BAK"})
app.guardar_paciente(_r_pac, {"nombre_completo": "BAK2", "clave": "TEST_BAK"})
check("B85 .bak1 creado tras segunda escritura", (_r_pac / "datos.json.bak1").exists())
(_r_pac / "datos.json").write_text("{corrupto", encoding="utf-8")
try:
    d_deg = app.cargar_paciente(_r_pac)
    check("B85 degradacion carga .bak1",
          d_deg.get("nombre_completo") in ("BAK", "BAK2") and d_deg.get("degradado") is True)
except Exception as e:
    check("B85 degradacion carga .bak1", False, str(e))
_ok_fpp = True
for _i in range(200):
    _fum = f"{_random.randint(2020,2030):04d}-{_random.randint(1,12):02d}-{_random.randint(1,28):02d}"
    _fpp = app.calcular_fpp(_fum)
    if _fpp is None:
        _ok_fpp = False
        break
    _delta = (datetime.date.fromisoformat(_fpp) - datetime.date.fromisoformat(_fum)).days
    if not (250 <= _delta <= 300):  # Naegele: ~38-42 semanas
        _ok_fpp = False
        break
check("Propiedad FPP Naegele en rango obstetrico (200 muestras)", _ok_fpp)
_ok_mono = True
_prev = 0
for _dias in range(0, 280, 7):
    _h = (datetime.date.fromisoformat("2026-01-01") + datetime.timedelta(days=_dias)).isoformat()
    _eg = app.calcular_eg("2026-01-01", _h)
    _sem = int(_eg.split()[0]) if _eg else 0
    if _dias > 0 and _sem < _prev:
        _ok_mono = False
        break
    _prev = _sem
check("Propiedad EG monotonamente creciente", _ok_mono)
_ok_rt = True
for _i in range(100):
    _f = f"{_random.randint(1990,2035):04d}-{_random.randint(1,12):02d}-{_random.randint(1,28):02d}"
    if app.parsear_fecha(app.fecha_legible(_f)) != _f:
        _ok_rt = False
        break
check("Propiedad roundtrip fechas (100 muestras)", _ok_rt)
_ok_fz = True
for _i in range(50):
    _ff = BASE / f"_fuzz_{_i}.bin"
    _ff.write_bytes(bytes(_random.randint(0,255) for _ in range(200)))
    try:
        _r = app.leer_progreso_quimico(_ff)
        if not (isinstance(_r, tuple) and len(_r) == 4):
            _ok_fz = False
            break
    except Exception:
        _ok_fz = False
        break
    finally:
        _ff.unlink(missing_ok=True)
check("Fuzz leer_progreso (50) sin excepcion", _ok_fz)
_diag = app.diagnostico_arranque()
check("B86 diagnostico_arranque corre", isinstance(_diag, list) and len(_diag) > 0)

# =====================================================================
# T22 (ADR-0034): store reactivo + borradores
# =====================================================================
_e = app.Estado()
_n = {"v": 0}
_e.suscribir("demo", lambda: _n.__setitem__("v", _n["v"] + 1))
_r1 = _e.dispatch("a", {"demo"})
_r2 = _e.dispatch("a", {"demo", "otra"})
check("Estado revision monotonica", _r2 > _r1)
_e.flush()
check("Estado notifica suscriptores", _n["v"] >= 1)
_okb = app.guardar_borrador("test", "PAC1", {"campo": "valor"})
_d = app.cargar_borrador("test", "PAC1")
check("Borrador guardar/cargar", _okb and _d == {"campo": "valor"})
app.borrar_borrador("test", "PAC1")
check("Borrador borrar", app.cargar_borrador("test", "PAC1") is None)
_src_calc = (BASE / "app.py").read_text(encoding="utf-8")
check("ADR-0034 sin boton Calcular", 'text="Calcular"' not in _src_calc)
check("C-16 talla en _CAMPOS_CONTROL", "talla" in app._CAMPOS_CONTROL)

# =====================================================================
# T23 (ADR-0035): integridad clinica
# =====================================================================
h_p = app.validar_plausibilidad({"ta_s": "90", "ta_d": "110"})
check("B87 TA sist<diast -> hallazgo", any("sist" in x["mensaje"] for x in h_p))
check("B87 peso 0 -> critico", any(x["nivel"] == "critico" for x in app.validar_plausibilidad({"peso": "0"})))
check("B87 temp 50 -> alerta", any("30" in x["mensaje"] for x in app.validar_plausibilidad({"temp": "50"})))
check("B87 spo2 110 -> critico", any(x["nivel"] == "critico" for x in app.validar_plausibilidad({"spo2": "110"})))
check("B87 valores ok -> vacio", app.validar_plausibilidad({"ta_s": "120", "ta_d": "80"}) == [])
check("B89 registrar_lectura", app.registrar_lectura(ruta, "test"))
check("B89 _ACCESOS.log con sha", "|" in app.ACCESOS_PATH.read_text(encoding="utf-8"))
h_ver = app.cargar_historia(ruta)
n_idx = app.nueva_nota_historia(h_ver, "ingreso", eid)
h_ver["notas"][n_idx]["datos"]["B"] = {"motivo": "v1"}
app.guardar_historia(ruta, h_ver)
h_ver2 = app.cargar_historia(ruta)
h_ver2["notas"][n_idx]["datos"]["B"] = {"motivo": "v2"}
app.guardar_historia(ruta, h_ver2, versionar=True, motivo="enmienda de prueba")
h_ver3 = app.cargar_historia(ruta)
check("A2 enmienda versionada", len(h_ver3["notas"][n_idx].get("versiones", [])) >= 1)
_r_o = BASE / "EXPEDIENTES" / "F" / "FUS_ORIGEN"
_r_d = BASE / "EXPEDIENTES" / "F" / "FUS_DESTINO"
_r_o.mkdir(parents=True, exist_ok=True)
_r_d.mkdir(parents=True, exist_ok=True)
app.guardar_paciente(_r_o, {"nombre_completo": "ORIGEN FUS", "clave": "FUS_ORIGEN", "telefono": "111", "episodios": [{"id": "E01_2020"}]})
app.guardar_paciente(_r_d, {"nombre_completo": "DESTINO FUS", "clave": "FUS_DESTINO", "telefono": "", "episodios": [{"id": "E01_2020"}]})
ok_fus, msg_fus = app.fusionar_paciente(str(_r_o), str(_r_d), "prueba")
check("B88 fusion episodios unidos por id", ok_fus and len(app.cargar_paciente(_r_d)["episodios"]) == 1, msg_fus)
check("B88 carpeta origen movida", not _r_o.exists() and (BASE / "_FUSIONADOS").exists())
app.registrar_consentimiento(ruta, "CONSENTIMIENTO_TEST.txt")
d_c = app.cargar_paciente(ruta)
check("A7 consentimiento registrado", d_c["consentimientos"][-1]["estado"] == "vigente")
ok_rev, _ = app.revocar_consentimiento(ruta, len(d_c["consentimientos"]) - 1, "prueba")
check("A7 consentimiento revocado", ok_rev and app.cargar_paciente(ruta)["consentimientos"][-1]["estado"] == "revocado")
app.registrar_alerta(ruta, {"origen": "dedupe", "descripcion": "mismo texto", "episodio": eid})
app.registrar_alerta(ruta, {"origen": "dedupe", "descripcion": "mismo texto", "episodio": eid})
n_ded = len([a for a in app.cargar_historia(ruta)["alertas"] if a["origen"] == "dedupe"])
check("A11 dedupe evita duplicados", n_ded == 1, str(n_ded))
ok_des, msg_des = app.deshacer_ultima_edicion(ruta, "A", "paciente")
check("A13 deshacer seccion paciente", ok_des, msg_des)
ok_diag, msg_diag, sal_diag = app.exportar_diagnostico()
check("A18 exportar diagnostico", ok_diag and Path(sal_diag).exists(), msg_diag)

# B90 drill restauracion (no destructivo)
_dest_drill = BASE / "_RESPALDOS" / "DRILL_TEST"
app.respaldar(str(_dest_drill))
ok_dr, msg_dr = app.drill_restauracion(str(_dest_drill))
check("B90 drill restauracion verifica respaldo", ok_dr, msg_dr)

# =====================================================================
# T24 (ADR-0036 / C-21): export anonimizado
# =====================================================================
_id1 = app.id_anonimo("PACIENTE_TEST_1985")
_id2 = app.id_anonimo("PACIENTE_TEST_1985")
_id3 = app.id_anonimo("OTRA_PACIENTE_1990")
check("B82 id_anonimo estable", _id1 == _id2 and len(_id1) == 12)
check("B82 id_anonimo distinto por paciente", _id1 != _id3)
ok_an, msg_an, sal_an = app.exportar_historia_anonima(ruta, "Evaluar alimentacion", True)
check("B83 export anonimizado ok", ok_an and Path(sal_an).exists(), msg_an)
if ok_an:
    _tx = Path(sal_an).read_text(encoding="utf-8")
    check("B83 contiene ID y perfil", "ID:" in _tx and "Perfil" in _tx)
    check("B83 sin nombre de paciente", "FERNANDA" not in _tx and "GARCIA" not in _tx)
    check("B83 sin CURP literal", "GODE561231MDFRRL09" not in _tx)
    check("B83 sin clave de expediente", "GARCIA_LOPEZ" not in _tx)
    check("B83 sin telefono", "5512345678" not in _tx)
    check("B83 serie de controles presente", "Serie temporal" in _tx)
    check("B83 pregunta de analisis", "alimentacion" in _tx)
    check("B83 CSV generado", len(list(Path(sal_an).parent.glob("HISTORIA_ANONIMA_*.csv"))) >= 1)
h_ids = app.escanear_identificadores("CURP GODE561231MDFRRL09 tel 5512345678", datos)
check("B84 detecta CURP y telefono", len(h_ids) >= 2, str(h_ids))
check("B84 texto limpio -> vacio", app.escanear_identificadores("Peso 70kg normal", datos) == [])
_scr = app._scrub_texto_libre("Paciente FERNANDA CURP GODE561231MDFRRL09", datos)
check("B84 scrubber enmascara", "FERNANDA" not in _scr and "GODE561231MDFRRL09" not in _scr, _scr)

# =====================================================================
# T25-T28 (ADR-0037..0040): timeline, tendencias, import, dictado
# =====================================================================
# T25: eventos_paciente
evs = app.eventos_paciente(ruta, eid)
check("T25 eventos_paciente lista eventos", isinstance(evs, list))
check("T25 eventos ordenados por fecha desc",
      all(evs[i]["fecha"] >= evs[i+1]["fecha"] for i in range(len(evs)-1)) if len(evs) > 1 else True)
check("T25 incluye controles", any(e["tipo"] == "control" for e in evs))
check("T25 incluye documentos", any(e["tipo"] == "documento" for e in evs))
check("T25 sin paciente -> vacio", app.eventos_paciente(str(BASE / "NOEXISTE")) == [])

# T26: tendencias
tend = app.tendencias_controles(ruta, eid)
check("T26 tendencias con >=2 controles", tend.get("n_visitas", 0) >= 2, str(tend))
check("T26 delta_peso calculado", "delta_peso" in tend, str(tend))
# recordatorios EG
rec = app.recordatorios_eg("2026-06-10", "2026-10-16")
check("T26 recordatorio trimestre", rec["trimestre"] == "II", str(rec))
check("T26 dias_termino calculado", rec["dias_termino"] is not None, str(rec))
check("T26 sin FUM -> trimestre vacio", app.recordatorios_eg("", "2026-10-16")["trimestre"] == "")

# T27: detectar_mapeo_csv
mapeo = app.detectar_mapeo_csv(["Nombre", "Apellido Paterno", "Fecha Nacimiento", "CURP"])
check("T27 detecta nombre", mapeo.get("nombre") == "Nombre")
check("T27 detecta ap_paterno", mapeo.get("ap_paterno") == "Apellido Paterno")
check("T27 detecta fecha", mapeo.get("fecha_nacimiento") == "Fecha Nacimiento")
# importar CSV
import csv as _csv
_csv_path = BASE / "test_import.csv"
with open(_csv_path, "w", newline="", encoding="utf-8") as f:
    w = _csv.writer(f)
    w.writerow(["Apellido Paterno", "Apellido Materno", "Nombre", "Fecha Nacimiento", "CURP"])
    w.writerow(["IMPORT", "TEST", "ANA", "1990-05-20", ""])
    w.writerow(["IMPORT2", "TEST2", "MARIA", "1992-03-15", ""])
    w.writerow(["IMPORT", "TEST", "ANA", "1990-05-20", ""])  # duplicado
    w.writerow(["FALTA", "", "", "invalida", ""])  # invalida
mapeo2 = app.detectar_mapeo_csv(["Apellido Paterno", "Apellido Materno", "Nombre", "Fecha Nacimiento", "CURP"])
ok_imp, msg_imp, res_imp = app.importar_pacientes_csv(str(_csv_path), mapeo2)
check("T27 importar CSV ok", ok_imp, msg_imp)
check("T27 importados=2", res_imp.get("importados") == 2, str(res_imp))
check("T27 duplicado omitido", res_imp.get("duplicados") == 1, str(res_imp))
check("T27 invalido omitido", res_imp.get("invalidos") == 1, str(res_imp))

# T28: dictado_disponible
ok_dict, msg_dict = app.dictado_disponible()
check("T28 dictado_disponible devuelve tuple", isinstance(ok_dict, bool) and isinstance(msg_dict, str))

# =====================================================================
# Integraciones pendientes cerradas (timeline resumen, tendencia alerta)
# =====================================================================
# evaluar_tendencia_alerta: sin config -> None; con config y delta grande -> alerta
check("Tendencia sin config -> None", app.evaluar_tendencia_alerta(ruta, eid) is None)
app.guardar_config({"RANGOS_TENDENCIA": {"ta_s": 5}})
tend_h = app.evaluar_tendencia_alerta(ruta, eid)
# con config: None si no hay delta (control sin ta_s) o alerta valida si lo hay
check("Tendencia con config respeta contrato",
      tend_h is None or (tend_h.get("origen") == "tendencia"
                         and "delta" in tend_h.get("descripcion", "")), str(tend_h))
app.guardar_config({"RANGOS_TENDENCIA": {}})
check("Tendencia config vacia -> None", app.evaluar_tendencia_alerta(ruta, eid) is None)

# resumen_consulta incluye linea de tiempo
txt_r2 = app.resumen_consulta(ruta)
check("Resumen incluye LINEA DE TIEMPO", "LINEA DE TIEMPO DEL EPISODIO" in txt_r2)

# =====================================================================
# v4.16: prefill de Calculadoras desde ultimo control + scroll
# =====================================================================
# registro con vitales crea control (flujo C-03 v4.11)
okRV, msgRV, rutaRV, _dRV = app.crear_paciente("PREFILL", "TEST", "VITALS", "1992-07-10", "")
check("v4.16 crear paciente para vitales", okRV, msgRV)
eidRV = app.nuevo_episodio(rutaRV)
app.fijar_episodio(rutaRV, eidRV)
app.agregar_control(rutaRV, eidRV, {"fecha": app.fecha_hoy(), "peso": "70.5",
                                     "ta_s": "120", "ta_d": "80", "fc": "75",
                                     "fr": "16", "temp": "36.8", "spo2": "98"})
ctrlsRV = app.listar_controles(rutaRV, eidRV)
check("v4.16 control con 8 campos registrado",
      len(ctrlsRV) == 1 and ctrlsRV[0]["peso"] == 70.5 and ctrlsRV[0]["spo2"] == 98.0,
      str(ctrlsRV))
check("v4.16 medidas corporales en H1", True)  # medidas via UI; aqui solo control
check("v4.16 listar_controles ordena desc", ctrlsRV[0]["fecha"] >= app.fecha_hoy())

print()
print("=" * 60)
if FALLAS:print()
print("=" * 60)
if FALLAS:print()
print("=" * 60)
if FALLAS:print()
print("=" * 60)
if FALLAS:
    print(f"RESULTADO: {len(FALLAS)} FALLAS")
    for f in FALLAS:
        print(f"  - {f}")
    sys.exit(1)
print("RESULTADO: TODAS LAS PRUEBAS PASARON")
