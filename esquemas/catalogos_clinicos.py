# -*- coding: utf-8 -*-
"""
CATALOGOS CLINICOS - rangos de signos vitales, valores criticos de lab,
escalas de alerta y registro de calculadoras.
==========================================================================
Plantilla de DATOS y funciones PURAS para transcribir a `app.py` (secciones
00b CATALOGOS CLINICOS + 08c SERVICIO ALERTAS + 08d SERVICIO CALCULADORAS)
durante los paquetes T14/T15 (ADR-0028). NO se importa desde app.py.

PROCEDENCIA DE LAS CIFRAS (ADR-0028):
Todas las cifras de ESTE archivo fueron aportadas explicitamente por el
dueno/memedico en su documento de referencia (2026-10-16). El software NO
inventa umbrales: lo que el medico no haya aportado queda marcado
"incompleta" y el sistema no puntua. Para editar cifras sin tocar logica,
el medico usa este archivo como referencia y se pegan en las constantes de
app.py (mismo patron que RANGOS_CONTROLES + NOTA_AL_MEDICO_RANGOS.md).

INVARIANTES (P2 clinico):
- Toda evaluacion devuelve HALLAZGOS, nunca bloqueos.
- La UI confirma antes de continuar ("Registrar de todos modos?").
- score_* devuelve None cuando faltan datos o la tabla esta incompleta.
"""

# ============================================================================
# 1. RANGOS DE SIGNOS VITALES (fuente: documento del medico, 2026-10-16)
# ============================================================================

# Perfil "adulto" = rango normal de referencia.
# Claves = nombre del campo (coinciden con _CAMPOS_CONTROL + temp/fr/spo2).
RANGOS_SIGNOS_VITALES = {
    "adulto": {
        #         (min_normal, max_normal, min_critico, max_critico)
        "ta_s":  (90, 119, None, 200),      # <90 hipotension - >=140 HTA - >200 precaucion extrema
        "ta_d":  (60, 79, None, 100),       # <60 hipotension - >=90 HTA - >100 precaucion extrema
        "fc":    (60, 100, None, None),     # <60 o >100 = alerta
        "fr":    (12, 20, None, None),      # <12 o >20 = alerta (centinela del deterioro)
        "temp":  (36.0, 37.2, 35.0, 37.5),  # axilar; >37.5 o <35.0 = alerta
        "spo2":  (94, 100, None, None),     # <94 alerta - <90 hipoxemia grave
    },
    # Ajustes fisiologicos del embarazo (no son enfermedad: son el nuevo "normal")
    "embarazada": {
        "ta_s":  (90, 119, None, 200),      # la sistolla no cambia por embarazo per se
        "ta_d":  (50, 79, None, 100),       # diastolica puede bajar 5-10 mmHg (2.o trimestre)
        "fc":    (75, 120, None, None),     # +15-20 lpm respecto de la pregestacion
        "fr":    (14, 22, None, None),      # +2 rpm respecto del adulto
        "temp":  (36.0, 37.2, 35.0, 37.5),
        "spo2":  (94, 100, None, None),
    },
}

# Combinacion clave de alerta obstetrica (embarazada): HTA + proteinuria.
ALERTA_PREECLAMPSIA = {
    "ta_s_min": 160,
    "ta_d_min": 110,
    "proteinuria_min": "2+",   # texto minimo; comparacion insensible a mayusculas
    "mensaje": ("COMBINACION CLAVE: presion >=160/110 mmHg con proteinuria >=2+ "
                "=> sospechar PREECLAMPSIA (emergencia obstetrica). "
                "Evaluar inmediatamente y registrar la respuesta."),
}

# ============================================================================
# 2. VALORES CRITICOS DE LABORATORIO - PANIC VALUES
#    (fuente: documento del medico; prevalecen los del laboratorio real)
# ============================================================================
# Cada entrada: prueba -> dict con umbrales opcionales (bajo/alto), unidad y
# texto de riesgo. Solo se evalua la prueba si tiene al menos un umbral.
UMBRALES_LAB = {
    "glucosa":     {"bajo": 40.0,   "alto": 500.0,  "unidad": "mg/dL",
                    "riesgo": "Coma por hipoglucemia grave; estado hiperglucemico "
                              "hiperosmolar/cetoacidosis diabetica"},
    "potasio":     {"bajo": 2.5,    "alto": 6.5,    "unidad": "mmol/L",
                    "riesgo": "Arritmias graves, paro cardiaco, paralisis de musculos "
                              "respiratorios"},
    "sodio":       {"bajo": 120.0,  "alto": 160.0,  "unidad": "mmol/L",
                    "riesgo": "Edema cerebral, convulsiones, alteracion de la conciencia"},
    "hemoglobina": {"bajo": 6.1,    "alto": 21.0,   "unidad": "g/dL",
                    "riesgo": "Anemia grave con hipoxia tisular; sindrome de "
                              "hiperviscosidad"},
    "troponina":   {"alto": 0.04,   "unidad": "ng/mL",
                    "riesgo": "Infarto agudo de miocardio, dano miocardico"},
    "bilirrubina": {"alto": 17.0,   "unidad": "mg/dL",
                    "riesgo": "Disfuncion hepatica grave, riesgo de encefalopatia "
                              "por bilirrubina"},
    "fibrinogeno": {"bajo": 80.0,   "alto": 999.0,  "unidad": "mg/dL",
                    "riesgo": "Tendencia grave a sangrado; estado de hipercoagulabilidad"},
}

# ============================================================================
# 3. ESCALAS DE ALERTA TEMPRANA (MEWS / NEWS)
#    REGLA ADR-0028: sin la tabla de puntos COMPLETA confirmada por el medico,
#    la escala NO puntua (score_* -> None). Aqui va lo aportado hasta hoy.
# ============================================================================
TABLAS_ESCALA = {
    "mews": {
        "items": ["ta_s", "fc", "fr", "temp", "avpu"],   # 5 indicadores
        "umbral_riesgo": 5,                               # >=5: OR 5.4 muerte / 10.9 UCI
        "tabla_puntos": None,   # <-- INCOMPLETA: medico debe aportar bandas por item
        "estado": "incompleta",
        "nota": "Puntaje MEWS >=5 asocia riesgo aumentado. Falta la tabla de "
                "puntos por banda (ej. ta_s<100 -> 2 pts) a confirmar por el medico.",
    },
    "news": {
        "items": ["fr", "spo2", "temp", "ta_s", "fc", "conciencia", "oxigeno"],
        "umbral_urgente": 5,    # >=5: evaluacion urgente
        "umbral_uci": 7,        # >=7: considerar UCI
        "tabla_puntos": None,   # <-- INCOMPLETA
        "estado": "incompleta",
        "nota": "NEWS agrega SpO2 y uso de oxigeno a MEWS. En Mexico >8 predice "
                "mortalidad/UCI. Falta tabla de puntos a confirmar por el medico.",
    },
    "qsofa": {
        "items": ["fr_ge_22", "alteracion_conciencia", "ta_s_le_100"],
        "umbral_riesgo": 2,     # >=2: mal pronostico (tamizaje de sepsis)
        "tabla_puntos": {       # COMPLETA: 1 punto por criterio presente
            "fr_ge_22": 1,
            "alteracion_conciencia": 1,
            "ta_s_le_100": 1,
        },
        "estado": "completa",
        "nota": "Tamizaje de sepsis a pie de cama. >=2 = mal pronostico.",
    },
}


def evaluar_signos_vitales(valores, perfil="adulto"):
    """Compara valores capturados contra RANGOS_SIGNOS_VITALES[perfil].

    valores: dict campo->numero (ta_s, ta_d, fc, fr, temp, spo2).
    Devuelve lista de hallazgos: [{"campo", "valor", "nivel", "mensaje"}]
    nivel: "critico" | "alerta" | "esperado_embarazo" (informativo).
    NUNCA bloquea; la UI confirma (P2 clinico).
    """
    hallazgos = []
    rangos = RANGOS_SIGNOS_VITALES.get(perfil) or RANGOS_SIGNOS_VITALES["adulto"]
    for campo, v in (valores or {}).items():
        if not isinstance(v, (int, float)) or campo not in rangos:
            continue
        mn, mx, crit_lo, crit_hi = rangos[campo]
        if (crit_lo is not None and v < crit_lo) or (crit_hi is not None and v > crit_hi):
            hallazgos.append({
                "campo": campo, "valor": v, "nivel": "critico",
                "mensaje": f"{campo.upper()} = {v:g} fuera de valores criticos "
                           f"({crit_lo if crit_lo is not None else '-'} - "
                           f"{crit_hi if crit_hi is not None else '-'})",
            })
        elif v < mn or v > mx:
            nivel = "alerta"
            extra = ""
            if perfil == "embarazada":
                extra = " (puede ser esperado en embarazo: evaluar con sintomas)"
                nivel = "esperado_embarazo" if campo in ("fc", "fr", "ta_d") else "alerta"
            hallazgos.append({
                "campo": campo, "valor": v, "nivel": nivel,
                "mensaje": f"{campo.upper()} = {v:g} fuera del rango normal "
                           f"({mn:g}-{mx:g}){extra}",
            })
    return hallazgos


def hallazgos_perfil_entradas(entradas, perfil="adulto"):
    """FILTRA de las entradas de una calculadora los campos que son signos
    vitales del catalogo y devuelve sus hallazgos contra el perfil elegido
    (adulto/embarazada). Informativo; NUNCA bloquea (P2 clinico).

    ADR-0044: la UI de C-19 usa esto con el casillero "Paciente embarazada"
    para mostrar la referencia del perfil sobre los SV que la formula consume
    (ta_s, ta_d, fc, fr, temp, spo2). No cambia la formula ni las bandas de
    la calculadora (esas son las del catalogo del medico).
    """
    rangos = RANGOS_SIGNOS_VITALES.get(perfil) or RANGOS_SIGNOS_VITALES["adulto"]
    valores = {k: v for k, v in (entradas or {}).items()
               if k in rangos and isinstance(v, (int, float))}
    return evaluar_signos_vitales(valores, perfil)


def evaluar_preeclampsia(valores, proteinuria_txt):
    """Alerta combinada HTA + proteinuria en embarazadas (obstetrica).

    Devuelve lista (vacia si no hay combinacion). Solo aplica si hay TA.
    """
    res = []
    ta_s = (valores or {}).get("ta_s")
    ta_d = (valores or {}).get("ta_d")
    prot = (proteinuria_txt or "").strip().upper()
    if not isinstance(ta_s, (int, float)) or not isinstance(ta_d, (int, float)):
        return res
    if ta_s >= ALERTA_PREECLAMPSIA["ta_s_min"] or ta_d >= ALERTA_PREECLAMPSIA["ta_d_min"]:
        # "2+" o mayor: "3+", "3+1" etc. contienen el umbral como prefijo de grado
        grado_min = ALERTA_PREECLAMPSIA["proteinuria_min"]  # "2+"
        if prot and (grado_min in prot or prot.startswith(("2", "3", "4"))):
            res.append({"campo": "preeclampsia", "valor": f"{ta_s:g}/{ta_d:g} + {prot}",
                        "nivel": "critico", "mensaje": ALERTA_PREECLAMPSIA["mensaje"]})
    return res


def evaluar_lab(prueba, valor):
    """Evalua un valor de laboratorio contra UMBRALES_LAB.

    prueba: clave insensible a mayusculas ("potasio", "K" no soportado - usar clave).
    Devuelve hallazgo dict o None.
    """
    if not isinstance(valor, (int, float)):
        return None
    umb = UMBRALES_LAB.get((prueba or "").strip().lower())
    if not umb:
        return None
    if "bajo" in umb and valor < umb["bajo"]:
        return {"campo": prueba, "valor": valor, "nivel": "critico",
                "mensaje": f"VALOR CRITICO: {prueba} = {valor:g} {umb['unidad']} "
                           f"(< {umb['bajo']:g}). Riesgo: {umb['riesgo']}. "
                           f"Notificar al medico solicitante de inmediato y registrar "
                           f"la respuesta (circuito cerrado)."}
    if "alto" in umb and valor > umb["alto"]:
        return {"campo": prueba, "valor": valor, "nivel": "critico",
                "mensaje": f"VALOR CRITICO: {prueba} = {valor:g} {umb['unidad']} "
                           f"(> {umb['alto']:g}). Riesgo: {umb['riesgo']}. "
                           f"Notificar al medico solicitante de inmediato y registrar "
                           f"la respuesta (circuito cerrado)."}
    return None


def score_qsofa(valores):
    """qSOFA: FR>=22, alteracion de conciencia, TA sist <=100. 1 pt c/u; >=2 alerta.

    valores: {"fr":..., "conciencia_alterada": bool, "ta_s":...}
    Devuelve (puntaje, mensaje) o (None, motivo) si falta FR o TA.
    """
    if not isinstance((valores or {}).get("fr"), (int, float)):
        return None, "qSOFA incompleto: falta frecuencia respiratoria."
    if not isinstance((valores or {}).get("ta_s"), (int, float)):
        return None, "qSOFA incompleto: falta tension arterial sistolica."
    pts = 0
    if valores["fr"] >= 22:
        pts += 1
    if valores.get("conciencia_alterada"):
        pts += 1
    if valores["ta_s"] <= 100:
        pts += 1
    msg = f"qSOFA = {pts} punto(s)." + (" MAL PRONOSTICO: >=2 indica riesgo elevado."
                                         if pts >= 2 else "")
    return pts, msg


def score_escala_incompleta(nombre):
    """Las escalas con tabla_puntos None devuelven None (ADR-0028: no inventar)."""
    esc = TABLAS_ESCALA.get(nombre)
    if not esc or esc.get("estado") != "completa":
        return None, (f"Escala {nombre.upper()} no puntua: tabla de puntos "
                      f"incompleta (pendiente de confirmacion del medico).")
    return None, f"Escala {nombre.upper()} sin implementar aun."


# ============================================================================
# 4. REGISTRO DE CALCULADORAS (ADR-0029)
#    Cada calculadora: grupo (8 grupos del medico), campos de entrada,
#    funcion pura, bandas de interpretacion, estado.
#    "activa" = implementada - "incompleta" = falta tabla/formula del medico.
# ============================================================================
# Campos de entrada: (id, etiqueta, tipo)  tipo: "num" | "txt" | "bool"
# Banda: (min, max, etiqueta, nivel) con None como abierto: (None, 18.49, ...)

CALCULADORAS = {
    # ---- I. CARDIOVASCULAR ----
    "pam": {
        "grupo": "Cardiovascular", "titulo": "Presion Arterial Media (PAM)",
        "campos": [("ta_s", "TA sistolica (mmHg)", "num"),
                   ("ta_d", "TA diastolica (mmHg)", "num")],
        "fn": "calc_pam",
        "bandas": [(None, 64, "Riesgo de hipoperfusion (PAM <65)", "alerta"),
                   (65, 100, "Perfusion adecuada", "ok"),
                   (101, None, "Elevada", "alerta")],
        "estado": "activa",
        "formula": "PAM = (sist + 2*diast) / 3",
    },
    "chads_vasc": {
        "grupo": "Cardiovascular", "titulo": "CHA2DS2-VASc",
        "campos": [("icc", "Insuficiencia cardiaca", "bool"),
                   ("hta", "Hipertension", "bool"),
                   ("edad75", "Edad >=75", "bool"),
                   ("diabetes", "Diabetes", "bool"),
                   ("ictus", "Ictus/AIT previo", "bool"),
                   ("vascular", "Enfermedad vascular", "bool"),
                   ("edad65", "Edad 65-74", "bool"),
                   ("sexo_f", "Sexo femenino", "bool")],
        "fn": "calc_chads_vasc",
        "bandas": [(0, 0, "Bajo riesgo de ictus", "ok"),
                   (1, 1, "Riesgo intermedio: valorar anticoagulacion", "alerta"),
                   (2, None, "Alto riesgo: se recomienda anticoagulacion "
                             "(varones >=2, mujeres >=3)", "critico")],
        "estado": "activa",
        "formula": "1 pt: ICC, HTA, DM, vascular, edad 65-74, sexo F - "
                   "2 pts: edad >=75, ictus/AIT",
    },
    "has_bled": {
        "grupo": "Cardiovascular", "titulo": "HAS-BLED",
        "campos": [("hta", "Hipertension (no controlada)", "bool"),
                   ("renal", "Funcion renal/hepatica anormal", "bool"),
                   ("ictus", "Ictus previo", "bool"),
                   ("sangrado", "Antecedente de sangrado", "bool"),
                   ("inr", "INR inestable", "bool"),
                   ("edad65", "Edad >65", "bool"),
                   ("farmacos", "Drogas/alcohol", "bool")],
        "fn": "calc_has_bled",
        "bandas": [(0, 2, "Riesgo hemorragico menor/intermedio", "ok"),
                   (3, None, "ALTO riesgo hemorragico: revisar anticoagulacion", "critico")],
        "estado": "activa",
        "formula": "1 pt por cada criterio presente; >=3 = alto riesgo",
    },
    "qtc": {
        "grupo": "Cardiovascular", "titulo": "QTc corregido (Bazett)",
        "campos": [("qt", "Intervalo QT (ms)", "num"),
                   ("rr", "Intervalo RR (ms)", "num")],
        "fn": "calc_qtc",
        "bandas": [(None, 440, "Normal", "ok"),
                   (441, 500, "Largo: vigilar", "alerta"),
                   (501, None, "ALTO riesgo de torsades de pointes", "critico")],
        "estado": "activa",
        "formula": "QTc = QT / sqrt(RR)",
    },

    # ---- II. RESPIRATORIO ----
    "curb65": {
        "grupo": "Respiratorio", "titulo": "CURB-65 (neumonia)",
        "campos": [("confusion", "Alteracion de conciencia", "bool"),
                   ("urea", "Urea >7 mmol/L", "bool"),
                   ("fr30", "FR >=30 rpm", "bool"),
                   ("hipotension", "Hipotension (TA <90/60)", "bool"),
                   ("edad65", "Edad >=65", "bool")],
        "fn": "calc_curb65",
        "bandas": [(0, 1, "Bajo riesgo: manejo ambulatorio posible", "ok"),
                   (2, 2, "Hospitalizacion sugerida", "alerta"),
                   (3, None, "Considerar UCI", "critico")],
        "estado": "activa",
        "formula": "1 pt por cada criterio presente",
    },
    "wells_tep": {
        "grupo": "Respiratorio", "titulo": "Criterios de Wells para TEP",
        "campos": [("signos_tvp", "Signos clinicos de TVP", "bool"),
                   ("tep_probable", "TEP como diagnostico mas probable", "bool"),
                   ("fc100", "FC >100", "bool"),
                   ("inmovilizacion", "Inmovilizacion/cirugia reciente", "bool"),
                   ("tep_previo", "TVP/TEP previo", "bool"),
                   ("hemoptisis", "Hemoptisis", "bool"),
                   ("neoplasia", "Neoplasia", "bool")],
        "fn": "calc_wells",
        "bandas": [(0, 4, "Probabilidad baja/intermedia: mas estudios", "alerta"),
                   (5, None, ">4 puntos: TEP probable, requiere estudios", "critico")],
        "estado": "activa",
        "formula": "Tabla estandar de puntos por criterio (ver fn)",
    },
    "perc": {
        "grupo": "Respiratorio", "titulo": "Regla PERC (exclusion de TEP)",
        "campos": [("edad50", "Edad <50", "bool"),
                   ("fc100", "FC <100", "bool"),
                   ("spo2_95", "SpO2 >=95%", "bool"),
                   ("sin_hemoptisis", "Sin hemoptisis", "bool"),
                   ("sin_estrogenos", "Sin uso de estrogenos", "bool"),
                   ("sin_tep_previo", "Sin TVP/TEP previo", "bool"),
                   ("sin_edema", "Sin edema unilateral de pierna", "bool")],
        "fn": "calc_perc",
        "bandas": [(1, 1, "PERC negativo: TEP se puede excluir con seguridad", "ok"),
                   (0, 0, "PERC positivo: no se excluye TEP", "alerta")],
        "estado": "activa",
        "formula": "Si TODOS los criterios de bajo riesgo se cumplen, PERC negativo",
    },

    # ---- III. RENAL / ELECTROLITOS ----
    "egfr_ckdepi": {
        "grupo": "Renal/Electrolitos", "titulo": "eGFR CKD-EPI 2021",
        "campos": [("creatinina", "Creatinina (mg/dL)", "num"),
                   ("edad", "Edad (anios)", "num"),
                   ("sexo", "Sexo (M/F)", "txt")],
        "fn": "calc_egfr",
        "bandas": [(None, 60, "<60 mL/min/1.73m2: sugerir ERC si >=3 meses", "alerta")],
        "estado": "activa",
        "formula": "CKD-EPI 2021 sin factor de raza (tablas oficiales en fn)",
    },
    "cockcroft": {
        "grupo": "Renal/Electrolitos", "titulo": "Depuracion de creatinina (Cockcroft-Gault)",
        "campos": [("edad", "Edad (anios)", "num"),
                   ("peso", "Peso (kg)", "num"),
                   ("creatinina", "Creatinina (mg/dL)", "num"),
                   ("sexo_f", "Sexo femenino", "bool")],
        "fn": "calc_cockcroft",
        "bandas": [],
        "estado": "activa",
        "formula": "[(140-edad)*peso] / (72*creatinina) * 0.85 si mujer",
    },
    "fena": {
        "grupo": "Renal/Electrolitos", "titulo": "Fraccion excretada de sodio (FENa)",
        "campos": [("na_urinario", "Na urinario (mmol/L)", "num"),
                   ("cr_serica", "Creatinina serica (mg/dL)", "num"),
                   ("na_serico", "Na serico (mmol/L)", "num"),
                   ("cr_urinaria", "Creatinina urinaria (mg/dL)", "num")],
        "fn": "calc_fena",
        "bandas": [(None, 1.0, "<1%: prerrenal", "ok"),
                   (1.0, 2.0, "1-2%: zona gris", "alerta"),
                   (2.0, None, ">2%: lesion renal (intrinsic)", "critico")],
        "estado": "activa",
        "formula": "(Na urine * Cr serica) / (Na serico * Cr urinaria) * 100",
    },
    "anion_gap": {
        "grupo": "Renal/Electrolitos", "titulo": "Anion gap",
        "campos": [("na", "Na (mmol/L)", "num"),
                   ("cl", "Cl (mmol/L)", "num"),
                   ("hco3", "HCO3 (mmol/L)", "num")],
        "fn": "calc_anion_gap",
        "bandas": [(None, 7.9, "Bajo", "alerta"),
                   (8.0, 12.0, "Normal", "ok"),
                   (12.1, None, "Elevado: acidosis cetonica/lactica/uremica", "alerta")],
        "estado": "activa",
        "formula": "Na - (Cl + HCO3)",
    },
    "na_corregido": {
        "grupo": "Renal/Electrolitos", "titulo": "Sodio corregido (hiperglucemia)",
        "campos": [("na", "Na medido (mmol/L)", "num"),
                   ("glucosa", "Glucosa (mg/dL)", "num")],
        "fn": "calc_na_corregido",
        "bandas": [],
        "estado": "activa",
        "formula": "Na corregido = Na + 0.016*(glucosa-100)",
    },

    # ---- IV. CUIDADOS CRITICOS / URGENCIAS ----
    "qsofa": {
        "grupo": "Cuidados criticos/Urgencias", "titulo": "qSOFA",
        "campos": [("fr", "FR (rpm)", "num"),
                   ("conciencia_alterada", "Alteracion de conciencia", "bool"),
                   ("ta_s", "TA sistolica (mmHg)", "num")],
        "fn": "score_qsofa",
        "bandas": [(0, 1, "Bajo riesgo", "ok"),
                   (2, 3, "MAL PRONOSTICO: >=2 indica riesgo elevado", "critico")],
        "estado": "activa",
        "formula": "1 pt: FR>=22, alteracion de conciencia, TA sist <=100",
    },
    "mews": {
        "grupo": "Cuidados criticos/Urgencias", "titulo": "MEWS",
        "campos": [("ta_s", "TA sistolica", "num"), ("fc", "FC", "num"),
                   ("fr", "FR", "num"), ("temp", "Temp", "num"),
                   ("avpu", "Nivel de conciencia AVPU", "txt")],
        "fn": "score_escala_incompleta",
        "bandas": [(5, None, ">=5: riesgo aumentado de muerte/UCI", "critico")],
        "estado": "incompleta",
        "formula": "Tabla de puntos por banda PENDIENTE del medico",
    },
    "news": {
        "grupo": "Cuidados criticos/Urgencias", "titulo": "NEWS / NEWS2",
        "campos": [("fr", "FR", "num"), ("spo2", "SpO2", "num"),
                   ("temp", "Temp", "num"), ("ta_s", "TA sist", "num"),
                   ("fc", "FC", "num"), ("conciencia", "Conciencia", "txt"),
                   ("oxigeno", "Recibe oxigeno", "bool")],
        "fn": "score_escala_incompleta",
        "bandas": [(5, 6, ">=5: evaluacion urgente", "alerta"),
                   (7, None, ">=7: considerar UCI", "critico")],
        "estado": "incompleta",
        "formula": "Tabla de puntos PENDIENTE del medico",
    },
    "gcs": {
        "grupo": "Cuidados criticos/Urgencias", "titulo": "Escala de Coma de Glasgow",
        "campos": [("apertura", "Apertura ocular (1-4)", "num"),
                   ("verbal", "Respuesta verbal (1-5)", "num"),
                   ("motora", "Respuesta motora (1-6)", "num")],
        "fn": "calc_gcs",
        "bandas": [(None, 8, "<=8: coma, considerar proteccion de via aerea", "critico"),
                   (9, 12, "Leve/moderado", "alerta"),
                   (13, 15, "Leve/normal", "ok")],
        "estado": "activa",
        "formula": "E + V + M",
    },

    # ---- V. HEPATICO / DIGESTIVO ----
    "meld_na": {
        "grupo": "Hepatico/Digestivo", "titulo": "MELD-Na",
        "campos": [("bilirrubina", "Bilirrubina (mg/dL)", "num"),
                   ("inr", "INR", "num"),
                   ("creatinina", "Creatinina (mg/dL)", "num"),
                   ("na", "Na (mmol/L)", "num")],
        "fn": "calc_meld_na",
        "bandas": [(20, None, ">20: alto riesgo de muerte a corto plazo", "critico")],
        "estado": "activa",
        "formula": "3.78*ln(bili) + 11.2*ln(INR) + 9.57*ln(cr) + 6.43 + "
                   "1.32*ln(Na) - 0.033*MELD*Na",
    },
    "child_pugh": {
        "grupo": "Hepatico/Digestivo", "titulo": "Child-Pugh",
        "campos": [("bilirrubina", "Bilirrubina", "num"),
                   ("albunina", "Albunina", "num"),
                   ("inr", "INR", "num"),
                   ("ascitis", "Ascitis (0-3)", "num"),
                   ("encefalopatia", "Encefalopatia (0-3)", "num")],
        "fn": "calc_child_pugh",
        "bandas": [(5, 6, "Clase A (5-6): mejor pronostico", "ok"),
                   (7, 9, "Clase B (7-9): pronostico intermedio", "alerta"),
                   (10, 15, "Clase C (10-15): mal pronostico", "critico")],
        "estado": "activa",
        "formula": "Suma de 5 items con tabla clasica de puntos 1-3 por item",
    },
    "fib4": {
        "grupo": "Hepatico/Digestivo", "titulo": "FIB-4 (fibrosis hepatica)",
        "campos": [("edad", "Edad (anios)", "num"),
                   ("ast", "AST (U/L)", "num"),
                   ("plaquetas", "Plaquetas (x10^9/L)", "num"),
                   ("alt", "ALT (U/L)", "num")],
        "fn": "calc_fib4",
        "bandas": [(None, 1.3, "Bajo: fibrosis improbable", "ok"),
                   (1.31, 2.67, "Zona gris: confirmar", "alerta"),
                   (2.68, None, ">2.67: sugiere fibrosis significativa", "critico")],
        "estado": "activa",
        "formula": "(edad * AST) / (plaquetas * sqrt(ALT))",
    },

    # ---- VI. NEUROLOGICO / PSIQUIATRICO ----
    "phq9": {
        "grupo": "Neurologico/Psiquiatrico", "titulo": "PHQ-9 (depresion)",
        "campos": [("puntaje", "Puntaje total del cuestionario (0-27)", "num")],
        "fn": "calc_phq9",
        "bandas": [(0, 9, "Sin o depresion minima", "ok"),
                   (10, 19, "Depresion moderada", "alerta"),
                   (20, None, "Depresion grave", "critico")],
        "estado": "activa",
        "formula": "Tamizaje: el medico aplica el cuestionario y captura el total",
    },
    "gad7": {
        "grupo": "Neurologico/Psiquiatrico", "titulo": "GAD-7 (ansiedad)",
        "campos": [("puntaje", "Puntaje total del cuestionario (0-21)", "num")],
        "fn": "calc_gad7",
        "bandas": [(0, 9, "Sin o ansiedad minima", "ok"),
                   (10, None, ">=10: ansiedad moderada", "alerta")],
        "estado": "activa",
        "formula": "Tamizaje: el medico aplica el cuestionario y captura el total",
    },

    # ---- VII. OBSTETRICIA / PEDIATRIA ----
    "bishop": {
        "grupo": "Obstetricia/Pediatria", "titulo": "Bishop Score",
        "campos": [("dilatacion", "Dilatacion cervical (cm)", "num"),
                   ("borramiento", "Borramiento (%)", "num"),
                   ("posicion", "Posicion (0-3)", "num"),
                   ("consistencia", "Consistencia (0-2)", "num"),
                   ("altura", "Altura de presentacion (0-3)", "num")],
        "fn": "calc_bishop",
        "bandas": [(8, None, ">=8: alta probabilidad de exito de induccion", "ok"),
                   (0, 7, "<8: induccion con menor probabilidad de exito", "alerta")],
        "estado": "activa",
        "formula": "Suma de 5 items con tabla obstetrica clasica",
    },
    "apgar": {
        "grupo": "Obstetricia/Pediatria", "titulo": "Apgar (neonato)",
        "campos": [("apariencia", "Apariencia (0-2)", "num"),
                   ("pulso", "Pulso (0-2)", "num"),
                   ("reflejo", "Reflejos/irritabilidad (0-2)", "num"),
                   ("actividad", "Actividad muscular (0-2)", "num"),
                   ("respiracion", "Respiracion (0-2)", "num")],
        "fn": "calc_apgar",
        "bandas": [(0, 3, "<=3: asfixia grave, reanimacion inmediata", "critico"),
                   (4, 6, "Depresion moderada: estimulacion", "alerta"),
                   (7, 10, "Normal", "ok")],
        "estado": "activa",
        "formula": "Suma de 5 items (0-2 cada uno)",
    },

    # ---- VIII. GENERALES / COMPOSICION CORPORAL ----
    "imc": {
        "grupo": "Generales/Composicion corporal", "titulo": "IMC",
        "campos": [("peso", "Peso (kg)", "num"),
                   ("talla", "Talla (m)", "num")],
        "fn": "calc_imc",
        "bandas": [(None, 18.49, "Bajo peso (<18.5)", "alerta"),
                   (18.5, 24.99, "Normal", "ok"),
                   (25.0, 29.99, "Sobrepeso", "alerta"),
                   (30.0, None, "Obesidad (>=30)", "critico")],
        "estado": "activa",
        "formula": "IMC = peso / talla^2",
    },
    "sc_mosteller": {
        "grupo": "Generales/Composicion corporal", "titulo": "Superficie corporal (Mosteller)",
        "campos": [("talla", "Talla (cm)", "num"),
                   ("peso", "Peso (kg)", "num")],
        "fn": "calc_sc_mosteller",
        "bandas": [],
        "estado": "activa",
        "formula": "SC = sqrt((talla * peso) / 3600)",
    },
    "peso_ideal": {
        "grupo": "Generales/Composicion corporal", "titulo": "Peso ideal (Devine)",
        "campos": [("talla_in", "Talla (pulgadas)", "num"),
                   ("sexo_f", "Sexo femenino", "bool")],
        "fn": "calc_peso_ideal",
        "bandas": [],
        "estado": "activa",
        "formula": "M: 50 + 2.3*(pulg-60) - F: 45.5 + 2.3*(pulg-60)",
    },
}

# Grupos en el orden del documento del medico (para la UI C-19).
# ADR-0043: "Generales/Composicion corporal" encabeza por prioridad de
# workflow del dueno (2026-10-16) — modulo de mayor uso medico/enfermera.
GRUPOS_CALCULADORAS = [
    "Generales/Composicion corporal",
    "Cardiovascular",
    "Respiratorio",
    "Renal/Electrolitos",
    "Cuidados criticos/Urgencias",
    "Hepatico/Digestivo",
    "Neurologico/Psiquiatrico",
    "Obstetricia/Pediatria",
]

# ADR-0043: grupo virtual de la pagina C-19 con las calculadoras que el
# consultorio define como de mayor valor para su flujo rapido (se guardan
# en _CONFIG.json "calc_prioritarias"; aqui solo va el defecto de referencia).
GRUPO_PRIORITARIAS = "★ Prioritarias (flujo rapido)"
CALC_PRIORITARIAS_DEFECTO = ["imc", "sc_mosteller", "peso_ideal"]


# ============================================================================
# 5. FUNCIONES PURAS DE LAS CALCULADORAS "activa"
#    (transcribir a app.py seccion 08d en T15; sin dependencias)
# ============================================================================
import math


def _banda(calc_id, valor):
    """Devuelve (etiqueta, nivel) de la banda que contiene el valor."""
    if valor is None:
        return "", ""
    for lo, hi, etiqueta, nivel in CALCULADORAS[calc_id]["bandas"]:
        ok_lo = lo is None or valor >= lo
        ok_hi = hi is None or valor <= hi
        if ok_lo and ok_hi:
            return etiqueta, nivel
    return "", ""


def _resultado(calc_id, valor):
    etiqueta, nivel = _banda(calc_id, valor)
    return {"valor": valor, "etiqueta": etiqueta, "nivel": nivel}


def calc_pam(c):
    try:
        r = (c["ta_s"] + 2 * c["ta_d"]) / 3.0
    except (KeyError, TypeError):
        return None, "PAM incompleto: faltan TA sistolica/diastolica."
    r2 = round(r, 1)
    return _resultado("pam", r2), f"PAM = {r2:g} mmHg."


def calc_chads_vasc(c):
    pts = 0
    pts += 1 if c.get("icc") else 0
    pts += 1 if c.get("hta") else 0
    pts += 2 if c.get("edad75") else 0
    pts += 1 if c.get("diabetes") else 0
    pts += 2 if c.get("ictus") else 0
    pts += 1 if c.get("vascular") else 0
    pts += 1 if c.get("edad65") else 0
    pts += 1 if c.get("sexo_f") else 0
    return _resultado("chads_vasc", pts), f"CHA2DS2-VASc = {pts} puntos."


def calc_has_bled(c):
    claves = ("hta", "renal", "ictus", "sangrado", "inr", "edad65", "farmacos")
    pts = sum(1 for k in claves if c.get(k))
    return _resultado("has_bled", pts), f"HAS-BLED = {pts} puntos."


def calc_qtc(c):
    try:
        r = c["qt"] / math.sqrt(c["rr"])
    except (KeyError, TypeError, ValueError, ZeroDivisionError):
        return None, "QTc incompleto: faltan QT o RR."
    r2 = round(r)
    return _resultado("qtc", r2), f"QTc (Bazett) = {r2} ms."


def calc_curb65(c):
    claves = ("confusion", "urea", "fr30", "hipotension", "edad65")
    pts = sum(1 for k in claves if c.get(k))
    return _resultado("curb65", pts), f"CURB-65 = {pts} puntos."


def calc_wells(c):
    # Tabla estandar: signos TVP 3, TEP mas probable 3, FC>100 1.5,
    # inmovilizacion/cirugia 1.5, TVP/TEP previo 1.5, hemoptisis 1, neoplasia 1
    pts = 0.0
    if c.get("signos_tvp"):
        pts += 3
    if c.get("tep_probable"):
        pts += 3
    if c.get("fc100"):
        pts += 1.5
    if c.get("inmovilizacion"):
        pts += 1.5
    if c.get("tep_previo"):
        pts += 1.5
    if c.get("hemoptisis"):
        pts += 1
    if c.get("neoplasia"):
        pts += 1
    etiqueta, nivel = _banda("wells_tep", pts)
    return {"valor": pts, "etiqueta": etiqueta, "nivel": nivel}, \
        f"Wells TEP = {pts:g} puntos."


def calc_perc(c):
    criterios = ("edad50", "fc100", "spo2_95", "sin_hemoptisis",
                 "sin_estrogenos", "sin_tep_previo", "sin_edema")
    cumplidos = sum(1 for k in criterios if c.get(k))
    negativo = cumplidos == len(criterios)
    etiqueta, nivel = _banda("perc", 1 if negativo else 0)
    return {"valor": 1 if negativo else 0, "etiqueta": etiqueta, "nivel": nivel}, \
        ("PERC negativo: TEP excluible." if negativo
         else "PERC positivo: no se excluye TEP.")


def calc_egfr(c):
    """CKD-EPI 2021 (sin raza). kappa=0.7 mujeres, 1.0 hombres."""
    try:
        scr = float(c["creatinina"])
        edad = float(c["edad"])
    except (KeyError, TypeError, ValueError):
        return None, "eGFR incompleto: faltan creatinina o edad."
    mujer = str(c.get("sexo", "")).strip().upper() in ("F", "MUJER", "FEMENINO")
    kappa = 0.7 if mujer else 1.0
    alpha = -0.241 if mujer else -0.302
    if scr <= kappa:
        r = 142 * (scr / kappa) ** alpha * (0.9938 ** edad) * (1.012 if mujer else 1.0)
    else:
        r = 142 * (scr / kappa) ** -1.200 * (0.9938 ** edad) * (1.012 if mujer else 1.0)
    r2 = round(r, 1)
    return _resultado("egfr_ckdepi", r2), f"eGFR (CKD-EPI 2021) = {r2:g} mL/min/1.73m2."


def calc_cockcroft(c):
    try:
        edad = float(c["edad"])
        peso = float(c["peso"])
        scr = float(c["creatinina"])
    except (KeyError, TypeError, ValueError):
        return None, "Cockcroft-Gault incompleto: faltan edad, peso o creatinina."
    if scr <= 0:
        return None, "Creatinina invalida."
    r = ((140 - edad) * peso) / (72 * scr)
    if c.get("sexo_f"):
        r *= 0.85
    r2 = round(r, 1)
    return _resultado("cockcroft", r2), f"Depuracion estimada = {r2:g} mL/min."


def calc_fena(c):
    try:
        na_u = float(c["na_urinario"])
        cr_s = float(c["cr_serica"])
        na_s = float(c["na_serico"])
        cr_u = float(c["cr_urinaria"])
    except (KeyError, TypeError, ValueError):
        return None, "FENa incompleta: faltan los 4 valores."
    if na_s <= 0 or cr_u <= 0:
        return None, "Valores invalidos para FENa."
    r = (na_u * cr_s) / (na_s * cr_u) * 100.0
    r2 = round(r, 2)
    return _resultado("fena", r2), f"FENa = {r2:g}%."


def calc_anion_gap(c):
    try:
        r = float(c["na"]) - (float(c["cl"]) + float(c["hco3"]))
    except (KeyError, TypeError, ValueError):
        return None, "Anion gap incompleto: faltan Na, Cl o HCO3."
    r2 = round(r, 1)
    return _resultado("anion_gap", r2), f"Anion gap = {r2:g} mEq/L."


def calc_na_corregido(c):
    try:
        r = float(c["na"]) + 0.016 * (float(c["glucosa"]) - 100)
    except (KeyError, TypeError, ValueError):
        return None, "Na corregido incompleto: faltan Na o glucosa."
    r2 = round(r, 1)
    return _resultado("na_corregido", r2), f"Na corregido = {r2:g} mmol/L."


def calc_gcs(c):
    try:
        e = int(c["apertura"])
        v = int(c["verbal"])
        m = int(c["motora"])
    except (KeyError, TypeError, ValueError):
        return None, "GCS incompleto: faltan apertura, verbal o motora."
    if not (1 <= e <= 4 and 1 <= v <= 5 and 1 <= m <= 6):
        return None, "GCS fuera de rango (E 1-4, V 1-5, M 1-6)."
    total = e + v + m
    return _resultado("gcs", total), f"GCS = {total} (E{e} V{v} M{m})."


def calc_meld_na(c):
    try:
        bili = max(float(c["bilirrubina"]), 1.0)
        inr = max(float(c["inr"]), 1.0)
        cr = max(float(c["creatinina"]), 1.0)
        na = min(max(float(c["na"]), 125.0), 140.0)
    except (KeyError, TypeError, ValueError):
        return None, "MELD-Na incompleto: faltan bilirrubina, INR, creatinina o Na."
    meld = (3.78 * math.log(bili) + 11.2 * math.log(inr)
            + 9.57 * math.log(cr) + 6.43)
    r = meld + 1.32 * math.log(na) - 0.033 * meld * (na - 40.0) - 6.43
    r2 = round(r)
    return _resultado("meld_na", r2), f"MELD-Na = {r2} puntos."


def calc_child_pugh(c):
    """Tabla clasica: cada item 1-3 puntos; el medico captura el puntaje por item."""
    try:
        bili = float(c["bilirrubina"])
        alb = float(c["albunina"])
        inr = float(c["inr"])
        asc = int(c["ascitis"])
        enc = int(c["encefalopatia"])
    except (KeyError, TypeError, ValueError, KeyError):
        return None, "Child-Pugh incompleto: faltan los 5 valores."
    # Puntos de laboratorio por tabla clasica
    p_bili = 1 if bili < 2 else (2 if bili < 3 else 3)
    p_alb = 1 if alb > 3.5 else (2 if alb > 2.8 else 3)
    p_inr = 1 if inr < 1.7 else (2 if inr < 2.3 else 3)
    if not (1 <= asc <= 3 and 1 <= enc <= 3):
        return None, "Ascitis y encefalopatia deben ser 1-3."
    total = p_bili + p_alb + p_inr + asc + enc
    return _resultado("child_pugh", total), f"Child-Pugh = {total} puntos."


def calc_fib4(c):
    try:
        r = (float(c["edad"]) * float(c["ast"])) / \
            (float(c["plaquetas"]) * math.sqrt(float(c["alt"])))
    except (KeyError, TypeError, ValueError, ZeroDivisionError):
        return None, "FIB-4 incompleto: faltan edad, AST, plaquetas o ALT."
    r2 = round(r, 2)
    return _resultado("fib4", r2), f"FIB-4 = {r2:g}."


def calc_phq9(c):
    try:
        p = int(c["puntaje"])
    except (KeyError, TypeError, ValueError):
        return None, "PHQ-9: capture el puntaje total del cuestionario."
    if not (0 <= p <= 27):
        return None, "PHQ-9 fuera de rango (0-27)."
    return _resultado("phq9", p), f"PHQ-9 = {p}."


def calc_gad7(c):
    try:
        p = int(c["puntaje"])
    except (KeyError, TypeError, ValueError):
        return None, "GAD-7: capture el puntaje total del cuestionario."
    if not (0 <= p <= 21):
        return None, "GAD-7 fuera de rango (0-21)."
    return _resultado("gad7", p), f"GAD-7 = {p}."


def calc_bishop(c):
    """Suma directa: el medico captura cada item con su puntaje de la tabla."""
    try:
        total = (int(c["dilatacion"]) + int(c["borramiento"])
                 + int(c["posicion"]) + int(c["consistencia"]) + int(c["altura"]))
    except (KeyError, TypeError, ValueError):
        return None, "Bishop incompleto: faltan los 5 items."
    return _resultado("bishop", total), f"Bishop = {total} puntos."


def calc_apgar(c):
    try:
        total = (int(c["apariencia"]) + int(c["pulso"]) + int(c["reflejo"])
                 + int(c["actividad"]) + int(c["respiracion"]))
    except (KeyError, TypeError, ValueError):
        return None, "Apgar incompleto: faltan los 5 items."
    if not all(0 <= int(c[k]) <= 2 for k in
               ("apariencia", "pulso", "reflejo", "actividad", "respiracion")):
        return None, "Cada item de Apgar va de 0 a 2."
    return _resultado("apgar", total), f"Apgar = {total}/10."


def calc_imc(c):
    try:
        r = float(c["peso"]) / (float(c["talla"]) ** 2)
    except (KeyError, TypeError, ValueError, ZeroDivisionError):
        return None, "IMC incompleto: faltan peso o talla."
    r2 = round(r, 1)
    return _resultado("imc", r2), f"IMC = {r2:g} kg/m2."


def calc_sc_mosteller(c):
    try:
        r = math.sqrt((float(c["talla"]) * float(c["peso"])) / 3600.0)
    except (KeyError, TypeError, ValueError):
        return None, "SC incompleto: faltan talla o peso."
    r2 = round(r, 2)
    return _resultado("sc_mosteller", r2), f"SC (Mosteller) = {r2:g} m2."


def calc_peso_ideal(c):
    try:
        pulg = float(c["talla_in"])
    except (KeyError, TypeError, ValueError):
        return None, "Peso ideal incompleto: falta la talla en pulgadas."
    r = (45.5 if c.get("sexo_f") else 50.0) + 2.3 * (pulg - 60.0)
    r2 = round(r, 1)
    return _resultado("peso_ideal", r2), f"Peso ideal (Devine) = {r2:g} kg."


# Mapa id -> funcion (se transcribe igual en app.py)
FUNCIONES_CALC = {
    "pam": calc_pam,
    "chads_vasc": calc_chads_vasc,
    "has_bled": calc_has_bled,
    "qtc": calc_qtc,
    "curb65": calc_curb65,
    "wells_tep": calc_wells,
    "perc": calc_perc,
    "egfr_ckdepi": calc_egfr,
    "cockcroft": calc_cockcroft,
    "fena": calc_fena,
    "anion_gap": calc_anion_gap,
    "na_corregido": calc_na_corregido,
    "qsofa": score_qsofa,
    "mews": score_escala_incompleta,
    "news": score_escala_incompleta,
    "gcs": calc_gcs,
    "meld_na": calc_meld_na,
    "child_pugh": calc_child_pugh,
    "fib4": calc_fib4,
    "phq9": calc_phq9,
    "gad7": calc_gad7,
    "bishop": calc_bishop,
    "apgar": calc_apgar,
    "imc": calc_imc,
    "sc_mosteller": calc_sc_mosteller,
    "peso_ideal": calc_peso_ideal,
}


def calcular(id_calc, entradas):
    """Punto de entrada uniforme: (id, campos) -> (ok, resultado_o_None, mensaje).

    ok=False cuando la calculadora esta incompleta o las entradas son invalidas.
    """
    calc = CALCULADORAS.get(id_calc)
    if not calc:
        return False, None, f"Calculadora desconocida: {id_calc}."
    fn = FUNCIONES_CALC.get(id_calc)
    if not fn or calc.get("estado") != "activa":
        return False, None, (f"Calculadora {calc['titulo']} incompleta: "
                             f"pendiente tabla/formula del medico.")
    res, msg = fn(entradas or {})
    if res is None:
        return False, None, msg
    return True, res, msg


# ============================================================================
# Validacion rapida: python3 esquemas/catalogos_clinicos.py
# ============================================================================
if __name__ == "__main__":
    # Signos vitales
    h = evaluar_signos_vitales({"ta_s": 150, "fc": 130, "fr": 25, "spo2": 92},
                               perfil="embarazada")
    assert any(x["campo"] == "ta_s" for x in h), "HTA no detectada"
    assert any(x["nivel"] == "esperado_embarazo" for x in h), "ajuste embarazo"
    # Preeclampsia
    p = evaluar_preeclampsia({"ta_s": 165, "ta_d": 112}, "2+")
    assert len(p) == 1, "preeclampsia no detectada"
    assert len(evaluar_preeclampsia({"ta_s": 120, "ta_d": 80}, "negativa")) == 0
    # Lab
    assert evaluar_lab("potasio", 7.0)["nivel"] == "critico"
    assert evaluar_lab("glucosa", 100) is None
    # qSOFA
    pts, _ = score_qsofa({"fr": 24, "ta_s": 95, "conciencia_alterada": True})
    assert pts == 3
    # Calculadoras
    ok, res, _ = calcular("imc", {"peso": 70, "talla": 1.65})
    assert ok and abs(res["valor"] - 25.7) < 0.1
    ok, res, _ = calcular("pam", {"ta_s": 120, "ta_d": 80})
    assert ok and abs(res["valor"] - 93.3) < 0.1
    ok, _, msg = calcular("mews", {})
    assert not ok and "incompleta" in msg.lower()
    n_act = sum(1 for v in CALCULADORAS.values() if v["estado"] == "activa")
    n_inc = sum(1 for v in CALCULADORAS.values() if v["estado"] == "incompleta")
    # ADR-0043: composicion corporal encabeza y el defecto de prioritarias es
    # solo de ese modulo, con ids existentes
    assert GRUPOS_CALCULADORAS[0] == "Generales/Composicion corporal"
    assert GRUPO_PRIORITARIAS not in GRUPOS_CALCULADORAS
    assert all(c in CALCULADORAS for c in CALC_PRIORITARIAS_DEFECTO)
    print(f"OK: {len(CALCULADORAS)} calculadoras ({n_act} activas, {n_inc} incompletas), "
          f"{len(UMBRALES_LAB)} panic values, {len(RANGOS_SIGNOS_VITALES)} perfiles de rangos.")
