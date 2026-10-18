# -*- coding: utf-8 -*-
"""
ESQUEMA NOM-004 COMPLETO — SECCIONES A–M (fuente de verdad del checklist)
=========================================================================
Plantilla de DATOS para transcribir a `app.py` (seccion 08b SERVICIO HISTORIA)
durante el paquete T12 (ADR-0027). NO se importa desde app.py: el monolito
sigue stdlib-only y autocontenido (ADR-0002). Este archivo es la CANONICA:
si se divergen, manda aqui y se re-transcribe.

FORMATO DE ITEM:
  (id, tipo, etiqueta[, ops])
    tipo:
      "auto"  derivado del expediente, solo lectura (cero tecleo)
      "texto" texto libre
      "numero" numerico (coercion _a_numero)
      "fecha" ISO via parsear_fecha
      "enum"  opciones cerradas (ops)
      "sn"    Sí / No / (vacio = sin investigar)
      "snns"  Sí / No / No sabe
      "snna"  Sí / No / No aplica
      "snv"   Sí / No + campo "valor" libre (signos vitales, dolor)
      "snc"   Sí / No + extras (E1 edad/tratamiento; E4/E5/F7 especificar; F4 habito)
      "snh"   Sí / No + campo "hallazgo" (exploracion por regiones)
      "libre" narrativa clinica del medico (una por seccion)

ESTADO TRI-STATE ALMACENADO: "si" | "no" | "no_sabe" | "no_aplica" | ""
  "" (vacio) = SIN INVESTIGAR — distinto de "no" (omision auditable).

NIVEL:
  "paciente"  → HISTORIA.json["paciente"][seccion]  (unica, acumulativa)
  "nota"      → HISTORIA.json["notas"][i][seccion]  (ingreso / evolucion)
  "egreso"    → HISTORIA.json["notas"][i] tipo=egreso

TRANSCRIPCION FIEL del checklist del dueno (2026-10-16). Sin inventar nada.
"""

from datetime import date


def _item(id_, tipo, etiqueta, ops=None):
    it = {"id": id_, "t": tipo, "lbl": etiqueta}
    if ops:
        it["ops"] = ops
    return it


def _sn(*pares):
    """Expande (id, etiqueta) a items tipo sn."""
    return [_item(i, "sn", l) for i, l in pares]


def _snns(*pares):
    return [_item(i, "snns", l) for i, l in pares]


# ============================================================================
# SECCION A — DATOS DE IDENTIFICACION Y FILIACION (nivel paciente)
# ============================================================================
SECC_A = [
    _item("nombre", "auto", "Nombre completo registrado"),
    _item("edad", "auto", "Edad cumplida (anios/meses/dias segun corresponda)"),
    _item("sexo", "enum", "Sexo registrado", ["M", "F", "OTRO"]),
    _item("fecha_nac", "auto", "Fecha de nacimiento registrada"),
    _item("lugar_nac", "texto", "Lugar de nacimiento registrado"),
    _item("nacionalidad", "texto", "Nacionalidad registrada"),
    _item("domicilio", "texto", "Domicilio completo registrado"),
    _item("telefono", "auto", "Telefono de contacto registrado"),
    _item("expediente", "auto", "Numero de expediente asignado"),
    _item("nss", "texto", "Numero de Seguridad Social (NSS)"),
    _item("curp", "auto", "CURP / identificacion oficial registrada"),
    _item("estado_civil", "enum", "Estado civil registrado",
          ["SOLTERO(A)", "CASADO(A)", "UNION LIBRE", "VIUDO(A)", "DIVORCIADO(A)"]),
    _item("ocupacion", "texto", "Ocupacion registrada"),
    _item("religion", "texto", "Religion registrada"),
    _item("escolaridad", "texto", "Escolaridad registrada"),
    _item("acompanante", "texto", "Nombre del acompanante o responsable (si aplica)"),
]

# ============================================================================
# SECCION B — MOTIVO DE LA ATENCION (nivel nota)
# ============================================================================
SECC_B = [
    _item("motivo", "texto", "Motivo de la atencion expresado por la paciente"),
    _item("fecha_hora", "auto", "Fecha y hora de la atencion"),
    _item("sintoma_principal", "texto", "Sintoma principal identificado"),
    _item("tiempo_evolucion", "texto", "Tiempo de evolucion del sintoma principal"),
    _item("tratamientos_previos", "texto", "Tratamientos previos investigados"),
    _item("automedicacion", "texto", "Automedicacion investigada"),
    _item("expectativa", "texto", "Expectativa de la paciente sobre la consulta"),
]

# ============================================================================
# SECCION C1 — ENFERMEDAD ACTUAL: INICIO Y EVOLUCION (nivel nota)
# ============================================================================
SECC_C1 = [
    _item("forma_inicio", "enum", "Forma de inicio",
          ["SUBITO", "GRADUAL", "NO PRECISA"]),
    _item("curso", "enum", "Curso",
          ["PROGRESIVO", "INTERMITENTE", "EN CRISIS", "ESTACIONARIO", "NO PRECISA"]),
    _item("duracion_total", "texto", "Duracion total del padecimiento"),
    _item("desencadenantes", "texto", "Factores desencadenantes identificados"),
    _item("agravantes", "texto", "Factores agravantes identificados"),
    _item("atenuantes", "texto", "Factores atenuantes identificados"),
    _item("acompanantes", "texto", "Sintomas acompanantes investigados"),
    _item("estado_actual", "texto", "Estado actual del padecimiento"),
]

# ============================================================================
# SECCION C2 — CARACTERIZACION DEL DOLOR (si aplica; nivel nota)
# ============================================================================
SECC_C2 = [
    _item("localizacion", "snv", "Localizacion precisa del dolor"),
    _item("irradiacion", "snv", "Irradiacion del dolor"),
    _item("caracter", "snv", "Tipo o caracter (punzante, ardor, colico, sordo)"),
    _item("intensidad", "snv", "Intensidad (escala 1-10)"),
    _item("duracion_episodio", "snv", "Duracion de cada episodio de dolor"),
    _item("frecuencia", "snv", "Frecuencia de los episodios"),
    _item("relacion_alimentos", "snv", "Relacion con alimentos, movimiento, respiracion"),
    _item("relacion_ciclo", "snv", "Relacion con el ciclo menstrual"),
]

# ============================================================================
# SECCION C3 — SINTOMAS GENERALES ASOCIADOS (nivel nota)
# ============================================================================
SECC_C3 = _sn(
    ("fiebre", "Fiebre"),
    ("escalofrios", "Escalofrios"),
    ("sudoracion_nocturna", "Sudoracion nocturna"),
    ("perdida_peso", "Perdida de peso involuntaria"),
    ("aumento_peso", "Aumento de peso involuntario"),
    ("astenia", "Astenia (cansancio)"),
    ("adinamia", "Adinamia (falta de fuerza)"),
    ("anorexia", "Anorexia (falta de apetito)"),
    ("polidipsia", "Polidipsia (sed excesiva)"),
    ("poliuria", "Poliuria (orina excesiva)"),
    ("polifagia", "Polifagia (hambre excesiva)"),
    ("insomnio", "Insomnio"),
    ("somnolencia", "Somnolencia excesiva"),
)

# ============================================================================
# SECCION D — ANTECEDENTES HEREDOFAMILIARES (nivel paciente)
# ============================================================================
SECC_D = _snns(
    ("dm", "Diabetes mellitus en familiares"),
    ("hta", "Hipertension arterial en familiares"),
    ("cardiopatias", "Cardiopatias en familiares"),
    ("cancer", "Cancer en familiares"),
    ("renal", "Enfermedad renal en familiares"),
    ("tiroidea", "Enfermedad tiroidea en familiares"),
    ("epilepsia", "Epilepsia en familiares"),
    ("asma", "Asma en familiares"),
    ("obesidad", "Obesidad en familiares"),
    ("mental", "Enfermedad mental en familiares"),
    ("tb", "Tuberculosis en familiares"),
    ("alcoholismo", "Alcoholismo en familiares"),
    ("toxicomanias", "Toxicomanias en familiares"),
    ("malformaciones", "Malformaciones congenitas en familiares"),
    ("abortos_recurrentes", "Abortos recurrentes en familiares"),
    ("consanguinidad", "Consanguinidad entre padres"),
    ("muerte_desconocida", "Muerte de familiar por causa desconocida"),
) + [
    _item("edad_fallecimiento_padres", "snv",
          "Edad de fallecimiento de padres registrada"),
    _item("causa_muerte_padres", "snv",
          "Causa de muerte de padres registrada"),
]

# ============================================================================
# SECCION E1 — ENFERMEDADES PREVIAS (nivel paciente; estado + edad dx + tratamiento)
# ============================================================================
_E1 = [
    ("dm", "Diabetes mellitus"),
    ("hta", "Hipertension arterial"),
    ("cardiopatia", "Cardiopatia"),
    ("asma", "Asma"),
    ("epoc", "EPOC"),
    ("tb", "Tuberculosis"),
    ("epilepsia", "Epilepsia"),
    ("tiroidea", "Enfermedad tiroidea"),
    ("renal", "Enfermedad renal"),
    ("hepatica", "Enfermedad hepatica"),
    ("cancer", "Cancer"),
    ("ets", "Enfermedad de transmision sexual"),
    ("vih", "VIH/SIDA"),
    ("hepatitis", "Hepatitis B o C"),
    ("mental", "Enfermedad mental"),
    ("gastritis", "Gastritis o ulcera"),
    ("colon", "Colon irritable"),
    ("artritis", "Artritis o artrosis"),
    ("osteoporosis", "Osteoporosis"),
    ("anemia", "Anemia"),
    ("coagulopatia", "Coagulopatia"),
    ("otra_cronica", "Otra enfermedad cronica"),
]
SECC_E1 = [
    _item(i, "snc", l + "  [estado | edad de diagnostico | tratamiento actual]")
    for i, l in _E1
]

# ============================================================================
# SECCION E2 — ANTECEDENTES QUIRURGICOS (nivel paciente)
# ============================================================================
SECC_E2 = [
    _item("operada", "sn", "Ha sido operada alguna vez"),
    _item("tipo_cirugia", "snv", "Tipo de cirugia registrado"),
    _item("fecha_cirugia", "snv", "Fecha de cirugia registrada"),
    _item("complicaciones", "snv", "Complicaciones quirurgicas investigadas"),
    _item("transfusiones", "sn", "Transfusiones sanguineas recibidas"),
    _item("reaccion_transfusional", "sn", "Reaccion transfusional previa"),
]

# ============================================================================
# SECCION E3 — ANTECEDENTES TRAUMATICOS (nivel paciente)
# ============================================================================
SECC_E3 = _sn(
    ("fracturas", "Fracturas previas"),
    ("perdida_conciencia", "Perdida de conciencia por trauma"),
    ("tce", "Traumatismo craneoencefalico previo"),
    ("heridas_sutura", "Heridas que requirieron sutura"),
    ("quemaduras", "Quemaduras previas"),
)

# ============================================================================
# SECCION E4 — ALERGIAS (nivel paciente; estado + especificar)
# ============================================================================
SECC_E4 = [
    _item(i, "snc", l + "  [estado | especificar]")
    for i, l in [
        ("medicamentos", "Alergia a medicamentos"),
        ("alimentos", "Alergia a alimentos"),
        ("insectos", "Alergia a picaduras de insectos"),
        ("ambientales", "Alergia a factores ambientales"),
        ("metales_latex", "Alergia a metales o latex"),
        ("anafilaxia", "Reaccion alergica grave previa (anafilaxia)"),
    ]
]

# ============================================================================
# SECCION E5 — INMUNIZACIONES (nivel paciente; estado + ultima dosis)
# ============================================================================
SECC_E5 = [
    _item(i, "snc", l + "  [si/no/no sabe | ultima dosis]")
    for i, l in [
        ("bcg", "BCG"),
        ("hep_b", "Hepatitis B"),
        ("influenza", "Influenza (ultimo anio)"),
        ("covid", "COVID-19"),
        ("tetanos", "Tetanos"),
        ("sarampion", "Sarampion/Rubeola"),
        ("vph", "VPH"),
        ("neumococo", "Neumococo"),
        ("otra", "Otra vacuna"),
    ]
]

# ============================================================================
# SECCION F1 — VIVIENDA (nivel paciente)
# ============================================================================
SECC_F1 = [
    _item("zona", "enum", "Zona de residencia", ["URBANA", "RURAL"]),
    _item("tipo_vivienda", "enum", "Tipo de vivienda",
          ["CASA", "DEPARTAMENTO", "CUARTO", "RANCHO", "OTRO"]),
    _item("habitantes", "numero", "Numero de habitantes en la vivienda"),
    _item("cuartos", "numero", "Numero de cuartos"),
    _item("hacinamiento", "sn", "Hacinamiento identificado"),
    _item("agua_potable", "sn", "Agua potable en la vivienda"),
    _item("drenaje", "sn", "Drenaje en la vivienda"),
    _item("electricidad", "sn", "Electricidad en la vivienda"),
    _item("animales", "sn", "Convivencia con animales domesticos"),
    _item("humo_lena", "sn", "Exposicion a humo de lena o carbon"),
]

# ============================================================================
# SECCION F2 — HABITOS ALIMENTICIOS (nivel paciente)
# ============================================================================
SECC_F2 = [
    _item("comidas_dia", "numero", "Numero de comidas al dia"),
    _item("tipo_alimentacion", "texto", "Tipo de alimentacion descrito"),
    _item("frutas_verduras", "sn", "Consumo de frutas y verduras investigado"),
    _item("ultraprocesados", "sn", "Consumo de alimentos ultraprocesados investigado"),
    _item("agua_diaria", "sn", "Consumo de agua diario investigado"),
    _item("dieta_especial", "texto", "Dieta especial o restriccion alimentaria"),
]

# ============================================================================
# SECCION F3 — HIGIENE (nivel paciente)
# ============================================================================
SECC_F3 = [
    _item("bano_diario", "sn", "Bano diario"),
    _item("cepillado", "snv", "Cepillado dental (frecuencia registrada)"),
    _item("cambio_ropa", "sn", "Cambio de ropa diario"),
    _item("lavado_manos", "sn", "Lavado de manos investigado"),
]

# ============================================================================
# SECCION F4 — TOXICOMANIAS Y HABITOS (nivel paciente; habito completo)
# ============================================================================
SECC_F4 = [
    _item(i, "snc", l + "  [si/no | frecuencia | cantidad | anios de consumo]")
    for i, l in [
        ("tabaquismo", "Tabaquismo"),
        ("alcoholismo", "Alcoholismo"),
        ("marihuana", "Marihuana"),
        ("cocaina", "Cocaina"),
        ("anfetaminas", "Anfetaminas"),
        ("opiacos", "Opiaceos"),
        ("benzodiacepinas", "Benzodiacepinas sin receta"),
        ("cafeina", "Cafeina excesiva"),
    ]
]

# ============================================================================
# SECCION F5 — ACTIVIDAD FISICA (nivel paciente)
# ============================================================================
SECC_F5 = [
    _item("ejercicio", "sn", "Realiza ejercicio regularmente"),
    _item("tipo_ejercicio", "texto", "Tipo de ejercicio registrado"),
    _item("frecuencia_semanal", "texto", "Frecuencia semanal registrada"),
    _item("duracion_sesion", "texto", "Duracion por sesion registrada"),
    _item("sedentarismo", "sn", "Sedentarismo identificado"),
]

# ============================================================================
# SECCION F6 — OCUPACION Y EXPOSICION LABORAL (nivel paciente)
# ============================================================================
SECC_F6 = [
    _item("ocupacion_actual", "snv", "Ocupacion actual registrada"),
    _item("toxicos", "sn", "Exposicion a toxicos en el trabajo"),
    _item("carcinogenos", "sn", "Exposicion a carcinogenos"),
    _item("ruido", "sn", "Exposicion a ruido excesivo"),
    _item("radiacion", "sn", "Exposicion a radiacion"),
    _item("epp", "sn", "Uso de equipo de proteccion personal"),
]

# ============================================================================
# SECCION F7 — ANTECEDENTES GINECO-OBSTETRICOS (paciente femenina; nivel paciente)
# ============================================================================
SECC_F7 = [
    _item("menarca", "snv", "Menarca (edad de primera regla)"),
    _item("ciclos_regulares", "snv", "Ciclos regulares"),
    _item("duracion_ciclo", "snv", "Duracion del ciclo"),
    _item("duracion_menstruacion", "snv", "Duracion de la menstruacion"),
    _item("fum", "fecha", "Fecha de ultima menstruacion (FUM)"),
    _item("dismenorrea", "snv", "Dismenorrea (dolor menstrual)"),
    _item("menorragia", "snv", "Menorragia (sangrado abundante)"),
    _item("metrorragia", "snv", "Metrorragia (sangrado entre ciclos)"),
    _item("amenorrea", "snv", "Amenorrea (ausencia de regla)"),
    _item("menopausia", "snv", "Menopausia"),
    _item("edad_menopausia", "snv", "Edad de menopausia"),
    _item("sintomas_climatericos", "snv", "Sintomas climatericos"),
    _item("vida_sexual", "sn", "Vida sexual activa"),
    _item("num_parejas", "snv", "Numero de parejas sexuales"),
    _item("metodo_anticonceptivo", "snv", "Uso de metodo anticonceptivo"),
    _item("tipo_metodo", "texto", "Tipo de metodo anticonceptivo"),
    _item("gestaciones", "numero", "Gestaciones (numero)"),
    _item("partos", "numero", "Partos (numero)"),
    _item("cesareas", "numero", "Cesareas (numero)"),
    _item("abortos", "numero", "Abortos (numero)"),
    _item("hijos_vivos", "numero", "Hijos vivos (numero)"),
    _item("hijos_fallecidos", "snv", "Hijos fallecidos (numero y causa)"),
    _item("complicaciones_embarazo", "snv", "Complicaciones en embarazos previos"),
    _item("papanicolaou", "sn", "Papanicolaou realizado"),
    _item("fecha_pap", "fecha", "Fecha del ultimo Papanicolaou"),
    _item("resultado_pap", "texto", "Resultado del ultimo Papanicolaou"),
    _item("mamografia", "sn", "Mamografia realizada"),
    _item("fecha_mamo", "fecha", "Fecha de ultima mamografia"),
    _item("resultado_mamo", "texto", "Resultado de ultima mamografia"),
    _item("autoexploracion", "sn", "Autoexploracion mamaria"),
    _item("ets", "snv", "Antecedente de enfermedad de transmision sexual"),
    _item("violencia_sexual", "snv", "Antecedente de violencia sexual"),
    _item("violencia_domestica", "snv", "Antecedente de violencia domestica"),
]

# ============================================================================
# SECCIONES G — INTERROGATORIO POR APARATOS Y SISTEMAS (nivel nota)
# ============================================================================
SECC_G1 = _sn(
    ("fiebre", "Fiebre"),
    ("escalofrios", "Escalofrios"),
    ("malestar", "Malestar general"),
    ("fatiga", "Fatiga"),
    ("perdida_peso", "Perdida de peso"),
    ("aumento_peso", "Aumento de peso"),
    ("sudoracion", "Sudoracion excesiva"),
    ("sed", "Sed excesiva"),
    ("apetito_aumentado", "Apetito aumentado"),
    ("apetito_disminuido", "Apetito disminuido"),
)

SECC_G2 = _sn(
    ("palpitaciones", "Palpitaciones"),
    ("dolor_precordial", "Dolor precordial"),
    ("opresion_toracica", "Opresion toracica"),
    ("disnea_esfuerzo", "Disnea de esfuerzo"),
    ("disnea_reposo", "Disnea en reposo"),
    ("ortopnea", "Ortopnea (dificultad para respirar acostada)"),
    ("dpn", "Disnea paroxistica nocturna"),
    ("edema_piernas", "Edema en piernas o tobillos"),
    ("cianosis", "Cianosis"),
    ("sincope", "Sincope (desmayo)"),
    ("mareo", "Mareo"),
    ("claudicacion", "Claudicacion intermitente"),
    ("varices", "Varices"),
    ("palidez", "Palidez"),
)

SECC_G3 = _sn(
    ("tos", "Tos"),
    ("tos_seca", "Tos seca"),
    ("tos_productiva", "Tos productiva"),
    ("hemoptisis", "Hemoptisis (sangre en flema)"),
    ("expectoracion", "Expectoracion color y cantidad"),
    ("disnea", "Disnea"),
    ("sibilancias", "Sibilancias"),
    ("estridor", "Estridor"),
    ("dolor_toracico_resp", "Dolor toracico al respirar"),
    ("ronquido", "Ronquido"),
    ("apnea_sueno", "Apnea del sueno"),
    ("congestion_nasal", "Congestion nasal"),
    ("rinorrea", "Rinorrea"),
    ("epistaxis", "Epistaxis"),
    ("dolor_faringeo", "Dolor faringeo"),
    ("disfonia", "Disfonia o ronquera"),
)

SECC_G4 = _sn(
    ("disfagia", "Disfagia (dificultad para tragar)"),
    ("odinofagia", "Odinofagia (dolor al tragar)"),
    ("pirosis", "Pirosis (acidez)"),
    ("regurgitacion", "Regurgitacion"),
    ("nausea", "Nausea"),
    ("vomito", "Vomito"),
    ("hematemesis", "Hematemesis (vomito con sangre)"),
    ("dolor_abdominal", "Dolor abdominal"),
    ("distension", "Distension abdominal"),
    ("meteorismo", "Meteorismo"),
    ("diarrea", "Diarrea"),
    ("estrenimiento", "Estrenimiento"),
    ("melena", "Melena (heces negras)"),
    ("rectorragia", "Rectorragia (sangre en heces)"),
    ("tenesmo", "Tenesmo"),
    ("ictericia", "Ictericia"),
    ("coluria", "Coluria (orina oscura)"),
    ("acolia", "Acolia (heces claras)"),
    ("hemorroides", "Hemorroides"),
    ("flatulencia", "Flatulencia excesiva"),
)

SECC_G5 = _sn(
    ("disuria", "Disuria (dolor al orinar)"),
    ("poliuria", "Poliuria"),
    ("oliguria", "Oliguria"),
    ("anuria", "Anuria"),
    ("nicturia", "Nicturia"),
    ("urgencia", "Urgencia urinaria"),
    ("incontinencia", "Incontinencia urinaria"),
    ("retencion", "Retencion urinaria"),
    ("hematuria", "Hematuria"),
    ("orina_turbia", "Orina turbia o fetida"),
    ("colico_renal", "Colico renal"),
    ("dolor_lumbar", "Dolor lumbar"),
    ("secrecion_uretral", "Secrecion uretral"),
)

SECC_G6 = _sn(
    ("leucorrea", "Leucorrea (flujo vaginal)"),
    ("prurito_vulvar", "Prurito vulvar"),
    ("dolor_pelvico", "Dolor pelvico"),
    ("dispareunia", "Dispareunia (dolor en relaciones)"),
    ("sangrado_poscoital", "Sangrado poscoital"),
    ("masas_mamarias", "Masas mamarias"),
    ("dolor_mamario", "Dolor mamario"),
    ("secrecion_pezon", "Secrecion por pezon"),
    ("cambios_pezon", "Cambios en la piel del pezon"),
    ("prolapso", "Prolapso uterino"),
    ("bochornos", "Bochornos"),
    ("sequedad_vaginal", "Sequedad vaginal"),
)

SECC_G7 = _sn(
    ("dolor_articular", "Dolor articular"),
    ("rigidez_matutina", "Rigidez matutina"),
    ("hinchazon_articular", "Hinchazon articular"),
    ("deformidad", "Deformidad articular"),
    ("dolor_muscular", "Dolor muscular"),
    ("debilidad_muscular", "Debilidad muscular"),
    ("calambres", "Calambres"),
    ("dolor_espalda", "Dolor de espalda"),
    ("dolor_cervical", "Dolor cervical"),
    ("limitacion_movimientos", "Limitacion de movimientos"),
    ("fracturas_fragilidad", "Fracturas por fragilidad"),
)

SECC_G8 = _sn(
    ("cefalea", "Cefalea"),
    ("migrana", "Migrana"),
    ("mareo_vertigo", "Mareo o vertigo"),
    ("sincope", "Sincope"),
    ("convulsiones", "Convulsiones"),
    ("parestesias", "Parestesias (hormigueo)"),
    ("anestesia", "Anestesia (perdida de sensibilidad)"),
    ("paresia_plejia", "Paresia o plejia"),
    ("temblor", "Temblor"),
    ("ataxia", "Ataxia"),
    ("disartria", "Disartria"),
    ("afasia", "Afasia"),
    ("amnesia", "Amnesia"),
    ("alteracion_conciencia", "Alteracion de la conciencia"),
    ("alteracion_sueno", "Alteracion del sueno"),
    ("alteracion_marcha", "Alteracion de la marcha"),
    ("perdida_equilibrio", "Perdida de equilibrio"),
    ("neuralgia", "Neuralgia"),
)

SECC_G9 = _sn(
    ("intolerancia_calor", "Intolerancia al calor"),
    ("intolerancia_frio", "Intolerancia al frio"),
    ("sudoracion_excesiva", "Sudoracion excesiva"),
    ("palpitaciones", "Palpitaciones"),
    ("temblor_fino", "Temblor fino"),
    ("cambio_peso", "Cambio de peso sin causa"),
    ("polidipsia", "Polidipsia"),
    ("poliuria", "Poliuria"),
    ("polifagia", "Polifagia"),
    ("alteracion_piel", "Alteracion en la piel"),
    ("caida_cabello", "Caida de cabello"),
    ("hirsutismo", "Hirsutismo"),
    ("galactorrea", "Galactorrea"),
    ("amenorrea", "Amenorrea"),
    ("bochornos", "Bochornos"),
    ("bocio", "Bocio"),
    ("exoftalmos", "Exoftalmos"),
)

SECC_G10 = _sn(
    ("palidez", "Palidez"),
    ("astenia", "Astenia"),
    ("adinamia", "Adinamia"),
    ("disnea", "Disnea"),
    ("palpitaciones", "Palpitaciones"),
    ("equimosis", "Equimosis (moretones)"),
    ("petequias", "Petequias"),
    ("sangrado_encias", "Sangrado de encias"),
    ("epistaxis", "Epistaxis"),
    ("hematomas", "Hematomas"),
    ("adenopatias", "Adenopatias"),
    ("esplenomegalia", "Esplenomegalia"),
    ("hepatomegalia", "Hepatomegalia"),
    ("fiebre", "Fiebre"),
    ("sudoracion_nocturna", "Sudoracion nocturna"),
    ("perdida_peso", "Perdida de peso"),
)

SECC_G11 = _sn(
    ("infecciones_recurrentes", "Infecciones recurrentes"),
    ("fiebre_recurrente", "Fiebre recurrente"),
    ("adenopatias", "Adenopatias"),
    ("artralgias", "Artralgias"),
    ("eritema", "Eritema"),
    ("ulceras_orales", "Ulceras orales"),
    ("fotofobia", "Fotofobia"),
    ("raynaud", "Fenomeno de Raynaud"),
    ("sequedad_mucosas", "Sequedad de mucosas"),
    ("alergias", "Alergias"),
)

SECC_G12 = _sn(
    ("prurito", "Prurito"),
    ("eritema", "Eritema"),
    ("urticaria", "Urticaria"),
    ("rash", "Rash"),
    ("lesiones_cutaneas", "Lesiones cutaneas"),
    ("ulceras", "Ulceras"),
    ("cambios_coloracion", "Cambios de coloracion"),
    ("nevos", "Nevos nuevos o cambiados"),
    ("alopecia", "Alopecia"),
    ("onicomicosis", "Onicomicosis"),
    ("sequedad_piel", "Sequedad de piel"),
    ("sudoracion_excesiva", "Sudoracion excesiva"),
)

SECC_G13 = _sn(
    ("agudeza_visual", "Disminucion de agudeza visual"),
    ("vision_borrosa", "Vision borrosa"),
    ("diplopia", "Diplopia"),
    ("fotofobia", "Fotofobia"),
    ("lagrimeo", "Lagrimeo"),
    ("ojo_seco", "Ojo seco"),
    ("dolor_ocular", "Dolor ocular"),
    ("hipoacusia", "Hipoacusia"),
    ("acufenos", "Acufenos"),
    ("vertigo", "Vertigo"),
    ("otorrea", "Otorrea"),
    ("otalgia", "Otalgia"),
    ("anosmia", "Anosmia"),
    ("hiposmia", "Hiposmia"),
    ("ageusia", "Ageusia"),
    ("disgeusia", "Disgeusia"),
)

SECC_G14 = _sn(
    ("tristeza", "Tristeza persistente"),
    ("anhedonia", "Anhedonia"),
    ("ansiedad", "Ansiedad"),
    ("crisis_angustia", "Crisis de angustia"),
    ("irritabilidad", "Irritabilidad"),
    ("alteracion_sueno", "Alteracion del sueno"),
    ("alteracion_apetito", "Alteracion del apetito"),
    ("dificultad_concentracion", "Dificultad de concentracion"),
    ("ideas_muerte", "Ideas de muerte"),
    ("ideacion_suicida", "Ideacion suicida"),
    ("intento_suicida", "Antecedente de intento suicida"),
    ("psicofarmacos", "Consumo de psicofarmacos"),
    ("hospitalizacion_psiquiatrica", "Antecedente de hospitalizacion psiquiatrica"),
)

# ============================================================================
# SECCION H1 — SIGNOS VITALES Y SOMATOMETRIA (nivel nota; auto desde controles)
# ============================================================================
SECC_H1 = [
    _item("temperatura", "auto", "Temperatura registrada (C)"),
    _item("ta", "auto", "Tension arterial registrada (mmHg)"),
    _item("fc", "auto", "Frecuencia cardiaca registrada (lpm)"),
    _item("fr", "auto", "Frecuencia respiratoria registrada (rpm)"),
    _item("spo2", "auto", "Saturacion de oxigeno registrada (%)"),
    _item("peso", "auto", "Peso registrado (kg)"),
    _item("talla", "auto", "Talla registrada (cm)"),
    _item("imc", "auto", "IMC calculado"),
    _item("pc", "snv", "Perimetro cefalico (pediatria)"),
    _item("perimetro_abdominal", "snv", "Perimetro abdominal"),
    _item("cintura", "snv", "Perimetro de cintura"),
    _item("cadera", "snv", "Perimetro de cadera"),
]

# ============================================================================
# SECCION H2 — HABITUS EXTERIOR (nivel nota)
# ============================================================================
SECC_H2 = [
    _item("aspecto_general", "texto", "Aspecto general descrito"),
    _item("estado_conciencia", "enum", "Estado de conciencia",
          ["LUCIDO", "SOMNOLIENTO", "CONFUSO", "COMA"]),
    _item("estado_nutricion", "enum", "Estado de nutricion",
          ["NORMOPESO", "BAJO PESO", "SOBREPESO", "OBESIDAD"]),
    _item("estado_hidratacion", "enum", "Estado de hidratacion",
          ["NORMAL", "LEVE", "MODERADA", "GRAVE"]),
    _item("coloracion", "texto", "Coloracion de piel y mucosas"),
    _item("facies", "texto", "Facies descrita"),
    _item("postura_marcha", "texto", "Postura y marcha"),
    _item("edad_aparente", "sn", "Edad aparente concordante"),
]

# ============================================================================
# SECCION H3 — EXPLORACION POR REGIONES (nivel nota; estado + hallazgo)
# ============================================================================
SECC_H3 = [
    _item(i, "snh", l + "  [explorado | hallazgo]")
    for i, l in [
        ("cabeza", "Cabeza"),
        ("cuello", "Cuello"),
        ("torax_anterior", "Torax anterior"),
        ("torax_posterior", "Torax posterior"),
        ("abdomen", "Abdomen"),
        ("pelvis", "Pelvis"),
        ("extremidades_superiores", "Extremidades superiores"),
        ("extremidades_inferiores", "Extremidades inferiores"),
        ("columna", "Columna vertebral"),
        ("genitales", "Genitales"),
        ("inguinal", "Region inguinal"),
        ("linfatico", "Sistema linfatico"),
        ("piel_anexos", "Piel y anexos"),
        ("neurologico", "Sistema neurologico"),
        ("fondo_ojo", "Fondo de ojo (si aplica)"),
        ("otoscopia", "Otoscopia"),
        ("orofaringe", "Orofaringe"),
        ("odontograma", "Odontograma"),
        ("periodontograma", "Periodontograma"),
    ]
]

# ============================================================================
# SECCION I — DIAGNOSTICOS Y PROBLEMAS CLINICOS (nivel nota)
# ============================================================================
SECC_I = [
    _item("dx_principal", "texto", "Diagnostico principal registrado"),
    _item("dx_secundario", "texto", "Diagnostico secundario registrado"),
    _item("dx_crecimiento", "texto", "Diagnostico de crecimiento y desarrollo (pediatria)"),
    _item("dx_nutricional", "texto", "Diagnostico nutricional"),
    _item("impresion_diagnostica", "texto", "Impresion diagnostica registrada"),
    _item("problemas", "texto", "Problemas clinicos identificados"),
    _item("cie10", "texto", "CIE-10 asignado"),
]

# ============================================================================
# SECCION J — PLAN DE ESTUDIOS Y TRATAMIENTO (nivel nota)
# ============================================================================
SECC_J = [
    _item("lab", "texto", "Estudios de laboratorio solicitados"),
    _item("gabinete", "texto", "Estudios de gabinete solicitados"),
    _item("terapeutica", "texto", "Terapeutica farmacologica indicada"),
    _item("dosis_via", "texto", "Dosis, via y frecuencia registradas"),
    _item("duracion", "texto", "Duracion del tratamiento registrada"),
    _item("alimentacion", "texto", "Plan de alimentacion indicado"),
    _item("cuidados", "texto", "Cuidados generales indicados"),
    _item("interconsulta", "snv", "Interconsulta solicitada"),
    _item("pronostico", "texto", "Pronostico registrado"),
    _item("proxima_cita", "fecha", "Fecha de proxima cita registrada"),
    _item("signos_alarma", "sn", "Signos de alarma explicados a la paciente"),
    _item("referencia", "snv", "Referencia a otro nivel de atencion"),
]

# ============================================================================
# SECCION K — TRAZABILIDAD Y RESPONSABILIDAD (nivel nota; pie automatico)
# ============================================================================
SECC_K = [
    _item("medico", "auto", "Nombre del medico registrado"),
    _item("firma", "auto", "Firma autografica del medico (linea impresa)"),
    _item("cedula", "texto", "Cedula profesional registrada"),
    _item("matricula", "texto", "Matricula institucional registrada"),
    _item("elaboracion", "auto", "Fecha y hora de elaboracion de la nota"),
    _item("folio_receta", "texto", "Folio de receta"),
    _item("consentimiento", "auto", "Carta de consentimiento informado firmada (derivada: existe CONSENTIMIENTO_*.txt en el episodio)"),
    _item("testigo1", "texto", "Testigo 1 (si aplica)"),
    _item("testigo2", "texto", "Testigo 2 (si aplica)"),
    _item("sha256", "auto", "SHA-256 del contenido de la nota"),
]

# ============================================================================
# SECCION L — NOTAS DE EVOLUCION (nota tipo=evolucion; una por visita de seguimiento)
# ============================================================================
SECC_L = [
    _item("fecha_hora", "auto", "Fecha y hora de la evolucion"),
    _item("signos_vitales", "sn", "Signos vitales actualizados"),
    _item("estado_actual", "texto", "Estado actual del paciente descrito"),
    _item("cambios", "texto", "Cambios en el cuadro clinico"),
    _item("resultados", "texto", "Resultados de estudios recibidos"),
    _item("ajuste_tratamiento", "texto", "Ajuste de tratamiento"),
    _item("pronostico", "texto", "Pronostico actualizado"),
    _item("firma", "auto", "Nombre y firma del medico"),
]

# ============================================================================
# SECCION M — NOTA DE EGRESO (nota tipo=egreso; una por episodio maximo)
# ============================================================================
SECC_M = [
    _item("fecha_ingreso", "fecha", "Fecha de ingreso"),
    _item("fecha_egreso", "fecha", "Fecha de egreso"),
    _item("dx_ingreso", "texto", "Diagnostico de ingreso"),
    _item("dx_egreso", "texto", "Diagnostico de egreso"),
    _item("resumen_evolucion", "texto", "Resumen de la evolucion"),
    _item("estudios", "texto", "Estudios realizados"),
    _item("tratamiento", "texto", "Tratamiento administrado"),
    _item("condiciones_egreso", "texto", "Condiciones de egreso"),
    _item("plan_ambulatorio", "texto", "Plan de manejo ambulatorio"),
    _item("recomendaciones", "texto", "Recomendaciones e indicaciones"),
    _item("signos_alarma", "sn", "Signos de alarma explicados"),
    _item("firma", "auto", "Nombre y firma del medico"),
]

# ============================================================================
# ENSAMBLADO — lista maestra (transcribir completa a app.py en T12)
# ============================================================================
def _sec(id_, titulo, nivel, items, aplica=None):
    s = {"id": id_, "titulo": titulo, "nivel": nivel, "items": items}
    if aplica:
        s["aplica"] = aplica
    return s


SECCIONES_HISTORIA = [
    _sec("A", "Datos de identificacion y filiacion", "paciente", SECC_A),
    _sec("B", "Motivo de la atencion", "nota", SECC_B),
    _sec("C1", "Enfermedad actual: inicio y evolucion", "nota", SECC_C1),
    _sec("C2", "Enfermedad actual: caracterizacion del dolor", "nota", SECC_C2,
         aplica="si_dolor"),
    _sec("C3", "Enfermedad actual: sintomas generales", "nota", SECC_C3),
    _sec("D", "Antecedentes heredofamiliares", "paciente", SECC_D),
    _sec("E1", "Antecedentes personales: enfermedades previas", "paciente", SECC_E1),
    _sec("E2", "Antecedentes personales: quirurgicos", "paciente", SECC_E2),
    _sec("E3", "Antecedentes personales: traumaticos", "paciente", SECC_E3),
    _sec("E4", "Antecedentes personales: alergias", "paciente", SECC_E4),
    _sec("E5", "Antecedentes personales: inmunizaciones", "paciente", SECC_E5),
    _sec("F1", "Antecedentes no patologicos: vivienda", "paciente", SECC_F1),
    _sec("F2", "Antecedentes no patologicos: habitos alimenticios", "paciente", SECC_F2),
    _sec("F3", "Antecedentes no patologicos: higiene", "paciente", SECC_F3),
    _sec("F4", "Antecedentes no patologicos: toxicomanias y habitos", "paciente", SECC_F4),
    _sec("F5", "Antecedentes no patologicos: actividad fisica", "paciente", SECC_F5),
    _sec("F6", "Antecedentes no patologicos: ocupacion y exposicion", "paciente", SECC_F6),
    _sec("F7", "Antecedentes gineco-obstetricos", "paciente", SECC_F7),
    _sec("G1", "Interrogatorio: generales", "nota", SECC_G1),
    _sec("G2", "Interrogatorio: cardiovascular", "nota", SECC_G2),
    _sec("G3", "Interrogatorio: respiratorio", "nota", SECC_G3),
    _sec("G4", "Interrogatorio: digestivo", "nota", SECC_G4),
    _sec("G5", "Interrogatorio: urinario", "nota", SECC_G5),
    _sec("G6", "Interrogatorio: reproductor femenino", "nota", SECC_G6),
    _sec("G7", "Interrogatorio: musculoesqueletico", "nota", SECC_G7),
    _sec("G8", "Interrogatorio: neurologico", "nota", SECC_G8),
    _sec("G9", "Interrogatorio: endocrino", "nota", SECC_G9),
    _sec("G10", "Interrogatorio: hematologico", "nota", SECC_G10),
    _sec("G11", "Interrogatorio: inmunologico", "nota", SECC_G11),
    _sec("G12", "Interrogatorio: piel y anexos", "nota", SECC_G12),
    _sec("G13", "Interrogatorio: organos de los sentidos", "nota", SECC_G13),
    _sec("G14", "Interrogatorio: salud mental y emocional", "nota", SECC_G14),
    _sec("H1", "Exploracion fisica: signos vitales y somatometria", "nota", SECC_H1),
    _sec("H2", "Exploracion fisica: habitus exterior", "nota", SECC_H2),
    _sec("H3", "Exploracion fisica: por regiones", "nota", SECC_H3),
    _sec("I", "Diagnosticos y problemas clinicos", "nota", SECC_I),
    _sec("J", "Plan de estudios y tratamiento", "nota", SECC_J),
    _sec("K", "Trazabilidad y responsabilidad legal", "nota", SECC_K),
    _sec("L", "Notas de evolucion (por visita de seguimiento)", "evolucion", SECC_L),
    _sec("M", "Nota de egreso", "egreso", SECC_M),
]

# ---- Validacion rapida: python3 esquemas/historia_nom004_esquema.py ----
if __name__ == "__main__":
    total_items = 0
    ids_seccion = set()
    for s in SECCIONES_HISTORIA:
        assert s["id"] not in ids_seccion, f"Seccion duplicada: {s['id']}"
        ids_seccion.add(s["id"])
        ids_item = set()
        for it in s["items"]:
            assert it["id"] not in ids_item, f"Item duplicado en {s['id']}: {it['id']}"
            ids_item.add(it["id"])
            assert it["t"] in ("auto", "texto", "numero", "fecha", "enum",
                               "sn", "snns", "snna", "snv", "snc", "snh", "libre"), \
                f"Tipo invalido: {it['t']}"
            total_items += 1
    print(f"OK: {len(SECCIONES_HISTORIA)} secciones, {total_items} items unicos.")
    niveles = {}
    for s in SECCIONES_HISTORIA:
        niveles[s["nivel"]] = niveles.get(s["nivel"], 0) + 1
    print("Por nivel:", niveles)
