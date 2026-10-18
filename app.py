# -*- coding: utf-8 -*-
"""
SISTEMA DE EXPEDIENTES MATERNIDAD - v1.0 (PROTOTIPO UTILIZABLE)
Monolito de un solo archivo. Solo stdlib de Python 3.8+. Offline total.
Los datos viven en carpetas abiertas junto a este archivo (del médico).

INDICE DEL ARCHIVO:
  00 CONSTANTES          Titulo, rutas, paleta
  01 UTILIDADES          Normalizacion, fechas, CURP, hashing, apertura
  02 MODELO DE DATOS     Estructura de carpetas, datos.json, episodios
  03 SERVICIO PACIENTE   Crear/cargar/buscar pacientes, homonimos, episodios
  04 SERVICIO BUSQUEDA   Busqueda normalizada, filtros, recientes
  05 SERVICIO DOCUMENTAL Plantillas, marcadores, docx/xlsx, progreso quimico
  06 SERVICIO AUDITORIA  Bitacora hash-chain, historial, log de errores
  07 SERVICIO RESPALDO   Respaldar, verificar, restaurar
  08 SERVICIO CUMPLIMIENTO Consentimiento, aviso privacidad, historia NOM-004
  09 SERVICIO EXPORTACION XLSX/CSV/PDF minimal, zip de expediente
  10 INTERFAZ            Ventana principal (tkinter)
  11 PUNTO DE ENTRADA    main()

REGLAS: datos en formato abierto (carpeta + json + docx/xlsx);
cero dependencias de red; instalacion = tener Python y ejecutar.
"""
import os
import sys
import re
import json
import shutil
import zipfile
import hashlib
import unicodedata
import calendar
import datetime
import traceback
import subprocess
import threading
import queue
import time
import base64
import csv
import io
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    TK_DISPONIBLE = True
except Exception:
    TK_DISPONIBLE = False
    tk = None

# P1 (ADR-0015/T9b): cryptography es OPCIONAL; sin ella el respaldo conserva
# integridad SHA-256 y la UI avisa. El core sigue stdlib-only (G10).
try:
    from cryptography.fernet import Fernet
    CRYPTO_DISPONIBLE = True
except Exception:
    Fernet = None
    CRYPTO_DISPONIBLE = False


# ===== 00 CONSTANTES =====

APP_TITLE = "SISTEMA DE EXPEDIENTES - MATERNIDAD v1.0"
VERSION = "1.0.0-prototipo"

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent

EXPEDIENTES_DIR = BASE_DIR / "EXPEDIENTES"
PLANTILLAS_DIR = BASE_DIR / "_PLANTILLAS"
RESPALDOS_DIR = BASE_DIR / "_RESPALDOS"
EXPORTACIONES_DIR = BASE_DIR / "EXPORTACIONES"
GENERAL_DIR = BASE_DIR / "DOCUMENTOS_SIN_PACIENTE"
BITACORA_PATH = BASE_DIR / "_BITACORA.log"
CONFIG_PATH = BASE_DIR / "_CONFIG.json"
ERROR_LOG_PATH = BASE_DIR / "_ERRORES.log"
LEEME_PATH = BASE_DIR / "LEEME.txt"

LETRAS = [chr(c) for c in range(ord("A"), ord("Z") + 1)]

CREMA = "#F5F0E1"
CREMA_CLARO = "#FAF7F0"
AZUL = "#0A2F5A"
AZUL_HOVER = "#1A3F6A"
TEXTO_OSCURO = "#1E1E1E"
TEXTO_CLARO = "#F5F0E1"
VERDE_OK = "#1B7A3D"
ROJO_AVISO = "#B00020"
BLANCO = "#FFFFFF"

# Design tokens (ADR-0022): espaciado, tipografia, botones y grises jerarquizados
ESP_XS, ESP_S, ESP_M, ESP_L = 4, 8, 12, 16
FUENTE_XS, FUENTE_S, FUENTE_M, FUENTE_L, FUENTE_XL, FUENTE_XXL = 9, 10, 11, 13, 15, 18
BTN_ALTO_MIN = 32   # px; filas de botones visibles (auditoria M10)
BTN_PADY = 6
GRIS_BORDE = "#CCCCCC"
GRIS_REJILLA = "#EEEEEE"
GRIS_TEXTO = "#555555"
GRIS_TERCIARIO = "#777777"
GRIS_APAGADO = "#888888"
GRIS_SECUNDARIO = "#666666"
AZUL_CHIP = "#DCE9F8"

# Navegacion y encabezados (ADR-0026): identidad por pantalla
NAV_BG = "#071A33"            # fondo profundo del sidebar
NAV_TARJETA = "#0E2A4D"       # tarjeta de contexto y hover
NAV_TEXTO = "#CFE0F5"         # texto de items inactivos
NAV_LEYENDA = "#7FA9DA"       # secciones (PACIENTES/ATENCION/...)
NAV_PIE = "#5F83AD"           # pie del sidebar
ACENTO = "#F2C94C"            # barra de la pagina activa (dorado)
ROJO_SALIR = "#E8B4B4"        # Salir en el sidebar

# ===== 00c ADAPTADOR DE PLATAFORMA (ADR-0051) =====
# TODA decision que depende del sistema operativo vive en esta seccion: el
# resto del monolito solo usa estas constantes y la funcion comando_abrir.
# Para soportar un SO nuevo basta agregar una rama aqui (y, si no hay Tk,
# sustituir la capa de interfaz — el nucleo secciones 01-09 no toca SO).
def _detectar_plataforma():
    p = sys.platform
    if p == "win32":
        return "windows"
    if p == "darwin":
        return "macos"
    if p.startswith("linux"):
        return "linux"
    return p or "desconocido"

PLATAFORMA = _detectar_plataforma()
ES_WINDOWS = PLATAFORMA == "windows"
ES_MACOS = PLATAFORMA == "macos"
ES_LINUX = PLATAFORMA == "linux"
# Movil (Android/iOS): Tk no existe alla. El dia que se pida, la capa Tk se
# sustituye (web local o Kivy/BeeWare) y el adaptador gana esas ramas; el
# nucleo (modelos, validacion, core clinico) no cambia (ADR-0051).
ES_MOVIL = False

# Familia tipografica por SO: Tk sustituye si falta, pero aqui se decide una
# sola vez (antes eran 271 literales "Segoe UI" repartidos por el archivo).
FUENTE_SISTEMA = ("Segoe UI" if ES_WINDOWS
                  else "Helvetica" if ES_MACOS
                  else "DejaVu Sans")


def comando_abrir(ruta):
    """(programa, args) para abrir un archivo o carpeta con la app del sistema.

    Pura y testeable (B101): es UNICO punto de decision para abrir cosas.
    En Windows el programa es el sentinel "startfile" (os.startfile solo
    existe alli); en macOS "open"; en el resto "xdg-open" con fallbacks.
    """
    ruta = str(ruta)
    if ES_WINDOWS:
        return "startfile", [ruta]
    if ES_MACOS:
        return "open", [ruta]
    return "xdg-open", [ruta]

MESES_CORTOS = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
                "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
MESES_TXT = {
    "ENE": 1, "ENERO": 1, "FEB": 2, "FEBRERO": 2, "MAR": 3, "MARZO": 3,
    "ABR": 4, "ABRIL": 4, "MAY": 5, "MAYO": 5, "JUN": 6, "JUNIO": 6,
    "JUL": 7, "JULIO": 7, "AGO": 8, "AGOSTO": 8, "SEP": 9, "SEPTIEMBRE": 9,
    "OCT": 10, "OCTUBRE": 10, "NOV": 11, "NOVIEMBRE": 11, "DIC": 12, "DICIEMBRE": 12,
}
MESES_LARGOS = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]



# ===== 00b ESQUEMA NOM-004 A-M (T16/ADR-0030) =====
# TRANScripcion fiel de esquemas/historia_nom004_esquema.py (canonico del dueno:
# 40 secciones, 477 items). El monolito no importa esquemas/ (ADR-0002); si el
# canonico cambia, se re-transcribe y el test A de conteo lo detecta.

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



# ===== 00c CATALOGOS CLINICOS (T18/ADR-0031) =====
# TRANScripcion de esquemas/catalogos_clinicos.py (canonico del medico:
# rangos de signos vitales, panic values de lab, preeclampsia, escalas
# MEWS/NEWS/qSOFA y registro de 26 calculadoras). NO se importa esquemas/
# (ADR-0002); si el canonico cambia, se re-transcribe y lo verifica el test A.
# CIFRAS DEL MEDICO: el software no inventa umbrales (P2 clinico).

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
# consultorio define como de mayor valor para su flujo rapido.
GRUPO_PRIORITARIAS = "★ Prioritarias (flujo rapido)"
# Defecto = el modulo que el dueno senalo como prioridad (composicion corporal).
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


# --- ADR-0043: calculadoras prioritarias del workflow (C-19) ---
# El consultorio define en _CONFIG.json ("calc_prioritarias") las
# calculadoras de mayor valor para su flujo rapido. Defecto: el modulo
# "Generales/Composicion corporal". Solo afecta presentacion/orden.

def calculadoras_prioritarias():
    """Ids validos de prioritarias desde _CONFIG.json (pura sobre config).

    Sin clave -> defecto (modulo Generales/Composicion corporal); con clave
    -> exactamente lo definido por el consultorio (permite quitar el defecto).
    """
    if "calc_prioritarias" not in leer_config():
        return [c for c in CALC_PRIORITARIAS_DEFECTO if c in CALCULADORAS]
    ids = leer_config().get("calc_prioritarias")
    if not isinstance(ids, list):
        return [c for c in CALC_PRIORITARIAS_DEFECTO if c in CALCULADORAS]
    return [c for c in ids if c in CALCULADORAS]


def alternar_prioritaria(cid):
    """Agrega o quita cid de las prioritarias y persiste en _CONFIG.json.

    Devuelve True si quedo marcada. Id desconocido -> False sin persistir.
    """
    if cid not in CALCULADORAS:
        return False
    actuales = calculadoras_prioritarias()
    if cid in actuales:
        actuales = [c for c in actuales if c != cid]
    else:
        actuales = actuales + [cid]
    guardar_config({"calc_prioritarias": actuales})
    return cid in actuales


# ============================================================================
# Validacion rapida: python3 esquemas/catalogos_clinicos.py
# ============================================================================


# ===== 01 UTILIDADES =====

def normalizar(texto):
    s = (texto or "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.replace("ñ", "n").replace("Ñ", "N")
    s = re.sub(r"[^A-Za-z0-9 .]", " ", s)
    s = re.sub(r"[\s.]+", "_", s)
    return s.strip("_").upper()


def fecha_hoy():
    return datetime.date.today().isoformat()


def hora_ahora():
    return datetime.datetime.now().strftime("%H:%M:%S")


def timestamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def fecha_legible(iso):
    try:
        d = datetime.date.fromisoformat(iso)
        return f"{d.day:02d}-{MESES_CORTOS[d.month - 1]}-{d.year}"
    except Exception:
        return iso or ""


def parsear_fecha(txt):
    txt = (txt or "").strip()
    if not txt:
        return None
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", txt)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", txt)
        if m:
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        else:
            m = re.fullmatch(r"(\d{1,2})[- ]([A-Za-z]+)[- ](\d{4})", txt)
            if not m:
                m = re.fullmatch(r"(\d{1,2})\s+de\s+([A-Za-z]+)\s+de\s+(\d{4})", txt, re.I)
            if not m:
                return None
            d, mes_txt, y = int(m.group(1)), m.group(2).upper(), int(m.group(3))
            if mes_txt not in MESES_TXT:
                return None
            mo = MESES_TXT[mes_txt]
    try:
        return datetime.date(y, mo, d).isoformat()
    except ValueError:
        return None


def dias_en_mes(anio, mes):
    return calendar.monthrange(int(anio), int(mes))[1]


def calcular_fpp(fum_iso):
    """FPP por Naegele: +7 dias primero (evita trampas de fin de mes), -3 meses, +1 anio."""
    try:
        f = datetime.date.fromisoformat(fum_iso)
    except Exception:
        return None
    d = f + datetime.timedelta(days=7)
    mes, anio = d.month - 3, d.year
    if mes <= 0:
        mes += 12
        anio -= 1
    anio += 1
    dia = min(d.day, calendar.monthrange(anio, mes)[1])
    try:
        return datetime.date(anio, mes, dia).isoformat()
    except ValueError:
        return None


def calcular_eg(fum_iso, hoy_iso=None):
    """Edad gestacional como '28 s 3 d'. None si FUM invalida o futura."""
    try:
        f = datetime.date.fromisoformat(fum_iso)
        hoy = datetime.date.fromisoformat(hoy_iso) if hoy_iso else datetime.date.today()
    except Exception:
        return None
    delta = (hoy - f).days
    if delta < 0:
        return None
    semanas, dias = divmod(delta, 7)
    return f"{semanas} s {dias} d"


# --- ADR-0027 (v2.5): cuentas regresivas y estadisticas de serie ---

def dias_desde(fecha_iso, hoy_iso=None):
    """Hoy menos fecha (negativo si la fecha es futura); None si no calculable."""
    try:
        f = datetime.date.fromisoformat(fecha_iso)
        h = (datetime.date.fromisoformat(hoy_iso) if hoy_iso
             else datetime.date.today())
    except (TypeError, ValueError):
        return None
    return (h - f).days


def dias_para_fpp(fpp_iso, hoy_iso=None):
    """Dias restantes hasta la FPP (negativo = ya paso); None si no calculable."""
    try:
        f = datetime.date.fromisoformat(fpp_iso)
        h = (datetime.date.fromisoformat(hoy_iso) if hoy_iso
             else datetime.date.today())
    except (TypeError, ValueError):
        return None
    return (f - h).days


def trimestre_desde_eg(eg):
    """Trimestre por semanas de EG (convencion aritmetica: I <14 s, II 14-27 s, III >=28 s)."""
    if not eg:
        return ""
    try:
        sem = int(str(eg).split()[0])
    except (ValueError, IndexError, TypeError):
        return ""
    if sem < 14:
        return "I"
    if sem < 28:
        return "II"
    return "III"


def stats_serie(puntos):
    """(ultimo, fecha_etiqueta, dias_hace, delta, minimo, maximo, media)
    de [(etiqueta, valor)]; None si vacia.

    ADR-0045/v6.0 (B96): extendido con min/max/media; los 4 primeros
    elementos son identicos a los de siempre (retrocompatible).
    """
    if not puntos:
        return None
    ultimo = puntos[-1][1]
    delta = None
    if len(puntos) >= 2 and isinstance(puntos[-2][1], (int, float)):
        delta = ultimo - puntos[-2][1]
    dias_hace = None
    etiqueta = puntos[-1][0]
    if etiqueta is not None:
        iso = parsear_fecha(str(etiqueta))
        if iso:
            dias_hace = dias_desde(iso)
    nums = [v for _, v in puntos if isinstance(v, (int, float))]
    minimo = min(nums) if nums else None
    maximo = max(nums) if nums else None
    media = (sum(nums) / len(nums)) if nums else None
    return (ultimo, (str(etiqueta) if etiqueta is not None else ""),
            dias_hace, delta, minimo, maximo, media)


# --- ADR-0045/v6.0 (B95): proyeccion temporal de series (C-17) ---
# Constantes configurables por el medico (precedente U85); solo afectan
# presentacion, nunca interpretacion clinica.
UMBRAL_HUECO_DIAS = 45   # hueco maximo entre mediciones que la linea une
GRAF_MAX_SERIES = 12     # series dibujadas antes del pie "+N no mostradas"
GRAF_MAX_CARDS = 6       # stat-cards visibles


def segmentos_temporales(puntos, umbral_dias=None):
    """Corta la serie donde el hueco entre mediciones consecutivas supera
    umbral_dias (fidelidad de representacion, NO juicio clinico).

    ADR-0045/v6.0 (B95): puntos = [(etiqueta_x, valor), ...] con etiquetas
    parseables por parsear_fecha. Devuelve (es_temporal, segmentos):
      - es_temporal True: cada segmento es [(dia_abs, valor, etiqueta), ...]
        con dia_abs = dias desde el primer punto de TODA la serie (proyeccion
        proporcional al tiempo real); el hueco > umbral corta el segmento.
      - es_temporal False (alguna etiqueta no parsea): 1 solo segmento en
        modo indice [(i, valor, etiqueta), ...] (espaciado igual = v1.x).
    Pura; sin UI. umbral_dias=None usa la constante UMBRAL_HUECO_DIAS.
    """
    if umbral_dias is None:
        umbral_dias = UMBRAL_HUECO_DIAS
    if not puntos:
        return True, []
    fechas = [parsear_fecha(str(et)) if et is not None else None
              for et, _ in puntos]
    if any(f is None for f in fechas):
        return False, [[(i, v, et) for i, (et, v) in enumerate(puntos)]]
    objs = [datetime.date.fromisoformat(f) for f in fechas]
    base = min(objs)
    pares = sorted(zip(objs, [v for _, v in puntos],
                       [str(et) for et, _ in puntos]), key=lambda t: t[0])
    segmentos = []
    actual = []
    prev = None
    for f, v, et in pares:
        if prev is not None and (f - prev).days > umbral_dias:
            segmentos.append(actual)
            actual = []
        actual.append((float((f - base).days), v, et))
        prev = f
    if actual:
        segmentos.append(actual)
    return True, segmentos


def validar_curp(curp):
    c = (curp or "").strip().upper()
    if not c:
        return True, "CURP vacia (opcional)"
    if len(c) != 18:
        return False, "La CURP debe tener 18 caracteres."
    if not re.fullmatch(r"[A-Z]{4}\d{6}[HM][A-Z]{5}[0-9A-Z]\d", c):
        return False, "Formato invalido. Ejemplo: GODE561231MDFRRL09"
    return True, "CURP valida."


# --- Tier A (v2.0, ADR-0024): utilidades puras de historia estructurada ---

def edad_anios(fecha_nacimiento_iso, hoy_iso=None):
    """Anios cumplidos a partir de fecha de nacimiento ISO; None si no calculable."""
    try:
        nac = datetime.date.fromisoformat(fecha_nacimiento_iso)
        hoy = (datetime.date.fromisoformat(hoy_iso) if hoy_iso
               else datetime.date.today())
    except (TypeError, ValueError):
        return None
    if nac > hoy:
        return None
    return hoy.year - nac.year - ((hoy.month, hoy.day) < (nac.month, nac.day))


_GPA_RE = re.compile(r"^[Gg]\d{1,2}[Pp]\d{1,2}[Aa]\d{1,2}$")


def validar_gpa(txt):
    """Validacion minima del G-P-A (vacio = opcional). Formato: G3P2A1."""
    t = (txt or "").strip()
    if not t:
        return True, "G-P-A vacio (opcional)"
    if not _GPA_RE.fullmatch(t):
        return False, "Formato G-P-A invalido. Ejemplo: G3P2A1"
    return True, "G-P-A valido."


def sha256_texto(texto):
    return hashlib.sha256(texto.encode("utf-8", "replace")).hexdigest()


def sha256_archivo(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(65536), b""):
            h.update(bloque)
    return h.hexdigest()


def escapar_xml(t):
    t = str(t)
    for a, b in (("&", "&amp;"), ("<", "&lt;"), (">", "&gt;"), ('"', "&quot;"), ("'", "&apos;")):
        t = t.replace(a, b)
    return t


def abrir_archivo(ruta):
    """Abre archivo/carpeta con la app del sistema (ADR-0051: decision unica
    en comando_abrir; en Linux cae por fallbacks si falta xdg-open)."""
    prog, args = comando_abrir(ruta)
    if prog == "startfile":
        os.startfile(str(ruta))  # solo Windows; el adaptador lo decide
        return
    intentos = [(prog, args)]
    if not ES_MACOS:
        # Linux y demas: orden de preferencia si xdg-open no esta instalado
        intentos += [["gio", ["open", str(ruta)]], ["kde-open", [str(ruta)]]]
    ultimo = None
    for cmd in intentos:
        try:
            subprocess.Popen(cmd, stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL)
            return
        except Exception as e:      # FileNotFoundError u otros
            ultimo = e
            continue
    raise RuntimeError(
        f"No se pudo abrir {ruta}: falta un gestor de archivos "
        f"(instala xdg-open/gio en Linux). Detalle: {ultimo}")


def abrir_carpeta(ruta):
    abrir_archivo(ruta)


def log_exc(err):
    try:
        with open(ERROR_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{fecha_hoy()} {hora_ahora()}] {err}\n")
            f.write(traceback.format_exc() + "\n")
    except Exception:
        pass


def leer_config():
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


# ===== 01b ESCRITURA SEGURA (ADR-0033 / T21) =====

JOURNAL_PATH = BASE_DIR / "_JOURNAL.log"
DIAG_PATH = BASE_DIR / "_DIAGNOSTICO.log"
_BAK_MAX = 3


def escribir_atomico(path, texto):
    """Escribe con tmp + fsync + os.replace (atomico POSIX y Windows)."""
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(texto)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, path)


def _journal(operacion, archivo, sha, estado):
    try:
        with open(JOURNAL_PATH, "a", encoding="utf-8") as f:
            f.write(f"{fecha_hoy()} {hora_ahora()}|{operacion}|{archivo}|{sha}|{estado}\n")
    except OSError:
        pass


def _rotar_bak(path):
    """Rota .bak (hasta _BAK_MAX generaciones FIFO) del archivo bueno."""
    path = Path(path)
    if not path.exists():
        return
    bak3 = path.with_suffix(path.suffix + ".bak3")
    bak2 = path.with_suffix(path.suffix + ".bak2")
    bak1 = path.with_suffix(path.suffix + ".bak1")
    if bak2.exists():
        os.replace(bak2, bak3)
    if bak1.exists():
        os.replace(bak1, bak2)
    os.replace(path, bak1)


def escribir_json_atomico(path, obj, rotar=False):
    """Escritura segura de JSON: journal pendiente -> (rotar .bak) -> atomico -> OK."""
    path = Path(path)
    texto = json.dumps(obj, ensure_ascii=False, indent=2)
    sha = sha256_texto(texto)
    _journal("WRITE", str(path), sha, "PENDIENTE")
    if rotar:
        _rotar_bak(path)
    escribir_atomico(path, texto)
    _journal("WRITE", str(path), sha, "OK")


def recuperar_journal():
    """Recupera escrituras pendientes del journal (arranque). -> lista de hallazgos."""
    hallazgos = []
    if not JOURNAL_PATH.exists():
        return hallazgos
    try:
        lineas = JOURNAL_PATH.read_text(encoding="utf-8").splitlines()
    except OSError:
        return hallazgos
    pendientes = {}
    for l in lineas:
        partes = l.split("|")
        if len(partes) == 5 and partes[4].strip() == "PENDIENTE":
            pendientes[partes[2]] = partes[3]
    # una escritura con OK posterior ya no esta pendiente
    for l in lineas:
        partes = l.split("|")
        if len(partes) == 5 and partes[4].strip() == "OK":
            pendientes.pop(partes[2], None)
    for archivo, sha in pendientes.items():
        destino = Path(archivo)
        tmp = destino.with_suffix(destino.suffix + ".tmp")
        if tmp.exists():
            try:
                contenido = tmp.read_text(encoding="utf-8")
                if sha256_texto(contenido) == sha:
                    os.replace(tmp, destino)
                    hallazgos.append(f"Recuperado de journal: {destino.name}")
                    _journal("RECOVER", archivo, sha, "OK")
                    continue
            except OSError:
                pass
        hallazgos.append(f"Escritura pendiente sin tmp valido: {destino.name}")
    return hallazgos


def diagnostico_arranque():
    """Verifica parseo de JSON, bitacora y journal al arranque. -> lista de lineas."""
    lineas = [f"[{fecha_hoy()} {hora_ahora()}] Diagnostico de arranque"]
    # journal
    rec = recuperar_journal()
    for r in rec:
        lineas.append(f"  JOURNAL: {r}")
    # bitacora
    ok_b, msg_b = verificar_bitacora()
    lineas.append(f"  BITACORA: {'OK' if ok_b else 'ERROR - ' + msg_b}")
    # JSONs
    n_ok = n_bad = 0
    for carpeta in EXPEDIENTES_DIR.glob("*/*"):
        dj = carpeta / "datos.json"
        if dj.exists():
            try:
                json.loads(dj.read_text(encoding="utf-8"))
                n_ok += 1
            except Exception as e:
                n_bad += 1
                lineas.append(f"  JSON corrupto: {carpeta.name}/datos.json ({e})")
        hj = carpeta / "HISTORIA.json"
        if hj.exists():
            try:
                json.loads(hj.read_text(encoding="utf-8"))
            except Exception as e:
                lineas.append(f"  JSON corrupto: {carpeta.name}/HISTORIA.json ({e})")
    lineas.append(f"  Expedientes: {n_ok} OK, {n_bad} con error")
    try:
        escribir_atomico(DIAG_PATH, "\n".join(lineas) + "\n")
    except OSError:
        pass
    return lineas


# ===== 01c STORE REACTIVO Y BORRADORES (ADR-0034 / T22) =====

BORRADORES_DIR = BASE_DIR / "_BORRADORES"


class Estado:
    """Store unidireccional con revision monotona, dirty-set y suscripciones.
    ADR-0034: la UI nunca calcula ni escribe datos; solo renderiza el estado."""

    def __init__(self):
        self.revision = 0
        self._subs = {}  # clave -> [callbacks]
        self._batch = None

    def suscribir(self, clave, cb):
        self._subs.setdefault(clave, []).append(cb)

    def dispatch(self, accion, dirty):
        """Aplica una accion (accion es solo documental) y notifica claves sucias.
        Coalesce en un solo tick si ya hay un batch abierto (after(0) en UI)."""
        self.revision += 1
        if self._batch is None:
            self._batch = set()
        self._batch.update(dirty)
        return self.revision

    def flush(self):
        """Notifica las claves sucias acumuladas. Llamar desde el hilo principal."""
        if not self._batch:
            return
        claves = set(self._batch)
        self._batch = None
        for clave in claves:
            for cb in self._subs.get(clave, []):
                try:
                    cb()
                except Exception:
                    log_exc(RuntimeError(f"Error en suscriptor de '{clave}'"))


def guardar_borrador(pantalla, clave, datos):
    """Borrador local de carril B (ADR-0034): nunca toca datos.json/HISTORIA."""
    try:
        BORRADORES_DIR.mkdir(exist_ok=True)
        nombre = f"{pantalla}_{clave}.json".replace("/", "_")
        escribir_atomico(BORRADORES_DIR / nombre,
                         json.dumps(datos, ensure_ascii=False))
        return True
    except Exception:
        return False


def cargar_borrador(pantalla, clave):
    nombre = f"{pantalla}_{clave}.json".replace("/", "_")
    p = BORRADORES_DIR / nombre
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


def borrar_borrador(pantalla, clave):
    nombre = f"{pantalla}_{clave}.json".replace("/", "_")
    try:
        (BORRADORES_DIR / nombre).unlink(missing_ok=True)
    except OSError:
        pass


def guardar_config(datos):
    cfg = leer_config()
    cfg.update(datos)
    escribir_atomico(CONFIG_PATH, json.dumps(cfg, ensure_ascii=False, indent=2))


def autor_actual():
    return leer_config().get("autor", "MEDICO")


# --- P2 (ADR-0015/T9a): PIN de apertura — hash + sal, nunca en claro ---

_PIN_SAL = "maternidad-pin-v1"


def hash_pin(pin, sal=_PIN_SAL):
    return sha256_texto(f"{sal}|{pin}")


def pin_configurado(cfg=None):
    cfg = leer_config() if cfg is None else cfg
    return bool(cfg.get("pin_activo") and cfg.get("pin_hash"))


def verificar_pin(pin, cfg=None):
    cfg = leer_config() if cfg is None else cfg
    return bool(pin) and hash_pin(pin) == str(cfg.get("pin_hash", ""))


def guardar_pin(pin):
    guardar_config({"pin_hash": hash_pin(pin), "pin_activo": True})


def quitar_pin():
    cfg = leer_config()
    cfg.pop("pin_hash", None)
    cfg.pop("pin_activo", None)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# --- P7 (ADR-0015/T9c): lock monousuario best-effort ---

LOCK_PATH = BASE_DIR / "_BLOQUEO.lock"


def crear_bloqueo():
    try:
        LOCK_PATH.write_text(json.dumps(
            {"pid": os.getpid(),
             "ts": datetime.datetime.now().isoformat()}), encoding="utf-8")
        return True
    except OSError:
        return False


def limpiar_bloqueo():
    try:
        LOCK_PATH.unlink(missing_ok=True)
    except OSError:
        pass


def bloqueo_reciente(minutos=30):
    """(reciente?, mensaje) — lock vigente con menos de 'minutos' de edad."""
    try:
        info = json.loads(LOCK_PATH.read_text(encoding="utf-8"))
        ts = datetime.datetime.fromisoformat(str(info["ts"]))
    except Exception:
        return False, ""
    edad_min = (datetime.datetime.now() - ts).total_seconds() / 60.0
    if 0 <= edad_min <= minutos:
        return True, (f"Otra instancia (pid {info.get('pid', '?')}) abrio el "
                      f"sistema hace {int(edad_min)} min.")
    return False, ""


def ir_a_ruta(ruta):
    p = Path(ruta)
    if p.is_dir():
        abrir_carpeta(p)
    elif p.exists():
        abrir_archivo(p)
    else:
        raise FileNotFoundError(str(ruta))


# ===== 02 MODELO DE DATOS =====

def inicializar_estructura():
    EXPEDIENTES_DIR.mkdir(exist_ok=True)
    PLANTILLAS_DIR.mkdir(exist_ok=True)
    RESPALDOS_DIR.mkdir(exist_ok=True)
    EXPORTACIONES_DIR.mkdir(exist_ok=True)
    GENERAL_DIR.mkdir(exist_ok=True)
    BORRADORES_DIR.mkdir(exist_ok=True)
    for L in LETRAS:
        (EXPEDIENTES_DIR / L).mkdir(exist_ok=True)
    if not BITACORA_PATH.exists():
        primera = f"BITACORA MATERNIDAD v{VERSION}|sistema|inicio"
        with open(BITACORA_PATH, "a", encoding="utf-8") as f:
            f.write(f"{primera}|{sha256_texto('0' * 64 + primera)}\n")
    if not CONFIG_PATH.exists():
        guardar_config({"autor": "MEDICO", "creado": f"{fecha_hoy()} {hora_ahora()}"})
    instr = PLANTILLAS_DIR / "INSTRUCCIONES.txt"
    if not instr.exists():
        with open(instr, "w", encoding="utf-8") as f:
            f.write(
                "COLOCA AQUI TUS PLANTILLAS (.docx, .xlsx, .txt)\n"
                "==============================================\n"
                "- El sistema las detecta y las usa para generar documentos.\n"
                "- Marcadores disponibles (escribelos seguidos, sin saltos):\n"
                "  {{PACIENTE}}  {{APELLIDOS}}  {{NOMBRES}}  {{CURP}}\n"
                "  {{FECHA_NAC}} {{FECHA}}  {{DOCTOR}}  {{TELEFONO}}\n"
                "  {{CORREO}}  {{EPISODIO}}  {{EXPEDIENTE}}\n"
                "  {{FUM}}  {{FPP}}  {{EG}}   <- fechas del embarazo activo\n"
                "    (la FUM se captura en 'Nuevo embarazo'; FPP y EG\n"
                "     se calculan solas, nunca se escriben a mano)\n"
                "  Tier A: {{EDAD}} {{GP_A}} {{ALERGIAS}} {{GRUPO}} {{PARTO_ANT}}\n"
                "    (se rellenan solos desde la Ficha y Datos del embarazo)\n"
                "  Controles: {{PESO}} {{TA}} {{AU}} {{FCF}}\n"
                "    (ultimo valor registrado en el boton 'Controles')\n"
                "- Progreso quimico: nombra tu plantilla PROGRESO_BASE.xlsx\n"
                "  y el sistema creara el archivo vivo de cada embarazo.\n"
            )
    if not LEEME_PATH.exists():
        with open(LEEME_PATH, "w", encoding="utf-8") as f:
            f.write(
                "SISTEMA DE EXPEDIENTES MATERNIDAD\n"
                "================================\n"
                "1. Ejecuta: python app.py  (o Iniciar.bat en Windows)\n"
                "2. Coloca tus plantillas .docx/.xlsx en _PLANTILLAS\n"
                "3. Registrar -> buscar -> generar documento.\n"
                "4. Respaldos: menu lateral. Verificar y Restaurar incluidos.\n"
                "5. Todo offline. Tus archivos viven en ESTA carpeta.\n"
                "6. Formatos abiertos: puedes abrir tus expedientes sin el\n"
                "   software (Word, Excel, cualquier editor de texto).\n"
                "Nota: soporte normativo, no certificacion oficial.\n"
            )


def clave_paciente(nombre_norm, fecha_iso):
    return f"{nombre_norm}_{fecha_iso.replace('-', '')}"


def ruta_por_clave(clave):
    return EXPEDIENTES_DIR / clave[0] / clave


# ===== 03 SERVICIO PACIENTE =====

def cargar_paciente(ruta):
    """Carga datos.json; si esta corrupto intenta .bak1/.bak2/.bak3 (ADR-0033)."""
    archivo = Path(ruta) / "datos.json"
    try:
        with open(archivo, encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError:
        for suf in (".bak1", ".bak2", ".bak3"):
            bak = Path(str(archivo) + suf)
            if bak.exists():
                try:
                    with open(bak, encoding="utf-8") as f:
                        d = json.load(f)
                    d["degradado"] = True
                    return d
                except Exception:
                    continue
        raise


def guardar_paciente(ruta, datos):
    datos["actualizado"] = f"{fecha_hoy()} {hora_ahora()}"
    escribir_json_atomico(Path(ruta) / "datos.json", datos, rotar=True)


def crear_episodio(datos, anio=None, fecha_ref=None, fum=None):
    """fecha_ref (ISO) tiene precedencia sobre anio para determinar el anio del id."""
    anio_ref = None
    if fecha_ref:
        try:
            anio_ref = datetime.date.fromisoformat(fecha_ref).year
        except Exception:
            anio_ref = None
    anio = anio_ref or anio or datetime.date.today().year
    n = len(datos.get("episodios", [])) + 1
    eid = f"E{n:02d}_{anio}"
    nuevo = {"id": eid, "anio": int(anio), "fecha": fecha_hoy()}
    if fecha_ref:
        nuevo["fecha_referencia"] = fecha_ref
    if fum:
        nuevo["fum"] = fum
    datos.setdefault("episodios", []).append(nuevo)
    datos["episodio_activo"] = eid
    return eid


def crear_paciente(ap_p, ap_m, nombres, fecha_txt, curp="", ventana_callback=None):
    nombre_completo = " ".join(x.strip() for x in (ap_p, ap_m, nombres) if x.strip()).upper()
    if not nombre_completo:
        return False, "El nombre no puede estar vacio.", None, None
    fecha_iso = parsear_fecha(fecha_txt)
    if not fecha_iso:
        return False, "Fecha invalida. Ejemplo: 15-MAR-1985, 15/03/1985 o 1985-03-15.", None, None
    if fecha_iso > fecha_hoy():
        # C2: advertencia con escape, nunca bloqueo duro (consistente con FUM futura)
        if not (ventana_callback and ventana_callback(
                f"La fecha de nacimiento ({fecha_legible(fecha_iso)}) es posterior a hoy.\n"
                "¿Deseas continuar de todos modos?")):
            return False, "La fecha de nacimiento es posterior a hoy.", None, None
    ok_c, msg_c = validar_curp(curp)
    if not ok_c:
        if not ventana_callback or not ventana_callback(f"{msg_c}\n¿Deseas continuar sin CURP valida?"):
            return False, msg_c, None, None
        curp = ""
    nombre_norm = normalizar(nombre_completo)
    clave = clave_paciente(nombre_norm, fecha_iso)
    ruta = ruta_por_clave(clave)

    if ruta.exists():
        try:
            existente = cargar_paciente(ruta)
        except Exception:
            existente = {}
        curp_existente = (existente.get("curp") or "").strip().upper()
        curp_nueva = (curp or "").strip().upper()
        crear_separado = False
        if curp_nueva and curp_existente and curp_nueva != curp_existente:
            crear_separado = bool(ventana_callback and ventana_callback(
                "Ya existe un expediente con ese nombre y fecha pero con OTRA CURP.\n"
                "Es probablemente otra paciente (homónima).\n\n¿Crear expediente separado?"
            ))
            if not crear_separado:
                return False, "Ya existe (misma identidad).", str(ruta), existente
        elif curp_nueva and not curp_existente:
            crear_separado = bool(ventana_callback and ventana_callback(
                "Ya existe un expediente con ese nombre y fecha SIN CURP registrada.\n\n"
                "Si es OTRA paciente (homónima): Si -> expediente separado.\n"
                "Si es la MISMA paciente: No -> se guarda la CURP en su expediente."
            ))
            if not crear_separado:
                existente["curp"] = curp_nueva
                guardar_paciente(ruta, existente)
                bitacora(f"CURP registrada en expediente existente: "
                         f"{existente.get('nombre_completo', clave)}", autor=autor_actual())
                return False, "Ya existe", str(ruta), existente
        else:
            return False, "Ya existe", str(ruta), existente
        if crear_separado:
            i = 2
            while ruta_por_clave(f"{clave}_{i}").exists():
                i += 1
            clave = f"{clave}_{i}"
            ruta = ruta_por_clave(clave)

    try:
        ruta.mkdir(parents=True, exist_ok=False)
        eid = crear_episodio({"episodios": []})
        datos = {
            "clave": clave,
            "nombre_completo": nombre_completo,
            "ap_paterno": ap_p.strip().upper(),
            "ap_materno": ap_m.strip().upper(),
            "nombres": nombres.strip().upper(),
            "fecha_nacimiento": fecha_iso,
            "curp": (curp or "").strip().upper(),
            "telefono": "",
            "correo": "",
            "medico_remitente": "",
            "notas": "",
            "episodios": [{"id": eid, "anio": datetime.date.today().year, "fecha": fecha_hoy()}],
            "episodio_activo": eid,
            "creado": f"{fecha_hoy()} {hora_ahora()}",
            "actualizado": f"{fecha_hoy()} {hora_ahora()}",
        }
        guardar_paciente(ruta, datos)
        (ruta / eid).mkdir(exist_ok=True)
        with open(ruta / "HISTORIAL.txt", "w", encoding="utf-8") as f:
            f.write(
                "HISTORIAL DE PACIENTE\n" + "=" * 40 + "\n"
                f"PACIENTE: {nombre_completo}\n"
                f"CURP: {datos['curp'] or '(sin registrar)'}\n"
                f"FECHA NAC: {fecha_legible(fecha_iso)}\n"
                f"APERTURA: {fecha_hoy()} {hora_ahora()}\n"
                f"EXPEDIENTE: {clave}\n" + "=" * 40 + "\n"
            )
        bitacora(f"Nueva paciente: {nombre_completo} ({clave})", autor=autor_actual())
        return True, f"Expediente creado: {nombre_completo}", str(ruta), datos
    except Exception as e:
        log_exc(e)
        return False, f"Error al crear: {e}", None, None


def listar_pacientes():
    res = []
    for L in LETRAS:
        base = EXPEDIENTES_DIR / L
        if not base.is_dir():
            continue
        try:
            for entry in sorted(base.iterdir()):
                if entry.is_dir() and (entry / "datos.json").exists():
                    res.append((entry.name, str(entry)))
        except Exception:
            continue
    return res


def listar_recientes(n=10):
    items = []
    for nombre, ruta in listar_pacientes():
        try:
            m = os.path.getmtime(Path(ruta) / "datos.json")
        except OSError:
            m = 0
        items.append((m, nombre, ruta))
    items.sort(reverse=True)
    return [(nom, ruta) for _, nom, ruta in items[:n]]


def nombre_mostrar(ruta):
    try:
        return cargar_paciente(ruta).get("nombre_completo", Path(ruta).name)
    except Exception:
        return Path(ruta).name


def guardar_datos_secundarios(ruta, telefono, correo, curp, medico, notas,
                              alergias=None, grupo_sanguineo=None):
    """None en alergias/grupo = no tocar la clave (retrocompatibilidad T7)."""
    datos = cargar_paciente(ruta)
    datos["telefono"] = telefono.strip()
    datos["correo"] = correo.strip().lower()
    if curp:
        datos["curp"] = curp.strip().upper()
    datos["medico_remitente"] = medico.strip()
    datos["notas"] = notas.strip()
    if alergias is not None:
        datos["alergias"] = alergias.strip()
    if grupo_sanguineo is not None:
        datos["grupo_sanguineo"] = grupo_sanguineo.strip()
    guardar_paciente(ruta, datos)
    bitacora(f"Datos adicionales actualizados: {datos['nombre_completo']}", autor=autor_actual())
    return True, "Datos guardados."


def nuevo_episodio(ruta, anio=None, fecha_ref=None, fum=None):
    datos = cargar_paciente(ruta)
    eid = crear_episodio(datos, anio, fecha_ref=fecha_ref, fum=fum)
    guardar_paciente(ruta, datos)
    (Path(ruta) / eid).mkdir(exist_ok=True)
    bitacora(f"Nuevo episodio {eid}: {datos['nombre_completo']}", autor=autor_actual())
    return eid


def episodio_dir(ruta, datos=None):
    datos = datos or cargar_paciente(ruta)
    eid = datos.get("episodio_activo")
    if not eid and datos.get("episodios"):
        eid = datos["episodios"][-1]["id"]
    if not eid:
        eid = crear_episodio(datos)
        guardar_paciente(ruta, datos)
    d = Path(ruta) / eid
    if not d.exists():
        d.mkdir(exist_ok=True)
    return d


def fijar_episodio(ruta, eid):
    datos = cargar_paciente(ruta)
    ids = [e["id"] for e in datos.get("episodios", [])]
    if eid in ids:
        datos["episodio_activo"] = eid
        guardar_paciente(ruta, datos)
        return True
    return False


def episodio_contexto(datos):
    """Episodio de contexto: el activo; si no esta en la lista, el ultimo (F0/M2)."""
    activo = datos.get("episodio_activo")
    epis = datos.get("episodios") or []
    for e in epis:
        if e.get("id") == activo:
            return e
    return epis[-1] if epis else None


def actualizar_episodio(ruta, eid, campos):
    """Actualiza campos de un episodio (fum, fecha_referencia, ...) en datos.json."""
    try:
        datos = cargar_paciente(ruta)
        for e in datos.get("episodios", []):
            if e.get("id") == eid:
                for k, v in campos.items():
                    if v is None:
                        e.pop(k, None)
                    else:
                        e[k] = v
                guardar_paciente(ruta, datos)
                bitacora(f"Episodio {eid} actualizado ({', '.join(campos)}): "
                         f"{datos['nombre_completo']}", autor=autor_actual())
                return True
        return False
    except Exception as e:
        log_exc(e)
        return False


# --- C-16 (v2.1, ADR-0025): controles prenatales por episodio ---

# El medico llena este catalogo con SUS rangos; vacio = sin validacion blanda.
# Ejemplo documentado (NO activo): {"peso": (35.0, 130.0), "ta_s": (70, 160),
#                                   "ta_d": (40, 110), "au": (10.0, 40.0)}
RANGOS_CONTROLES = {}

_CAMPOS_CONTROL = ("peso", "ta_s", "ta_d", "fc", "au", "fcf",
                   "temp", "fr", "spo2", "talla")


def rangos_control_fuera(campo, valor):
    """None si sin rango o dentro; mensaje si fuera del rango blando
    (C-16: confirmar, nunca bloquear duro)."""
    rango = RANGOS_CONTROLES.get(campo)
    if not rango or valor is None:
        return None
    lo, hi = rango
    if valor < lo or valor > hi:
        return f"{campo.upper()} = {valor:g} fuera del rango blando ({lo:g} - {hi:g})"
    return None


def listar_controles(ruta, eid):
    """Controles del episodio, mas reciente primero. Sin episodio -> []."""
    try:
        datos = cargar_paciente(ruta)
    except Exception:
        return []
    for e in datos.get("episodios", []):
        if e.get("id") == eid:
            ctrls = list(e.get("controles") or [])
            ctrls.sort(key=lambda c: str(c.get("fecha") or ""), reverse=True)
            return ctrls
    return []


def _preparar_control(campos):
    """Entrada de UI -> dict de control; (ctrl, None) o (None, mensaje)."""
    fecha_txt = str(campos.get("fecha") or "").strip()
    fecha_iso = parsear_fecha(fecha_txt) if fecha_txt else fecha_hoy()
    if not fecha_iso:
        return None, "Fecha invalida. Usa AAAA-MM-DD o el boton [Cal]."
    if fecha_iso > fecha_hoy():
        return None, "La fecha es posterior a hoy."
    ctrl = {"fecha": fecha_iso}
    for k in _CAMPOS_CONTROL:
        ctrl[k] = _a_numero(campos.get(k))
    ctrl["proteinuria"] = str(campos.get("proteinuria") or "").strip()
    ctrl["notas"] = str(campos.get("notas") or "").strip()
    return ctrl, None


def agregar_control(ruta, eid, campos):
    """Registra una visita en el episodio (bitacora + historial)."""
    try:
        datos = cargar_paciente(ruta)
    except Exception as e:
        return False, f"Paciente no encontrada: {e}", None
    epi = next((e for e in datos.get("episodios", []) if e.get("id") == eid), None)
    if epi is None:
        return False, f"El episodio {eid} no existe.", None
    ctrl, err = _preparar_control(campos)
    if err:
        return False, err, None
    epi.setdefault("controles", []).append(ctrl)
    guardar_paciente(ruta, datos)
    peso_txt = f"{ctrl['peso']:g} kg" if ctrl["peso"] is not None else "sin peso"
    bitacora(f"Control prenatal {ctrl['fecha']} en {eid} ({peso_txt}): "
             f"{datos['nombre_completo']}", autor=autor_actual())
    historial(ruta, f"Control prenatal {ctrl['fecha']}")
    return True, "Control registrado.", ctrl


def actualizar_control(ruta, eid, indice, campos):
    """Actualiza el control en la posicion 'indice' de la lista almacenada."""
    try:
        datos = cargar_paciente(ruta)
    except Exception as e:
        return False, f"Paciente no encontrada: {e}"
    epi = next((e for e in datos.get("episodios", []) if e.get("id") == eid), None)
    if epi is None:
        return False, f"El episodio {eid} no existe."
    ctrls = epi.get("controles") or []
    if not (0 <= indice < len(ctrls)):
        return False, "Control inexistente."
    # Merge: solo las claves enviadas se reemplazan (el formulario manda
    # todas; una llamada parcial conserva lo no enviado)
    base = dict(ctrls[indice])
    for k, v in campos.items():
        if k == "fecha" or k == "proteinuria" or k == "notas" or k in _CAMPOS_CONTROL:
            base[k] = v
    ctrl, err = _preparar_control(base)
    if err:
        return False, err
    ctrls[indice] = ctrl
    guardar_paciente(ruta, datos)
    bitacora(f"Control prenatal actualizado {ctrl['fecha']} en {eid}: "
             f"{datos['nombre_completo']}", autor=autor_actual())
    return True, "Control actualizado."


# ===== 04 SERVICIO BUSQUEDA =====

def buscar(query):
    q = normalizar(query)
    crudo = (query or "").strip().upper()
    if not q and not crudo:
        return []
    res = []
    for nombre_clave, ruta in listar_pacientes():
        if q and q in nombre_clave:
            res.append((nombre_mostrar(ruta), ruta))
            continue
        if crudo:
            try:
                d = cargar_paciente(ruta)
            except Exception:
                continue
            campos = " ".join([
                d.get("curp", ""), d.get("telefono", ""), d.get("correo", ""),
                d.get("medico_remitente", ""), d.get("nombre_completo", ""),
            ]).upper()
            if crudo in campos:
                res.append((d.get("nombre_completo", nombre_clave), ruta))
    return res


def filtrar(resultados, medico="", desde="", hasta="", curp="", telefono="",
            correo="", ocupacion="", alergias="", grupo="", eg_min="", eg_max="",
            episodios_min=""):
    """Filtra por medico, rango de nacimiento y campos de la paciente.
    alergias: "con"|"sin"|"" (vacio = sin filtro). eg_min/eg_max en semanas."""
    out = []
    medico = medico.strip().upper()
    curp = curp.strip().upper()
    telefono = telefono.strip()
    correo = correo.strip().lower()
    ocupacion = ocupacion.strip().upper()
    grupo = grupo.strip().upper()
    for nombre, ruta in resultados:
        try:
            d = cargar_paciente(ruta)
        except Exception:
            continue
        if medico and medico not in (d.get("medico_remitente") or "").upper():
            continue
        fn = d.get("fecha_nacimiento", "")
        if desde and fn and fn < desde:
            continue
        if hasta and fn and fn > hasta:
            continue
        if curp and curp not in (d.get("curp") or "").upper():
            continue
        if telefono and telefono not in (d.get("telefono") or ""):
            continue
        if correo and correo not in (d.get("correo") or "").lower():
            continue
        if ocupacion and ocupacion not in (d.get("ocupacion") or "").upper():
            continue
        if grupo and grupo not in (d.get("grupo_sanguineo") or "").upper():
            continue
        if alergias == "con" and not (d.get("alergias") or "").strip():
            continue
        if alergias == "sin" and (d.get("alergias") or "").strip():
            continue
        n_epi = len(d.get("episodios") or [])
        try:
            epi_min = int(episodios_min) if str(episodios_min).strip() else 0
        except ValueError:
            epi_min = 0
        if epi_min and n_epi < epi_min:
            continue
        if eg_min or eg_max:
            epi = episodio_contexto(d)
            fum = (epi or {}).get("fum") or ""
            eg = calcular_eg(fum) if fum else None
            if eg:
                try:
                    semanas = int(str(eg).split()[0])
                except (ValueError, IndexError):
                    semanas = None
            else:
                semanas = None
            if semanas is None:
                continue
            try:
                if eg_min and semanas < int(eg_min):
                    continue
                if eg_max and semanas > int(eg_max):
                    continue
            except ValueError:
                pass
        out.append((nombre, ruta))
    return out


# ===== 05 SERVICIO DOCUMENTAL =====

def listar_plantillas():
    try:
        return sorted(p.name for p in PLANTILLAS_DIR.iterdir() if p.is_file()
                      and not p.name.startswith("INSTRUCCIONES"))
    except Exception:
        return []


def mapa_marcadores(datos):
    epi = None
    for e in datos.get("episodios", []):
        if e.get("id") == datos.get("episodio_activo"):
            epi = e
            break
    if epi is None and datos.get("episodios"):
        epi = datos["episodios"][-1]
    fum = (epi or {}).get("fum") or ""
    fpp = calcular_fpp(fum) if fum else None
    eg = calcular_eg(fum) if fum else None
    # Tier A (v2.0): derivados y campos estructurados
    edad = edad_anios(datos.get("fecha_nacimiento", ""))
    fpa = (epi or {}).get("fecha_parto_anterior") or ""
    tipo_ant = (epi or {}).get("tipo_parto_anterior") or ""
    parto_ant = tipo_ant + (f" ({fecha_legible(fpa)})" if fpa else "")
    # C-16 (v2.1): ultimo valor REGISTRADO de cada campo (no exige que la
    # visita mas reciente tenga todos los campos)
    ctrls_m = (epi or {}).get("controles") or []

    def _ultimo_ctrl_con(*campos):
        con = [c for c in ctrls_m if all(isinstance(c.get(k), (int, float))
                                         for k in campos)]
        return (max(con, key=lambda c: str(c.get("fecha") or ""))
                if con else {})

    def _fmt_ctrl(v):
        return f"{v:g}" if isinstance(v, (int, float)) else ""
    ultimo_peso = _ultimo_ctrl_con("peso").get("peso")
    ultimo_au = _ultimo_ctrl_con("au").get("au")
    ultimo_fcf = _ultimo_ctrl_con("fcf").get("fcf")
    c_ta = _ultimo_ctrl_con("ta_s", "ta_d")
    ta_s_v, ta_d_v = c_ta.get("ta_s"), c_ta.get("ta_d")
    return {
        "{{PACIENTE}}": datos.get("nombre_completo", ""),
        "{{APELLIDOS}}": f"{datos.get('ap_paterno', '')} {datos.get('ap_materno', '')}".strip(),
        "{{NOMBRES}}": datos.get("nombres", ""),
        "{{CURP}}": datos.get("curp", ""),
        "{{FECHA_NAC}}": fecha_legible(datos.get("fecha_nacimiento", "")),
        "{{FECHA}}": fecha_legible(fecha_hoy()),
        "{{DOCTOR}}": datos.get("medico_remitente", ""),
        "{{TELEFONO}}": datos.get("telefono", ""),
        "{{CORREO}}": datos.get("correo", ""),
        "{{EPISODIO}}": datos.get("episodio_activo", ""),
        "{{EXPEDIENTE}}": datos.get("clave", ""),
        "{{FUM}}": fecha_legible(fum) if fum else "",
        "{{FPP}}": fecha_legible(fpp) if fpp else "",
        "{{EG}}": eg or "",
        "{{EDAD}}": str(edad) if edad is not None else "",
        "{{GP_A}}": (epi or {}).get("gpa") or "",
        "{{ALERGIAS}}": datos.get("alergias", "") or "",
        "{{GRUPO}}": datos.get("grupo_sanguineo", "") or "",
        "{{PARTO_ANT}}": parto_ant,
        "{{PESO}}": _fmt_ctrl(ultimo_peso),
        "{{TA}}": (f"{ta_s_v:g}/{ta_d_v:g}"
                   if isinstance(ta_s_v, (int, float)) and isinstance(ta_d_v, (int, float))
                   else ""),
        "{{AU}}": _fmt_ctrl(ultimo_au),
        "{{FCF}}": _fmt_ctrl(ultimo_fcf),
    }


def _marcador_dividido_regex(marca):
    return r"(?:<[^>]*>)*".join(re.escape(c) for c in marca)


def _reemplazar_en_xml(xml, mapa):
    for marca, valor in mapa.items():
        v = escapar_xml(valor)
        if marca in xml:
            xml = xml.replace(marca, v)
            continue
        patron = _marcador_dividido_regex(marca)
        xml = re.sub(patron, lambda m, vv=v: vv, xml, count=1)
    return xml


def _procesar_zip_docx(origen, destino, mapa):
    with zipfile.ZipFile(origen) as zin, zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            datos_zip = zin.read(info.filename)
            n = info.filename
            if n == "word/document.xml" or re.fullmatch(r"word/(header|footer)\d+\.xml", n):
                xml = datos_zip.decode("utf-8", "replace")
                datos_zip = _reemplazar_en_xml(xml, mapa).encode("utf-8")
            elif n == "xl/sharedStrings.xml" or n.startswith("xl/worksheets/sheet"):
                xml = datos_zip.decode("utf-8", "replace")
                datos_zip = _reemplazar_en_xml(xml, mapa).encode("utf-8")
            zout.writestr(info, datos_zip)


def _procesar_plantilla(origen, destino, mapa):
    ext = Path(origen).suffix.lower()
    if ext == ".docx" or ext == ".xlsx":
        _procesar_zip_docx(origen, destino, mapa)
    else:
        try:
            texto = Path(origen).read_text(encoding="utf-8")
            for marca, valor in mapa.items():
                texto = texto.replace(marca, valor)
            Path(destino).write_text(texto, encoding="utf-8")
        except Exception:
            shutil.copy2(origen, destino)


def generar_documento(ruta_paciente, nombre_plantilla, callback_existente=None,
                      nombre_forzado=None, directorio=None, mapa_extra=None):
    try:
        datos = cargar_paciente(ruta_paciente)
    except Exception as e:
        return False, f"Paciente no encontrada: {e}", None
    plantilla = PLANTILLAS_DIR / nombre_plantilla
    if not plantilla.exists():
        return False, f"La plantilla '{nombre_plantilla}' no existe.", None
    dest_dir = Path(directorio) if directorio else episodio_dir(ruta_paciente, datos)
    dest_dir.mkdir(parents=True, exist_ok=True)
    mapa = mapa_marcadores(datos)
    if mapa_extra:
        mapa.update(mapa_extra)
    base, ext = os.path.splitext(nombre_plantilla)
    slug = normalizar(base) or "DOCUMENTO"
    if nombre_forzado:
        nombre_final = nombre_forzado if ext in nombre_forzado else nombre_forzado + ext
        ruta_destino = dest_dir / nombre_final
    else:
        nombre_final = f"{timestamp()}_{slug}{ext}"
        ruta_destino = dest_dir / nombre_final
    if ruta_destino.exists():
        if callback_existente and callback_existente(ruta_destino.name):
            return True, f"Documento existente: {ruta_destino.name}", str(ruta_destino)
        i = 1
        while True:
            nombre_final = f"{timestamp()}_{slug}_{i}{ext}"
            ruta_destino = dest_dir / nombre_final
            if not ruta_destino.exists():
                break
            i += 1
    try:
        _procesar_plantilla(plantilla, ruta_destino, mapa)
        bitacora(f"Documento generado para {datos['nombre_completo']}: {nombre_final}",
                 autor=autor_actual())
        historial(ruta_paciente, f"Documento generado: {nombre_final}")
        return True, f"Documento creado: {nombre_final}", str(ruta_destino)
    except Exception as e:
        log_exc(e)
        return False, f"Error al generar: {e}", None


def plantilla_requiere_doctor(nombre_plantilla):
    p = PLANTILLAS_DIR / nombre_plantilla
    try:
        if p.suffix.lower() in (".docx", ".xlsx"):
            with zipfile.ZipFile(p) as z:
                for n in z.namelist():
                    if n.endswith(".xml") and ("document" in n or "sharedStrings" in n
                                               or n.startswith("xl/worksheets/")):
                        t = z.read(n).decode("utf-8", "replace")
                        if "{{DOCTOR}}" in t or "{{ DOCTOR }}" in t:
                            return True
            return False
        t = p.read_text(encoding="utf-8", errors="ignore")
        return "{{DOCTOR}}" in t
    except Exception:
        return False


def progreso_quimico(ruta_paciente, callback_existente=None):
    try:
        datos = cargar_paciente(ruta_paciente)
    except Exception as e:
        return False, f"Paciente no encontrada: {e}", None
    dest_dir = episodio_dir(ruta_paciente, datos)
    destino = dest_dir / "PROGRESO_QUIMICO.xlsx"
    if destino.exists():
        return True, "Progreso Quimico existente (abriendo).", str(destino)
    origen = None
    for p in PLANTILLAS_DIR.iterdir():
        if p.is_file() and p.name.upper().startswith("PROGRESO_BASE"):
            origen = p
            break
    mapa = mapa_marcadores(datos)
    try:
        if origen:
            _procesar_plantilla(origen, destino, mapa)
        else:
            _crear_xlsx_minimo(destino, [
                ["Fecha", "Glucosa", "Creatinina", "Proteinuria", "Observaciones"],
            ])
        bitacora(f"Progreso Quimico creado: {datos['nombre_completo']} ({datos.get('episodio_activo')})",
                 autor=autor_actual())
        historial(ruta_paciente, f"Progreso Quimico creado ({datos.get('episodio_activo')})")
        return True, "Progreso Quimico creado.", str(destino)
    except Exception as e:
        log_exc(e)
        return False, f"Error al crear progreso: {e}", None


def _crear_xlsx_minimo(destino, filas):
    def esc(t):
        return (str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    cols = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    rows_xml = []
    for r_i, fila in enumerate(filas, start=1):
        celdas = []
        for c_i, val in enumerate(fila[:26]):
            ref = f"{cols[c_i]}{r_i}"
            celdas.append(f'<c r="{ref}" t="inlineStr"><is><t>{esc(val)}</t></is></c>')
        rows_xml.append(f'<row r="{r_i}">' + "".join(celdas) + "</row>")
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             "<sheetData>" + "".join(rows_xml) + "</sheetData></worksheet>")
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Hoja1" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
               'relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
          'relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-'
          'officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.'
          'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet)


_NS_XLSX = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _serial_excel_a_iso(n):
    """Serial de fecha de Excel (epoca 1899-12-30) -> ISO; None si no lo es.

    Exacto para seriales >= 61 (fechas >= 1900-03-01), que es todo rango
    clinico real; antes de eso el bug historico de 1900 desfasa un dia.
    """
    try:
        v = float(n)
    except (TypeError, ValueError):
        return None
    if not (1 <= v <= 2958465):
        return None
    try:
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).isoformat()
    except (OverflowError, ValueError):
        return None


def leer_progreso_quimico(ruta_archivo):
    """Lee el Progreso Quimico (.xlsx) del episodio. Solo lectura (P3): nunca escribe.

    Soporta inlineStr (archivo creado por el sistema) y sharedStrings/numericos
    (archivo reescrito por Excel/LibreOffice tras editar). Devuelve
    (ok, msg, encabezados, filas).
    """
    try:
        with zipfile.ZipFile(ruta_archivo) as z:
            nombres = z.namelist()
            compartidas = []
            if "xl/sharedStrings.xml" in nombres:
                raiz_s = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in raiz_s.iter(f"{_NS_XLSX}si"):
                    compartidas.append("".join(t.text or "" for t in si.iter(f"{_NS_XLSX}t")))
            hoja = next((n for n in nombres
                         if n.startswith("xl/worksheets/") and n.endswith(".xml")), None)
            if not hoja:
                return False, "El archivo no tiene hoja de calculo.", None, None
            raiz = ET.fromstring(z.read(hoja))
    except zipfile.BadZipFile:
        return False, "El Progreso Quimico no es un .xlsx valido.", None, None
    except Exception as e:
        return False, f"No se pudo leer el Progreso Quimico: {e}", None, None

    def col_idx(ref):
        letras = "".join(ch for ch in (ref or "") if ch.isalpha())
        n = 0
        for ch in letras:
            n = n * 26 + (ord(ch.upper()) - 64)
        return n - 1

    matriz = []
    for row in raiz.iter(f"{_NS_XLSX}row"):
        celdas = {}
        for c in row.findall(f"{_NS_XLSX}c"):
            t = c.get("t", "n")
            valor = None
            if t == "inlineStr":
                is_el = c.find(f"{_NS_XLSX}is")
                if is_el is not None:
                    valor = "".join(x.text or "" for x in is_el.iter(f"{_NS_XLSX}t"))
            else:
                v_el = c.find(f"{_NS_XLSX}v")
                bruto = v_el.text if v_el is not None else None
                if bruto is None:
                    valor = None
                elif t == "s":
                    try:
                        valor = compartidas[int(bruto)]
                    except Exception:
                        valor = None
                else:
                    try:
                        valor = float(bruto)
                    except ValueError:
                        valor = bruto
            if valor is not None and valor != "":
                celdas[col_idx(c.get("r", ""))] = valor
        if celdas:
            matriz.append([celdas.get(i) for i in range(max(celdas) + 1)])
    if not matriz:
        return False, "El Progreso Quimico esta vacio.", None, None
    encabezados = [str(h) if h is not None else "" for h in matriz[0]]
    filas = []
    for fila in matriz[1:]:
        f = list(fila) + [None] * (len(encabezados) - len(fila))
        if isinstance(f[0], float):
            iso = _serial_excel_a_iso(f[0])
            if iso:
                f[0] = iso
        filas.append(f[:len(encabezados)])
    return True, f"{len(filas)} filas leidas.", encabezados, filas


def _a_numero(v):
    """Coercion a float para celdas numericas (incluye texto '70.5'/'70,5')."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None
    return None


def series_progreso(encabezados, filas, max_series=8):
    """Columnas numericas del Progreso -> [(titulo, [(etiqueta_x, valor), ...])].

    Alimenta los graficos de progreso (C-17). La primera columna es el eje X
    (fecha). Sin interpretacion clinica: solo valores que el medico ya registro.
    """
    series = []
    if not encabezados or not filas:
        return series
    for i, titulo in enumerate(encabezados[1:], start=1):
        pts = []
        for n, fila in enumerate(filas):
            v = _a_numero(fila[i] if i < len(fila) else None)
            if v is None:
                continue
            etiqueta = str(fila[0]) if fila and fila[0] is not None else str(n + 1)
            pts.append((etiqueta, v))
        if pts:
            series.append((str(titulo) or f"Columna {i + 1}", pts))
    return series[:max_series]


def generar_sin_paciente(nombre_doc, nombre_plantilla):
    plantilla = PLANTILLAS_DIR / nombre_plantilla
    if not plantilla.exists():
        return False, f"La plantilla '{nombre_plantilla}' no existe.", None
    ext = Path(nombre_plantilla).suffix
    base = normalizar(nombre_doc) or "DOCUMENTO"
    GENERAL_DIR.mkdir(exist_ok=True)
    nombre_final = f"{timestamp()}_{base}{ext}"
    destino = GENERAL_DIR / nombre_final
    mapa = {
        "{{FECHA}}": fecha_legible(fecha_hoy()),
        "{{PACIENTE}}": "(sin paciente)",
    }
    try:
        _procesar_plantilla(plantilla, destino, mapa)
        bitacora(f"Documento general generado: {nombre_final}", autor=autor_actual())
        return True, f"Documento creado: {nombre_final}", str(destino)
    except Exception as e:
        log_exc(e)
        return False, f"Error: {e}", None


# ===== 06 SERVICIO AUDITORIA =====

def bitacora(mensaje, autor=None):
    try:
        autor = autor or autor_actual()
        prev = "0" * 64
        if BITACORA_PATH.exists():
            lineas = BITACORA_PATH.read_text(encoding="utf-8").splitlines()
            for ln in reversed(lineas):
                partes = ln.split("|")
                if len(partes) >= 4:
                    prev = partes[-1]
                    break
        msg = str(mensaje).replace("|", "/")
        linea = f"{fecha_hoy()} {hora_ahora()}|{autor}|{msg}"
        h = sha256_texto(prev + linea)
        with open(BITACORA_PATH, "a", encoding="utf-8") as f:
            f.write(f"{linea}|{h}\n")
    except Exception as e:
        log_exc(e)


def verificar_bitacora():
    if not BITACORA_PATH.exists():
        return False, "No hay bitacora."
    prev = "0" * 64
    total = 0
    try:
        for i, ln in enumerate(BITACORA_PATH.read_text(encoding="utf-8").splitlines(), start=1):
            partes = ln.split("|")
            if len(partes) < 4:
                return False, f"Linea {i} malformada (posible alteracion)."
            fecha, autor, msg = partes[0], partes[1], "|".join(partes[2:-1])
            h_reg = partes[-1]
            linea = f"{fecha}|{autor}|{msg}"
            h_calc = sha256_texto(prev + linea)
            if h_calc != h_reg:
                return False, f"ALTERACION detectada en linea {i}."
            prev = h_reg
            total += 1
        return True, f"Bitacora integra ({total} registros verificados)."
    except Exception as e:
        return False, f"Error al verificar: {e}"


def historial(ruta_paciente, mensaje):
    try:
        with open(Path(ruta_paciente) / "HISTORIAL.txt", "a", encoding="utf-8") as f:
            f.write(f"{fecha_hoy()} {hora_ahora()} - {mensaje}\n")
    except Exception as e:
        log_exc(e)


# ===== 07 SERVICIO RESPALDO =====

COMPONENTES_DATOS = ["EXPEDIENTES", "_PLANTILLAS", "_CONFIG.json", "_BITACORA.log", "_ERRORES.log"]


# --- P1 (ADR-0015/T9b): cifrado opcional de respaldos (cryptography) ---

CRYPTO_BIN = "RESPALDO_CIFRADO.bin"


def _clave_desde_password(password, sal=b"maternidad-respaldo-v1"):
    """PBKDF2-SHA256 -> clave urlsafe base64 (para Fernet)."""
    kdf = hashlib.pbkdf2_hmac("sha256", str(password).encode("utf-8"), sal,
                              200_000, dklen=32)
    return base64.urlsafe_b64encode(kdf)


def _zip_en_memoria(dir_raiz):
    buf = io.BytesIO()
    raiz = Path(dir_raiz)
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(raiz.rglob("*")):
            if p.is_file():
                z.write(p, str(p.relative_to(raiz)))
    return buf.getvalue()


def _cifrar_respaldo(destino, password):
    """Zip del respaldo -> un solo .bin cifrado; elimina las copias en claro."""
    destino = Path(destino)
    if not CRYPTO_DISPONIBLE:
        return False, "cryptography no instalada (el respaldo queda en claro)."
    datos_zip = _zip_en_memoria(destino)
    fernet = Fernet(_clave_desde_password(password))
    (destino / CRYPTO_BIN).write_bytes(fernet.encrypt(datos_zip))
    for comp in COMPONENTES_DATOS:
        p = destino / comp
        if p.is_dir():
            shutil.rmtree(p, ignore_errors=True)
        elif p.exists():
            p.unlink()
    manif = destino / "MANIFIESTO_RESPALDO.json"
    if manif.exists():
        manif.unlink()
    return True, "Respaldo cifrado."


def _descifrar_a_temp(destino, password):
    """Descifra .bin -> zip -> extrae en tempdir; (ok, msg, ruta_temp|None)."""
    destino = Path(destino)
    cifrado = destino / CRYPTO_BIN
    if not cifrado.exists():
        return False, "No hay respaldo cifrado en ese destino.", None
    if not CRYPTO_DISPONIBLE:
        return False, "cryptography no instalada: no se puede descifrar.", None
    try:
        fernet = Fernet(_clave_desde_password(password))
        datos_zip = fernet.decrypt(cifrado.read_bytes())
    except Exception:
        return False, "Contrasena incorrecta o archivo danado.", None
    tmp = Path(tempfile.mkdtemp(prefix="respaldo_desc_"))
    try:
        with zipfile.ZipFile(io.BytesIO(datos_zip)) as z:
            z.extractall(tmp)
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True)
        return False, f"Zip danado: {e}", None
    return True, "Descifrado.", str(tmp)


def respaldar(destino, clave=None):
    try:
        destino = Path(destino)
        destino.mkdir(parents=True, exist_ok=True)

        def ignorar(dirpath, nombres):
            if Path(dirpath).resolve() == RESPALDOS_DIR.resolve():
                return nombres
            return [n for n in nombres if n in ("_RESPALDOS",)]

        for comp in COMPONENTES_DATOS:
            origen = BASE_DIR / comp
            if not origen.exists():
                continue
            if origen.is_dir():
                shutil.copytree(origen, destino / comp, dirs_exist_ok=True,
                                ignore=ignorar, copy_function=shutil.copy2)
            else:
                shutil.copy2(origen, destino / comp)
        manifiesto = {"fecha": f"{fecha_hoy()} {hora_ahora()}", "archivos": {}}
        for comp in COMPONENTES_DATOS:
            origen = destino / comp
            if origen.is_dir():
                for p in sorted(origen.rglob("*")):
                    if p.is_file():
                        rel = str(p.relative_to(destino))
                        manifiesto["archivos"][rel] = sha256_archivo(p)
            elif origen.exists():
                manifiesto["archivos"][comp] = sha256_archivo(origen)
        (destino / "MANIFIESTO_RESPALDO.json").write_text(
            json.dumps(manifiesto, indent=2, ensure_ascii=False), encoding="utf-8")
        n = len(manifiesto["archivos"])
        bitacora(f"Respaldo creado en {destino} ({n} archivos)", autor=autor_actual())
        base_msg = (f"Respaldo completado en:\n{destino}\n"
                    f"{n} archivos verificados por hash.")
        if clave:
            if not CRYPTO_DISPONIBLE:
                return True, (base_msg + "\nCifrado NO disponible (instala "
                              "cryptography); queda integridad SHA-256.")
            ok_c, msg_c = _cifrar_respaldo(destino, clave)
            if not ok_c:
                return False, f"Error al cifrar: {msg_c}"
            return True, base_msg + "\nRespaldo CIFRADO (Fernet + PBKDF2)."
        return True, base_msg
    except Exception as e:
        log_exc(e)
        return False, f"Error en respaldo: {e}"


def verificar_respaldo(destino, clave=None):
    try:
        destino = Path(destino)
        if (destino / CRYPTO_BIN).exists():
            if not clave:
                return False, "Respaldo CIFRADO: escribe la contraseña para verificar."
            ok_d, msg_d, tmp = _descifrar_a_temp(destino, clave)
            if not ok_d:
                return False, msg_d
            try:
                return verificar_respaldo(tmp, None)
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
        manifiesto_path = destino / "MANIFIESTO_RESPALDO.json"
        if not manifiesto_path.exists():
            return False, "No se encontro MANIFIESTO_RESPALDO.json en ese destino."
        manifiesto = json.loads(manifiesto_path.read_text(encoding="utf-8"))
        archivos = manifiesto.get("archivos", {})
        malos = []
        for rel, hash_esperado in archivos.items():
            p = destino / rel
            if not p.exists():
                malos.append(f"FALTA: {rel}")
            elif sha256_archivo(p) != hash_esperado:
                malos.append(f"ALTERADO: {rel}")
        if malos:
            muestra = "\n".join(malos[:10])
            return False, f"{len(malos)} archivos no coinciden:\n{muestra}"
        return True, f"Respaldo integro: {len(archivos)} archivos verificados por SHA-256."
    except Exception as e:
        return False, f"Error al verificar: {e}"


def restaurar_respaldo(destino, clave=None):
    try:
        destino = Path(destino)
        if (destino / CRYPTO_BIN).exists():
            if not clave:
                return False, "Respaldo CIFRADO: escribe la contraseña para restaurar."
            ok_d, msg_d, tmp = _descifrar_a_temp(destino, clave)
            if not ok_d:
                return False, msg_d
            try:
                return restaurar_respaldo(tmp, None)
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
        manifiesto_path = destino / "MANIFIESTO_RESPALDO.json"
        if not manifiesto_path.exists():
            return False, "Ese destino no es un respaldo valido (sin manifiesto)."
        ok, msg = verificar_respaldo(destino)
        if not ok:
            return False, f"Restauracion cancelada, respaldo no integro:\n{msg}"
        for comp in COMPONENTES_DATOS:
            origen = destino / comp
            if not origen.exists():
                continue
            actual = BASE_DIR / comp
            if origen.is_dir():
                if actual.exists():
                    shutil.rmtree(actual)
                shutil.copytree(origen, actual)
            else:
                shutil.copy2(origen, actual)
        bitacora(f"Restauracion desde {destino}", autor=autor_actual())
        return True, f"Restauracion completada desde:\n{destino}"
    except Exception as e:
        log_exc(e)
        return False, f"Error al restaurar: {e}"



# ===== 07b SERVICIO CONSULTA (ADR-0027 / T13) =====

# Umbral NEUTRAL de dias sin control: informativo, no es juicio clinico.
# El medico lo cambia aqui si su cadencia es otra (mismo patron que
# RANGOS_CONTROLES; ver NOTA_AL_MEDICO_RANGOS.md).
DIAS_SIN_CONTROL_AVISO = 42


def resumen_hoy(dias=7, umbral=None, hoy_iso=None):
    """Bloques neutrales del panel "Hoy" (ADR-0027).

    Devuelve {"hoy", "recientes": [(nombre, ruta, ultima_fecha, hace_dias)], "vencidos": [...]}
    - recientes: controles con hace_dias <= dias
    - vencidos: ultimo control con hace_dias >= umbral; sin controles se usa
      la fecha de creacion del expediente como referencia.
    Fechas puras: no emite juicio clinico.
    """
    umbral = DIAS_SIN_CONTROL_AVISO if umbral is None else int(umbral)
    hoy = (datetime.date.fromisoformat(hoy_iso) if hoy_iso
           else datetime.date.today())
    recientes, vencidos = [], []
    for nombre, ruta in listar_pacientes():
        try:
            d = cargar_paciente(ruta)
        except Exception:
            continue
        nombre_mostrado = d.get("nombre_completo", nombre)
        fechas = []
        for epi in d.get("episodios") or []:
            for c in epi.get("controles") or []:
                if c.get("fecha"):
                    fechas.append(str(c["fecha"]))
        if fechas:
            ult = max(fechas)
            try:
                delta = (hoy - datetime.date.fromisoformat(ult)).days
            except ValueError:
                continue
            if delta <= dias:
                recientes.append((nombre_mostrado, ruta, ult, delta))
            elif delta >= umbral:
                vencidos.append((nombre_mostrado, ruta, ult, delta))
        else:
            creado = str(d.get("creado", ""))[:10]
            try:
                delta = (hoy - datetime.date.fromisoformat(creado)).days
            except ValueError:
                continue
            if delta >= umbral:
                vencidos.append((nombre_mostrado, ruta, creado, delta))
    recientes.sort(key=lambda x: x[3])
    vencidos.sort(key=lambda x: -x[3])
    return {"hoy": hoy.isoformat(), "umbral": umbral,
            "recientes": recientes, "vencidos": vencidos}


def resumen_consulta(ruta):
    """Hoja de 1 pagina (texto) para la cabina; solo datos ya guardados.

    Incluye identidad, episodio de contexto, derivados gestacionales,
    Tier A, ultimos 3 controles y las ultimas filas del Progreso Quimico.
    """
    try:
        d = cargar_paciente(ruta)
    except Exception as e:
        return f"ERROR: paciente no encontrada: {e}"
    epi = episodio_contexto(d)
    fum = (epi or {}).get("fum") or ""
    fpp = calcular_fpp(fum) if fum else None
    eg = calcular_eg(fum) if fum else None
    edad = edad_anios(d.get("fecha_nacimiento", ""))
    L = []
    L.append("=" * 72)
    L.append("RESUMEN DE CONSULTA".center(72))
    L.append("=" * 72)
    L.append(f"Fecha de impresion : {fecha_legible(fecha_hoy())} {hora_ahora()}")
    L.append(f"Paciente           : {d.get('nombre_completo', '')}")
    L.append(f"Edad               : {str(edad) + ' anios' if edad is not None else '(no calculable)'}")
    L.append(f"CURP               : {d.get('curp') or '(sin registrar)'}")
    L.append(f"Telefono           : {d.get('telefono') or '-'}")
    L.append(f"Medico remitente   : {d.get('medico_remitente') or '-'}")
    L.append(f"Expediente         : {d.get('clave', '')}")
    L.append("-" * 72)
    L.append(f"EPISODIO ACTIVO    : {(epi or {}).get('id', '(ninguno)')}")
    L.append(f"FUM                : {fecha_legible(fum) if fum else '(sin FUM)'}")
    L.append(f"FPP                : {fecha_legible(fpp) if fpp else '(no calculada)'}")
    L.append(f"EG                 : {eg or '(no disponible)'}"
             + (f"   [trimestre {trimestre_desde_eg(eg)}]" if eg else ""))
    if fpp:
        dpf = dias_para_fpp(fpp)
        if dpf is not None:
            L.append(f"Parto estimado     : "
                     + (f"faltan {dpf} dias" if dpf >= 0
                        else f"hace {-dpf} dias de la FPP"))
    L.append(f"G-P-A              : {(epi or {}).get('gpa') or '(no registrado)'}")
    L.append(f"Tipo parto anterior: {(epi or {}).get('tipo_parto_anterior') or '(no registrado)'}")
    if (epi or {}).get("fecha_parto_anterior"):
        L.append(f"Fecha parto anterior: "
                 f"{fecha_legible(epi.get('fecha_parto_anterior'))}")
    L.append(f"Complicaciones previas: {(epi or {}).get('complicaciones_previas') or '(ninguna registrada)'}")
    L.append(f"Alergias           : {d.get('alergias') or '(no registradas)'}")
    L.append(f"Grupo sanguineo    : {d.get('grupo_sanguineo') or '(no registrado)'}")
    L.append("-" * 72)
    L.append("ULTIMOS CONTROLES PRENATALES")
    ctrls = listar_controles(ruta, (epi or {}).get("id", "") or "")
    if ctrls:
        for c in ctrls[:3]:
            L.append("  " + linea_control(c))
            if c.get("notas"):
                L.append(f"      notas: {c['notas']}")
    else:
        L.append("  (sin controles registrados en este episodio)")
    # Linea de tiempo del episodio (ADR-0037)
    L.append("-" * 72)
    L.append("LINEA DE TIEMPO DEL EPISODIO")
    evs = eventos_paciente(ruta, (epi or {}).get("id", ""))
    if evs:
        for ev in evs[:20]:
            L.append(f"  {ev['fecha']} {ev.get('hora','')} [{ev['tipo'].upper()}] "
                     f"{ev['titulo']}")
    else:
        L.append("  (sin eventos registrados)")
    ruta_prog = Path(ruta) / ((epi or {}).get("id") or "") / "PROGRESO_QUIMICO.xlsx"
    if ruta_prog.exists():
        ok_p, _, enc_p, filas_p = leer_progreso_quimico(ruta_prog)
        if ok_p and filas_p:
            L.append("-" * 72)
            L.append("ULTIMAS MEDICIONES DEL PROGRESO QUIMICO")
            for fila in filas_p[-3:]:
                pares = []
                for k, v in enumerate(fila):
                    if v is None or (enc_p[k] if k < len(enc_p) else "") == "":
                        continue
                    pares.append(f"{enc_p[k] if k < len(enc_p) else k}: {v:g}"
                                 if isinstance(v, (int, float))
                                 else f"{enc_p[k] if k < len(enc_p) else k}: {v}")
                if pares:
                    L.append("  " + " | ".join(pares))
    if d.get("notas"):
        L.append("-" * 72)
        L.append(f"NOTAS GENERALES    : {d['notas']}")
    L.append("=" * 72)
    L.append(f"Autor: {autor_actual()} · documento de apoyo, no certifica (ADR-0013).")
    return "\n".join(L)


# ===== 08 SERVICIO CUMPLIMIENTO =====

def generar_consentimiento(ruta_paciente, procedimiento, ventana_callback=None):
    try:
        datos = cargar_paciente(ruta_paciente)
    except Exception as e:
        return False, f"Paciente no encontrada: {e}", None
    if not procedimiento.strip():
        return False, "Escribe el procedimiento/acto clinico.", None
    dest_dir = episodio_dir(ruta_paciente, datos)
    nombre = f"CONSENTIMIENTO_{timestamp()}.txt"
    destino = dest_dir / nombre
    autor = autor_actual()
    id_prev = sha256_texto(str(datos.get("clave", "")) + "|" + nombre + "|" + fecha_hoy())[:16]
    contenido = (
        "=" * 60 + "\n"
        "CONSENTIMIENTO INFORMADO\n"
        "=" * 60 + "\n\n"
        f"Paciente: {datos['nombre_completo']}\n"
        f"CURP: {datos.get('curp') or '(no registrada)'}\n"
        f"Fecha de nacimiento: {fecha_legible(datos.get('fecha_nacimiento', ''))}\n"
        f"Expediente: {datos.get('clave', '')}\n"
        f"Episodio: {datos.get('episodio_activo', '')}\n\n"
        f"Procedimiento / acto clinico: {procedimiento.strip()}\n\n"
        "DESCRIPCION DEL PROCEDIMIENTO:\n"
        "Se me ha explicado en terminos comprensibles la naturaleza del\n"
        "procedimiento, sus riesgos, beneficios, alternativas y posibles\n"
        "complicaciones. He tenido oportunidad de preguntar y mis dudas\n"
        "fueron resueltas a mi satisfaccion.\n\n"
        "RIESGOS Y BENEFICIOS:\n"
        "Los riesgos y beneficios specificos del procedimiento descrito\n"
        "fueron expuestos por el personal medico.\n\n"
        "ALTERNATIVAS:\n"
        "Se me informo de las alternativas de tratamiento disponibles.\n\n"
        "DERECHO A RETIRAR EL CONSENTIMIENTO:\n"
        "Manifiesto que puedo retirar este consentimiento en cualquier\n"
        "momento, sin que ello afecte mi atencion medica.\n\n"
        "AUTORIZO / NO AUTORIZO (tachar lo que no aplique) el procedimiento\n"
        "descrito, otorgando este consentimiento de forma libre y voluntaria.\n\n"
        "-" * 60 + "\n"
        f"Nombre y firma de la paciente: _________________________\n\n"
        f"Nombre y firma del medico: {datos.get('medico_remitente') or autor}\n"
        f"Fecha: {fecha_legible(fecha_hoy())}  Hora: {hora_ahora()}\n"
        "Testigo: _________________________\n"
        "-" * 60 + "\n"
        f"Documento generado por el Sistema de Expedientes Maternidad\n"
        f"Autor registrado: {autor}\n"
        f"EXPEDIENTE-ID: {id_prev}\n"
    )
    destino.write_text(contenido, encoding="utf-8")
    h = sha256_archivo(destino)
    with open(destino, "a", encoding="utf-8") as f:
        f.write(f"SHA256-CONTENIDO: {h}\n")
    bitacora(f"Consentimiento informado generado: {nombre} ({datos['nombre_completo']})",
             autor=autor)
    historial(ruta_paciente, f"Consentimiento informado: {nombre}")
    return True, f"Consentimiento generado: {nombre}", str(destino)


def generar_aviso_privacidad(ruta_paciente):
    try:
        datos = cargar_paciente(ruta_paciente)
    except Exception as e:
        return False, f"Paciente no encontrada: {e}", None
    destino = Path(ruta_paciente) / "AVISO_PRIVACIDAD.txt"
    if destino.exists():
        return True, "El aviso de privacidad ya existe.", str(destino)
    contenido = (
        "=" * 60 + "\n"
        "AVISO DE PRIVACIDAD\n"
        "=" * 60 + "\n\n"
        f"Titular: {datos['nombre_completo']}\n"
        f"CURP: {datos.get('curp') or '(no registrada)'}\n"
        f"Fecha: {fecha_legible(fecha_hoy())}\n\n"
        "FINALIDADES:\n"
        "Sus datos personales y de salud seran utilizados unica y\n"
        "exclusivamente para la prestacion de servicios medicicos de\n"
        "ginecologia y obstetricia, integracion de su historia clinica,\n"
        "y actividades de vigilancia epidemiologica cuando legalmente\n"
        "proceda.\n\n"
        "DATOS RECABADOS:\n"
        "Nombre, CURP, fecha de nacimiento, telefono, correo, medico\n"
        "remitente y datos clinicos del expediente.\n\n"
        "BASE LEGAL:\n"
        "Ley General de Salud y demas disposiciones aplicables en materia\n"
        "de proteccion de datos personales en posesion de sujetos obligados.\n\n"
        "TRANSFERENCIAS:\n"
        "No se transfieren datos a terceros ni a servicios en la nube. Sus\n"
        "archivos permanecen en el equipo de este consultorio.\n\n"
        "DERECHOS ARCO Y REVOCACION:\n"
        "Puede acceder, rectificar, cancelar u oponerse al tratamiento de\n"
        "sus datos y revocar este aviso ante el responsable del consultorio.\n\n"
        "RESPONSABLE:\n"
        f"{autor_actual()} - Consultorio de Maternidad\n"
        "=" * 60 + "\n"
    )
    destino.write_text(contenido, encoding="utf-8")
    bitacora(f"Aviso de privacidad generado: {datos['nombre_completo']}", autor=autor_actual())
    historial(ruta_paciente, "Aviso de privacidad generado")
    return True, "Aviso de privacidad generado.", str(destino)


def linea_control(c):
    """('15-JUN-2026  peso 71.2 kg  TA 120/80  ...') para textos de consulta."""
    ta = ""
    if c.get("ta_s") is not None and c.get("ta_d") is not None:
        ta = f"{c['ta_s']:g}/{c['ta_d']:g}"

    def _f(v):
        return "-" if v is None else f"{v:g}"
    return (f"{fecha_legible(str(c.get('fecha') or ''))}  "
            f"peso {_f(c.get('peso'))} kg  TA {ta or '-'}  "
            f"FC {_f(c.get('fc'))}  AU {_f(c.get('au'))}  "
            f"FCF {_f(c.get('fcf'))}  prot. {c.get('proteinuria') or '-'}")


def historia_nom004_texto(ruta_paciente, campos):
    """Texto NOM-004 puro (ADR-0028/T14): (texto, error). Vista previa y
    generacion comparten EXACTAMENTE este contenido."""
    try:
        datos = cargar_paciente(ruta_paciente)
    except Exception as e:
        return None, f"Paciente no encontrada: {e}"
    # Tier A (v2.0): bloque 1.b automatico desde datos.json (nada tecleado)
    epi = episodio_contexto(datos)
    fum_b = (epi or {}).get("fum") or ""
    fpp_b = calcular_fpp(fum_b) if fum_b else None
    eg_b = calcular_eg(fum_b) if fum_b else None
    edad_b = edad_anios(datos.get("fecha_nacimiento", ""))
    fpa_b = (epi or {}).get("fecha_parto_anterior") or ""
    contenido = (
        "=" * 60 + "\n"
        "HISTORIA CLINICA (estructura NOM-004-SSA3-2012)\n"
        "=" * 60 + "\n\n"
        "1. DATOS DE IDENTIFICACION\n"
        f"   Paciente: {datos['nombre_completo']}\n"
        f"   CURP: {datos.get('curp') or '(no registrada)'}\n"
        f"   Fecha de nacimiento: {fecha_legible(datos.get('fecha_nacimiento', ''))}\n"
        f"   Expediente: {datos.get('clave', '')}  Episodio: {datos.get('episodio_activo', '')}\n"
        f"   Medic(o) referente: {datos.get('medico_remitente') or '(sin registrar)'}\n\n"
        "1.b. DATOS AUTOMATICOS (del expediente, no se capturan aqui)\n"
        f"   Edad: {str(edad_b) + ' anios' if edad_b is not None else '(no calculable)'}\n"
        f"   FUM: {fecha_legible(fum_b) if fum_b else '(no registrada)'}\n"
        f"   FPP: {fecha_legible(fpp_b) if fpp_b else '(no calculada)'}\n"
        f"   Edad gestacional al elaborar: {eg_b or '(no disponible)'}\n"
        f"   G-P-A: {datos_gpa if (datos_gpa := (epi or {}).get('gpa')) else '(no registrado)'}\n"
        f"   Alergias: {datos_alerg if (datos_alerg := datos.get('alergias')) else '(no registradas)'}\n"
        f"   Grupo sanguineo: {datos_grup if (datos_grup := datos.get('grupo_sanguineo')) else '(no registrado)'}\n"
        f"   Tipo de parto anterior: {tipo_ant_b if (tipo_ant_b := (epi or {}).get('tipo_parto_anterior')) else '(no registrado)'}\n"
        f"   Fecha parto anterior: {fecha_legible(fpa_b) if fpa_b else '(no registrada)'}\n\n"
        "1.c. ULTIMO CONTROL REGISTRADO (automatico, C-16)\n"
        + "   " + _ultima_linea_control(datos) + "\n\n"
        "2. ANTECEDENTES\n"
        f"   Heredofamiliares: {campos.get('heredofamiliares', '(no registrado)')}\n"
        f"   Personales patologicos: {campos.get('patologicos', '(no registrado)')}\n"
        f"   Gineco-obstetricos: {campos.get('gineco', '(no registrado)')}\n\n"
        "3. EXPLORACION FISICA\n"
        f"   {campos.get('exploracion', '(no registrado)')}\n\n"
        "4. ESTUDIOS / GABINETE\n"
        f"   {campos.get('estudios', '(no registrado)')}\n\n"
        "5. DIAGNOSTICO\n"
        f"   {campos.get('diagnostico', '(no registrado)')}\n\n"
        "6. PLAN Y TRATAMIENTO\n"
        f"   {campos.get('plan', '(no registrado)')}\n\n"
        "7. EVOLUCION\n"
        f"   {campos.get('evolucion', '(no registrado)')}\n\n"
        "-" * 60 + "\n"
        f"Fecha de elaboracion: {fecha_legible(fecha_hoy())} {hora_ahora()}\n"
        f"Elaboro: {autor_actual()}\n"
        "Nota: soporte normativo; la validez oficial depende de la firma\n"
        "del profesional conforme a la normatividad aplicable.\n"
        "=" * 60 + "\n"
    )
    return contenido, None


def _ultima_linea_control(datos):
    epi = episodio_contexto(datos)
    ctrls = (epi or {}).get("controles") or []
    if not ctrls:
        return "(sin controles registrados en este episodio)"
    ultimo = max(ctrls, key=lambda c: str(c.get("fecha") or ""))
    return linea_control(ultimo)


def generar_historia_nom004(ruta_paciente, campos):
    contenido, err = historia_nom004_texto(ruta_paciente, campos)
    if contenido is None:
        return False, err, None
    try:
        datos = cargar_paciente(ruta_paciente)
    except Exception as e:
        return False, f"Paciente no encontrada: {e}", None
    dest_dir = episodio_dir(ruta_paciente, datos)
    nombre = f"HISTORIA_CLINICA_{timestamp()}.txt"
    destino = dest_dir / nombre
    destino.write_text(contenido, encoding="utf-8")
    bitacora(f"Historia clinica NOM-004 generada: {nombre} ({datos['nombre_completo']})",
             autor=autor_actual())
    historial(ruta_paciente, f"Historia clinica: {nombre}")
    return True, f"Historia clinica generada: {nombre}", str(destino)


def documentos_episodio(ruta_paciente, eid):
    """(historias, consentimientos, progreso|None, dir) del episodio — T14."""
    base = Path(ruta_paciente) / str(eid or "")
    try:
        historias = sorted(base.glob("HISTORIA_CLINICA_*.txt"),
                           key=lambda p: p.name, reverse=True)
        consentimientos = sorted(base.glob("CONSENTIMIENTO_*.txt"),
                                 key=lambda p: p.name, reverse=True)
    except OSError:
        historias, consentimientos = [], []
    prog = base / "PROGRESO_QUIMICO.xlsx"
    return {"historias": historias, "consentimientos": consentimientos,
            "progreso": str(prog) if prog.exists() else None, "dir": str(base)}


# ===== 08b SERVICIO HISTORIA ESTRUCTURADA (C-18 / ADR-0030, T16) =====

def historia_default():
    return {"version": 1, "paciente": {}, "notas": [],
            "alertas": [], "calculos": []}


def ruta_historia(ruta_paciente):
    return Path(ruta_paciente) / "HISTORIA.json"


def cargar_historia(ruta_paciente):
    """B67: dict del HISTORIA.json o default vacio (retrocompatible)."""
    p = ruta_historia(ruta_paciente)
    try:
        h = json.loads(p.read_text(encoding="utf-8"))
        if not isinstance(h, dict):
            raise ValueError("no es un objeto")
        h.setdefault("paciente", {})
        h.setdefault("notas", [])
        h.setdefault("alertas", [])
        h.setdefault("calculos", [])
        return h
    except FileNotFoundError:
        return historia_default()
    except Exception as e:
        log_exc(e)
        return historia_default()


def guardar_historia(ruta_paciente, h, versionar=False, motivo=""):
    """B68: escribe HISTORIA.json + bitacora + historial.
    A2: versionar=True guarda snapshot previo de notas modificadas."""
    try:
        p = ruta_historia(ruta_paciente)
        if versionar:
            prev = cargar_historia(ruta_paciente)
            for n_nuevo in h.get("notas") or []:
                n_id = n_nuevo.get("id")
                for n_prev in prev.get("notas") or []:
                    if n_prev.get("id") == n_id and n_prev.get("datos") != n_nuevo.get("datos"):
                        n_nuevo.setdefault("versiones", []).append({
                            "fecha": fecha_hoy(), "hora": hora_ahora(),
                            "autor": autor_actual(), "motivo": motivo,
                            "datos_previos": n_prev.get("datos", {})})
                        n_nuevo["versiones"] = n_nuevo["versiones"][-5:]
        escribir_json_atomico(p, h)
        nombre = cargar_paciente(ruta_paciente).get("nombre_completo", "")
        bitacora(f"Historia estructurada (A-M) guardada: {nombre}", autor=autor_actual())
        historial(ruta_paciente, "Historia estructurada A-M guardada")
        return True, "Historia guardada."
    except Exception as e:
        log_exc(e)
        return False, f"Error al guardar la historia: {e}"


def _valor_global_historia(h, seccion, item_id, nivel):
    """Valor del item a nivel global: paciente; o el primer valor entre notas."""
    if nivel == "paciente":
        return (h.get("paciente") or {}).get(seccion, {}).get(item_id, "")
    for n in h.get("notas") or []:
        v = (n.get("datos") or {}).get(seccion, {}).get(item_id, "")
        if isinstance(v, dict):
            if v.get("e") or any(str(x).strip() for k, x in v.items() if k != "e"):
                return v
        elif v not in ("", None):
            return v
    return ""


def progreso_historia(h):
    """B69: {seccion: (investigados, total)}. '' = sin investigar (cuenta mal);
    'no' cuenta como investigado. Items 'auto' siempre cuentan (derivados)."""
    out = {}
    for s in SECCIONES_HISTORIA:
        inv = tot = 0
        for it in s["items"]:
            tot += 1
            if it["t"] == "auto":
                inv += 1
                continue
            v = _valor_global_historia(h, s["id"], it["id"], s["nivel"])
            if isinstance(v, dict):
                if v.get("e") or any(str(x).strip() for k, x in v.items() if k != "e"):
                    inv += 1
            elif v not in ("", None):
                inv += 1
        out[s["id"]] = (inv, tot)
    return out


def contexto_historia(ruta_paciente):
    """Bundle de datos derivados para items 'auto' (editor y render)."""
    datos = cargar_paciente(ruta_paciente)
    epi = episodio_contexto(datos)
    eid = (epi or {}).get("id", "")
    ctrls = listar_controles(ruta_paciente, eid)
    # ultimo valor REGISTRADO por campo (mismo criterio que los marcadores)
    ultimo = {}
    if ctrls:
        ultimo["fecha"] = max(str(c.get("fecha") or "") for c in ctrls)
        for campo in ("temp", "ta_s", "ta_d", "fc", "fr", "spo2", "peso"):
            con = [c for c in ctrls if c.get(campo) is not None]
            if con:
                peor = max(con, key=lambda c: str(c.get("fecha") or ""))
                ultimo[campo] = peor[campo]
    docs = documentos_episodio(ruta_paciente, eid)
    return {"ruta": ruta_paciente, "datos": datos, "epi": epi, "ctrl": ultimo,
            "docs": docs, "autor": autor_actual(),
            "ahora": f"{fecha_hoy()} {hora_ahora()}"}


def auto_historia(seccion, item, ctx):
    """Valor derivado de un item tipo 'auto'; '' si no tiene fuente aun."""
    d = ctx["datos"]
    c = ctx.get("ctrl") or {}

    def _n(v):
        return f"{v:g}" if isinstance(v, (int, float)) else ""
    if seccion == "A":
        return {
            "nombre": d.get("nombre_completo", ""),
            "edad": (f"{edad_anios(d.get('fecha_nacimiento', ''))} anios"
                     if edad_anios(d.get("fecha_nacimiento", "")) is not None else ""),
            "fecha_nac": fecha_legible(d.get("fecha_nacimiento", "")),
            "telefono": d.get("telefono", ""),
            "expediente": d.get("clave", ""),
            "curp": d.get("curp", ""),
        }.get(item, "")
    if seccion == "B" and item == "fecha_hora":
        return ctx["ahora"]
    if seccion == "H1":
        ta = ""
        if c.get("ta_s") is not None and c.get("ta_d") is not None:
            ta = f"{c['ta_s']:g}/{c['ta_d']:g}"
        talla = ""  # la talla no tiene fuente estructurada aun
        imc = ""
        peso = c.get("peso")
        if peso is not None:
            try:
                talla_m = float(talla) / 100 if talla else None
                if talla_m:
                    imc = f"{peso / (talla_m ** 2):.1f}"
            except ValueError:
                imc = ""
        return {"temperatura": _n(c.get("temp")), "ta": ta, "fc": _n(c.get("fc")),
                "fr": _n(c.get("fr")), "spo2": _n(c.get("spo2")),
                "peso": _n(c.get("peso")), "talla": talla, "imc": imc}.get(item, "")
    if seccion == "K":
        hay_consent = bool(ctx["docs"]["consentimientos"])
        return {"medico": ctx["autor"], "firma": "_" * 28,
                "elaboracion": ctx["ahora"],
                "consentimiento": ("Si (generado en el episodio)" if hay_consent
                                   else "No (aun sin consentimiento)"),
                "sha256": ""}.get(item, "")
    if seccion == "L" and item == "fecha_hora":
        nota = ctx.get("nota") or {}
        return str(nota.get("fecha") or "")
    if seccion == "L" and item == "firma":
        return ctx["autor"]
    return ""


def _formato_valor_historia(item, v):
    """Valor -> texto imprimible (estados, dicts de extras, crudos)."""
    if v is None or v == "":
        return "(sin investigar)" if item["t"] in ("sn", "snns", "snna", "snv",
                                                    "snc", "snh") else "-"
    estados = {"si": "Si", "no": "No", "no_sabe": "No sabe",
               "no_aplica": "No aplica"}
    if item["t"] in ("sn", "snns", "snna"):
        return estados.get(str(v), str(v))
    if item["t"] in ("snv", "snc", "snh") and isinstance(v, dict):
        base = estados.get(str(v.get("e", "")), str(v.get("e", "")) or "(sin investigar)")
        extras = [str(x) for k, x in v.items() if k != "e" and str(x).strip()]
        return f"{base} — {' / '.join(extras)}" if extras else base
    if isinstance(v, dict):
        return ", ".join(f"{k}={x}" for k, x in v.items())
    return str(v)


def nueva_nota_historia(h, tipo, eid):
    """Agrega una nota (nunca se borran; precedente C-16)."""
    nota = {"tipo": tipo, "fecha": fecha_hoy(), "hora": hora_ahora(),
            "episodio": eid or "", "datos": {}}
    h.setdefault("notas", []).append(nota)
    return len(h["notas"]) - 1

def render_historia_v2(ruta_paciente, eid=None):
    """B70 (C-18): texto A-M con estados, bloque 1.b, seccion K y SHA-256."""
    h = cargar_historia(ruta_paciente)
    ctx = contexto_historia(ruta_paciente)
    datos = ctx["datos"]
    epi = ctx["epi"]
    eid = eid or (epi or {}).get("id", "")
    base = None
    evoluciones = []
    egreso = None
    for n in h.get("notas") or []:
        if n.get("episodio") and eid and n["episodio"] != eid:
            continue
        tipo = n.get("tipo")
        if tipo == "evolucion":
            evoluciones.append(n)
        elif tipo == "egreso":
            egreso = n
        elif base is None:
            base = n
    base = base or {"datos": {}}

    L = []
    L.append("=" * 72)
    L.append("HISTORIA CLINICA COMPLETA — NOM-004-SSA3-2012 (A-M)".center(72))
    L.append("=" * 72)
    L.append(f"Paciente   : {datos.get('nombre_completo', '')}")
    L.append(f"Expediente : {datos.get('clave', '')}   Episodio: {eid or '-'}")
    L.append(f"Impresion  : {fecha_legible(fecha_hoy())} {hora_ahora()}")
    # bloque 1.b (misma fuente que la ruta rapida C-12)
    fum_b = (epi or {}).get("fum") or ""
    fpp_b = calcular_fpp(fum_b) if fum_b else None
    eg_b = calcular_eg(fum_b) if fum_b else None
    edad_b = edad_anios(datos.get("fecha_nacimiento", ""))
    L.append("")
    L.append("1.b. DATOS AUTOMATICOS DEL EXPEDIENTE")
    L.append(f"   Edad: {str(edad_b) + ' anios' if edad_b is not None else '(no calculable)'}")
    L.append(f"   FUM: {fecha_legible(fum_b) if fum_b else '(no registrada)'}"
             f"   FPP: {fecha_legible(fpp_b) if fpp_b else '(no calculada)'}"
             f"   EG: {eg_b or '(no disponible)'}"
             + (f"   Trimestre {trimestre_desde_eg(eg_b)}" if eg_b else ""))
    L.append(f"   G-P-A: {(epi or {}).get('gpa') or '(no registrado)'}"
             f"   Alergias: {datos.get('alergias') or '(ninguna)'}"
             f"   Grupo: {datos.get('grupo_sanguineo') or '(no registrado)'}")
    for s in SECCIONES_HISTORIA:
        L.append("")
        L.append(f"{s['id']}. {s['titulo'].upper()}   [{s['nivel']}]")
        if s["id"] == "L":
            if evoluciones:
                for i, n in enumerate(evoluciones, start=1):
                    ctx_n = dict(ctx)
                    ctx_n["nota"] = n
                    L.append(f"   -- Evolucion {i} ({n.get('fecha', '')} "
                             f"{n.get('hora', '')}) --")
                    src = (n.get("datos") or {}).get("L", {})
                    for it in s["items"]:
                        v = (auto_historia("L", it["id"], ctx_n)
                             if it["t"] == "auto" else src.get(it["id"], ""))
                        L.append(f"      {it['lbl']}: {_formato_valor_historia(it, v)}")
            else:
                L.append("   (sin notas de evolucion registradas)")
            continue
        if s["id"] == "M":
            if egreso:
                src = (egreso.get("datos") or {}).get("M", {})
                for it in s["items"]:
                    v = (auto_historia("M", it["id"], ctx)
                         if it["t"] == "auto" else src.get(it["id"], ""))
                    L.append(f"   {it['lbl']}: {_formato_valor_historia(it, v)}")
            else:
                L.append("   (sin nota de egreso registrada)")
            continue
        if s["nivel"] == "paciente":
            src = (h.get("paciente") or {}).get(s["id"], {})
            ctx_s = ctx
            ctx_s["nota"] = {}
        else:
            src = (base.get("datos") or {}).get(s["id"], {})
            ctx_s = dict(ctx)
            ctx_s["nota"] = base
        for it in s["items"]:
            v = (auto_historia(s["id"], it["id"], ctx_s)
                 if it["t"] == "auto" else src.get(it["id"], ""))
            L.append(f"   {it['lbl']}: {_formato_valor_historia(it, v)}")
    contenido = "\n".join(L)
    sha = sha256_texto(contenido)
    pie = [
        "",
        "=" * 72,
        "TRAZABILIDAD (K — PIE DEL DOCUMENTO)",
        f"SHA-256 del documento: {sha}",
        f"Elaboró: {autor_actual()}   Fecha: {fecha_legible(fecha_hoy())} {hora_ahora()}",
        "Nota: soporte normativo; la validez oficial depende de la firma del",
        "profesional conforme a la normatividad aplicable (ADR-0013).",
        "=" * 72,
    ]
    return contenido + "\n" + "\n".join(pie)


def generar_historia_v2(ruta_paciente, eid=None):
    """Escribe el render A-M en el episodio activo (bitacora + historial)."""
    try:
        datos = cargar_paciente(ruta_paciente)
    except Exception as e:
        return False, f"Paciente no encontrada: {e}", None
    texto = render_historia_v2(ruta_paciente, eid)
    dest_dir = episodio_dir(ruta_paciente, datos)
    nombre = f"HISTORIA_COMPLETA_{timestamp()}.txt"
    destino = dest_dir / nombre
    destino.write_text(texto, encoding="utf-8")
    bitacora(f"Historia completa A-M generada: {nombre} "
             f"({datos.get('nombre_completo', '')})", autor=autor_actual())
    historial(ruta_paciente, f"Historia completa A-M: {nombre}")
    return True, f"Historia completa generada: {nombre}", str(destino)


def exportar_checklist_xlsx(ruta_paciente):
    """B71 (C-18): checklist A-M a XLSX (Seccion, Item, Estado/valor)."""
    try:
        datos = cargar_paciente(ruta_paciente)
    except Exception as e:
        return False, f"Paciente no encontrada: {e}", None
    h = cargar_historia(ruta_paciente)
    filas = [["Seccion", "Item", "Estado/valor"]]
    for s in SECCIONES_HISTORIA:
        for it in s["items"]:
            v = _valor_global_historia(h, s["id"], it["id"], s["nivel"])
            filas.append([f"{s['id']}. {s['titulo']}", it["lbl"],
                          _formato_valor_historia(it, v)])
    EXPORTACIONES_DIR.mkdir(parents=True, exist_ok=True)
    destino = EXPORTACIONES_DIR / (f"CHECKLIST_{datos.get('clave', 'PACIENTE')}_"
                                   f"{timestamp()}.xlsx")
    try:
        _crear_xlsx_minimo(destino, filas)
        bitacora(f"Checklist A-M exportado: {destino}", autor=autor_actual())
        return True, f"Checklist exportado: {destino}", str(destino)
    except Exception as e:
        log_exc(e)
        return False, f"Error al exportar checklist: {e}", None


# ===== 08d SERVICIO ALERTAS (C-20 / ADR-0031, T18) =====

def perfil_control(ruta_paciente):
    """'embarazada' si el episodio de contexto tiene EG valida; si no 'adulto'."""
    try:
        epi = episodio_contexto(cargar_paciente(ruta_paciente))
    except Exception:
        epi = None
    fum = (epi or {}).get("fum") or ""
    eg = calcular_eg(fum) if fum else None
    return "embarazada" if eg else "adulto"


def hallazgos_control(ruta_paciente, campos):
    """(hallazgos, qsofa_pts, qsofa_msg, perfil) al guardar un control (C-20)."""
    perfil = perfil_control(ruta_paciente)
    valores = {k: _a_numero(campos.get(k)) for k in _CAMPOS_CONTROL}
    hall = evaluar_signos_vitales(valores, perfil)
    hall += evaluar_preeclampsia(valores, str(campos.get("proteinuria") or ""))
    pts, qmsg = score_qsofa({**valores,
                             "conciencia_alterada": bool(campos.get("conciencia_alterada"))})
    return hall, pts, qmsg, perfil


def _escribir_historia(ruta_paciente, h):
    escribir_json_atomico(ruta_historia(ruta_paciente), h)


def registrar_alerta(ruta_paciente, alerta):
    """B79: alta en HISTORIA.json.alertas[] (nace pendiente) + bitacora.
    A11: dedupe por (episodio, origen, hash) pendiente."""
    try:
        h = cargar_historia(ruta_paciente)
        dup = _dedupe_alerta(h, alerta)
        if dup:
            return True, dup  # ya existe pendiente identica
        ident = f"ALT-{fecha_hoy()}-{len(h.get('alertas') or []) + 1:03d}"
        registro = dict(alerta)
        registro.update({"id": ident, "fecha": fecha_hoy(), "hora": hora_ahora(),
                         "estado": "pendiente"})
        h.setdefault("alertas", []).append(registro)
        _escribir_historia(ruta_paciente, h)
        bitacora(f"Alerta registrada {ident}: {registro.get('origen', '?')} "
                 f"- {registro.get('descripcion', '')[:80]}", autor=autor_actual())
        return True, ident
    except Exception as e:
        log_exc(e)
        return False, f"Error al registrar alerta: {e}"


def atender_alerta(ruta_paciente, ident, respuesta):
    """B80: pendiente -> atendida con {texto, medico, cedula}; bitacora."""
    try:
        h = cargar_historia(ruta_paciente)
        alerta = next((a for a in (h.get("alertas") or [])
                       if a.get("id") == ident), None)
        if alerta is None:
            return False, "Alerta no encontrada."
        if alerta.get("estado") != "pendiente":
            return False, "La alerta ya no esta pendiente."
        alerta["estado"] = "atendida"
        alerta["respuesta"] = {
            "texto": str(respuesta.get("texto", "")).strip(),
            "medico": str(respuesta.get("medico", "")).strip() or autor_actual(),
            "cedula": str(respuesta.get("cedula", "")).strip(),
        }
        alerta["atendida_fecha"] = fecha_hoy()
        alerta["atendida_hora"] = hora_ahora()
        _escribir_historia(ruta_paciente, h)
        bitacora(f"Alerta atendida {ident} por {alerta['respuesta']['medico']}: "
                 f"{alerta['respuesta']['texto'][:80]}", autor=autor_actual())
        return True, "Alerta atendida."
    except Exception as e:
        log_exc(e)
        return False, f"Error al atender alerta: {e}"


def alertas_pendientes(ruta_paciente, eid=None):
    """Pendientes del episodio (o todas si eid es None)."""
    try:
        h = cargar_historia(ruta_paciente)
    except Exception:
        return []
    out = []
    for a in h.get("alertas") or []:
        if a.get("estado") != "pendiente":
            continue
        if eid and a.get("episodio") and a["episodio"] != eid:
            continue
        out.append(a)
    return out


def revisar_lab_progreso(ruta_paciente, eid=None):
    """B76 aplicado: columnas conocidas del Progreso vs UMBRALES_LAB
    (solo lectura, P3). -> (ok, msg, hallazgos)."""
    try:
        datos = cargar_paciente(ruta_paciente)
    except Exception as e:
        return False, f"Paciente no encontrada: {e}", []
    epi = episodio_contexto(datos)
    eid = eid or (epi or {}).get("id", "")
    ruta_xlsx = Path(ruta_paciente) / str(eid) / "PROGRESO_QUIMICO.xlsx"
    if not ruta_xlsx.exists():
        return False, "Este episodio aun no tiene PROGRESO_QUIMICO.xlsx.", []
    ok_l, msg_l, enc, filas = leer_progreso_quimico(ruta_xlsx)
    if not ok_l:
        return False, msg_l, []
    hallazgos = []
    for i, h_enc in enumerate(enc[1:], start=1):
        clave = next((uk for uk in UMBRALES_LAB
                      if uk in str(h_enc).strip().lower()), None)
        if not clave:
            continue
        for fila in filas:
            v = _a_numero(fila[i] if i < len(fila) else None)
            h = evaluar_lab(clave, v)
            if h:
                h["fecha_fila"] = (str(fila[0]) if fila and fila[0] is not None
                                   else "")
                hallazgos.append(h)
    msg = (f"{len(hallazgos)} valor(es) critico(s) en el Progreso Quimico."
           if hallazgos else "Sin valores criticos en las columnas conocidas.")
    return True, msg, hallazgos


# ===== 08f CONSENTIMIENTOS, ALERTAS Y RESPALDOS (ADR-0035 / T23) =====

def registrar_consentimiento(ruta, archivo, estado="vigente"):
    """A7: registra un consentimiento con estado en datos.json."""
    datos = cargar_paciente(ruta)
    datos.setdefault("consentimientos", []).append({
        "fecha": fecha_hoy(), "hora": hora_ahora(), "archivo": archivo,
        "estado": estado, "autor": autor_actual()})
    guardar_paciente(ruta, datos)
    bitacora(f"Consentimiento registrado ({estado}): {archivo}", autor=autor_actual())
    return True, "Consentimiento registrado."


def revocar_consentimiento(ruta, indice, motivo):
    """A7: revoca un consentimiento vigente."""
    datos = cargar_paciente(ruta)
    conss = datos.get("consentimientos") or []
    if not (0 <= indice < len(conss)):
        return False, "Consentimiento no encontrado."
    conss[indice]["estado"] = "revocado"
    conss[indice]["revocado_fecha"] = fecha_hoy()
    conss[indice]["revocado_motivo"] = motivo
    guardar_paciente(ruta, datos)
    bitacora(f"Consentimiento revocado ({indice}): {motivo}", autor=autor_actual())
    return True, "Consentimiento revocado."


def _dedupe_alerta(h, alerta):
    """A11: dedupe por (episodio, origen, hash descripcion) pendiente."""
    nuevo_hash = sha256_texto(str(alerta.get("descripcion", "")))
    for a in h.get("alertas") or []:
        if (a.get("estado") == "pendiente"
                and a.get("episodio") == alerta.get("episodio")
                and a.get("origen") == alerta.get("origen")
                and sha256_texto(str(a.get("descripcion", ""))) == nuevo_hash):
            return a.get("id")
    return None


# ===== 08e INTEGRIDAD CLINICA (ADR-0035 / T23) =====

ACCESOS_PATH = BASE_DIR / "_ACCESOS.log"


def validar_plausibilidad(valores):
    """A3: validacion de plausibilidad clinica. -> lista de hallazgos blando."""
    h = []
    ta_s = _a_numero(valores.get("ta_s"))
    ta_d = _a_numero(valores.get("ta_d"))
    if ta_s is not None and ta_d is not None and ta_s < ta_d:
        h.append({"campo": "ta", "nivel": "alerta",
                  "mensaje": f"TA sistólica ({ta_s:g}) < diastólica ({ta_d:g}): revisa los valores"})
    peso = _a_numero(valores.get("peso"))
    if peso is not None and peso <= 0:
        h.append({"campo": "peso", "nivel": "critico",
                  "mensaje": f"Peso ({peso:g}) no puede ser <= 0"})
    temp = _a_numero(valores.get("temp"))
    if temp is not None and not (30.0 <= temp <= 45.0):
        h.append({"campo": "temp", "nivel": "alerta",
                  "mensaje": f"Temperatura ({temp:g}°C) fuera de rango plausible (30–45)"})
    spo2 = _a_numero(valores.get("spo2"))
    if spo2 is not None and not (0 <= spo2 <= 100):
        h.append({"campo": "spo2", "nivel": "critico",
                  "mensaje": f"SpO₂ ({spo2:g}%) fuera de 0–100"})
    fc = _a_numero(valores.get("fc"))
    if fc is not None and not (20 <= fc <= 250):
        h.append({"campo": "fc", "nivel": "alerta",
                  "mensaje": f"FC ({fc:g}) fuera de 20–250 lpm"})
    return h


def registrar_lectura(ruta, contexto):
    """A5: access-log append-only con SHA encadenado. Solo hitos."""
    try:
        prev = ""
        if ACCESOS_PATH.exists():
            lineas = ACCESOS_PATH.read_text(encoding="utf-8").splitlines()
            if lineas:
                prev = lineas[-1].split("|")[-1]
        nombre = cargar_paciente(ruta).get("nombre_completo", "?") if Path(ruta).exists() else "?"
        linea = f"{fecha_hoy()} {hora_ahora()}|{contexto}|{nombre}|{autor_actual()}"
        h = sha256_texto(prev + linea)
        with open(ACCESOS_PATH, "a", encoding="utf-8") as f:
            f.write(f"{linea}|{h}\n")
        return True
    except Exception:
        return False


def fusionar_paciente(origen, destino, motivo):
    """A4: fusiona dos expedientes; episodios unidos por id; carpeta origen -> _FUSIONADOS/."""
    try:
        d_o = cargar_paciente(origen)
        d_d = cargar_paciente(destino)
    except Exception as e:
        return False, f"No se pudieron leer los expedientes: {e}"
    # unir episodios por id (sin pisar)
    ids_d = {e.get("id") for e in d_d.get("episodios", [])}
    for epi in d_o.get("episodios", []):
        if epi.get("id") not in ids_d:
            d_d.setdefault("episodios", []).append(epi)
    # campos del destino preservados; solo rellenar vacios desde origen
    for k, v in d_o.items():
        if k in ("episodios", "clave", "nombre_completo"):
            continue
        if not d_d.get(k) and v:
            d_d[k] = v
    guardar_paciente(destino, d_d)
    # mover carpeta origen a _FUSIONADOS
    fus_dir = BASE_DIR / "_FUSIONADOS"
    fus_dir.mkdir(exist_ok=True)
    destino_fus = fus_dir / f"{Path(origen).name}_{fecha_hoy()}_{hora_ahora().replace(':', '')}"
    try:
        Path(origen).rename(destino_fus)
    except OSError as e:
        return False, f"No se pudo mover la carpeta de origen: {e}"
    bitacora(f"Expediente fusionado: {Path(origen).name} -> {Path(destino).name} "
             f"(motivo: {motivo})", autor=autor_actual())
    return True, f"Fusionado en {Path(destino).name} (origen movido a _FUSIONADOS)"


def registrar_lectura_hitos(ruta):
    """Registra lectura en hitos de negocio (A5)."""
    return registrar_lectura(ruta, "hito")


def drill_restauracion(destino):
    """A8: prueba de restauracion no destructiva en tempdir. -> (ok, msg)."""
    import tempfile as _tf
    tmp = Path(_tf.mkdtemp(prefix="drill_"))
    try:
        ok, msg = verificar_respaldo(destino)
        if not ok:
            return False, f"Verificacion fallo: {msg}"
        # restaurar en temp y verificar parseo
        r = restaurar_respaldo(destino)
        shutil.rmtree(tmp, ignore_errors=True)
        return True, f"Drill OK: {msg[:120]}"
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True)
        return False, f"Drill fallo: {e}"


def deshacer_ultima_edicion(ruta, seccion, nivel):
    """A13: deshacer la ultima edicion de una seccion (si hay version previa)."""
    h = cargar_historia(ruta)
    if nivel == "paciente":
        sec = (h.get("paciente") or {}).get(seccion)
        if sec is None:
            return False, "No hay edicion que deshacer."
        h["paciente"][seccion] = {}
        guardar_historia(ruta, h)
        bitacora(f"Deshacer edicion seccion {seccion}", autor=autor_actual())
        return True, f"Deshacer de {seccion}."
    return False, "Deshacer disponible solo para secciones de paciente."


def exportar_diagnostico():
    """A18: bundle de diagnostico anonimo (sin bitacora, sin datos de pacientes)."""
    import tempfile as _tf
    tmp = Path(_tf.mkdtemp(prefix="diag_"))
    try:
        lineas = []
        lineas.append(f"Sistema Maternidad v{VERSION}")
        lineas.append(f"Fecha: {fecha_hoy()} {hora_ahora()}")
        lineas.append(f"Python: {sys.version}")
        # diag
        if DIAG_PATH.exists():
            lineas.append("--- _DIAGNOSTICO.log ---")
            lineas.append(DIAG_PATH.read_text(encoding="utf-8")[:4000])
        if ERROR_LOG_PATH.exists():
            lineas.append("--- _ERRORES.log (ultimas 50 lineas) ---")
            lineas.append("\n".join(ERROR_LOG_PATH.read_text(encoding="utf-8").splitlines()[-50:]))
        n_exp = len(list(EXPEDIENTES_DIR.glob("*/*")))
        lineas.append(f"Expedientes: {n_exp}")
        # conteos de historia/alertas (sin nombres)
        n_hist = n_alert = 0
        for carpeta in EXPEDIENTES_DIR.glob("*/*"):
            hj = carpeta / "HISTORIA.json"
            if hj.exists():
                try:
                    h = json.loads(hj.read_text(encoding="utf-8"))
                    n_hist += 1
                    n_alert += len(h.get("alertas") or [])
                except Exception:
                    pass
        lineas.append(f"HISTORIA.json: {n_hist} pacientes, {n_alert} alertas")
        destino = EXPORTACIONES_DIR / f"DIAGNOSTICO_{fecha_hoy()}_{hora_ahora().replace(':', '')}.txt"
        EXPORTACIONES_DIR.mkdir(parents=True, exist_ok=True)
        destino.write_text("\n".join(lineas), encoding="utf-8")
        return True, f"Diagnostico exportado: {destino}", str(destino)
    except Exception as e:
        return False, f"Error al exportar diagnostico: {e}", None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ===== 09b EXPORTACION ANONIMIZADA (ADR-0036 / C-21 / T24) =====

def id_anonimo(clave):
    """B82: SHA-256(clave+sal_local)[:12]. Sal nunca sale del equipo."""
    cfg = leer_config()
    sal = cfg.get("sal_anonima")
    if not sal:
        import secrets as _sec
        sal = _sec.token_hex(16)
        guardar_config({"sal_anonima": sal})
    return sha256_texto(clave + sal)[:12].upper()


def _valores_identificadores(datos):
    """Lista de valores literales que NUNCA deben aparecer en el export."""
    vals = []
    for k in ("nombre_completo", "ap_paterno", "ap_materno", "nombres",
              "clave", "curp", "telefono", "correo", "domicilio", "nss",
              "medico_remitente"):
        v = datos.get(k)
        if v and isinstance(v, str) and len(v.strip()) >= 3:
            vals.append(v.strip())
    # tokens de nombre (>=4 chars)
    nc = datos.get("nombre_completo", "")
    for tok in nc.split():
        if len(tok) >= 4:
            vals.append(tok)
    return vals


def escanear_identificadores(texto, datos):
    """B84: patrones de identificacion en el texto final. -> lista de hallazgos."""
    h = []
    t_up = texto.upper()
    # CURP patron
    import re as _re
    if _re.search(r"[A-Z]{4}\d{6}[HM][A-Z]{5}[0-9A-Z]\d", t_up):
        h.append("CURP (patron)")
    # telefono 10 digitos
    if _re.search(r"\b\d{10}\b", t_up):
        h.append("posible telefono (10 digitos)")
    # email
    if _re.search(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", t_up):
        h.append("email")
    # NSS 11 digitos
    if _re.search(r"\b\d{11}\b", t_up):
        h.append("posible NSS (11 digitos)")
    # valores literales
    for v in _valores_identificadores(datos):
        if v.upper() in t_up:
            h.append(f"valor literal: {v[:20]}")
    return h


def _scrub_texto_libre(txt, datos):
    """Enmascara identificadores en texto libre (segunda capa). No altera el resto."""
    import re as _re
    if not txt:
        return txt
    for v in _valores_identificadores(datos):
        txt = txt.replace(v, "[IDENTIFICADOR]")
        txt = txt.replace(v.upper(), "[IDENTIFICADOR]")
        txt = txt.replace(v.lower(), "[IDENTIFICADOR]")
    txt = _re.sub(r"[A-Za-z]{4}\d{6}[HhMm][A-Za-z]{5}[0-9A-Za-z]\d", "[CURP]", txt)
    txt = _re.sub(r"\b\d{10}\b", "[TEL]", txt)
    txt = _re.sub(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", "[EMAIL]", txt)
    return txt


def exportar_historia_anonima(ruta, pregunta="", incluir_csv=True):
    """B83: exporta historial clinico anonimizado (allowlist + puerta)."""
    try:
        datos = cargar_paciente(ruta)
    except Exception as e:
        return False, f"Paciente no encontrada: {e}", None
    try:
        historia = cargar_historia(ruta)
    except Exception:
        historia = historia_default()
    epi = episodio_contexto(datos)
    ctrls = listar_controles(ruta, (epi or {}).get("id", ""))
    docs = documentos_episodio(ruta, (epi or {}).get("id", ""))
    iid = id_anonimo(datos.get("clave", "X"))
    edad = edad_anios(datos.get("fecha_nacimiento", ""))

    L = []
    L.append(f"# Historial clinico anonimizado — ID: {iid}")
    L.append("> Sin nombre, CURP, telefono, domicilio ni identificadores.")
    L.append("> Datos medicos completos para analisis.")
    L.append(f"> Generado: {fecha_legible(fecha_hoy())}")
    L.append("")
    if pregunta:
        L.append("## Pregunta de analisis")
        L.append(pregunta)
        L.append("")
    L.append("## Perfil")
    L.append(f"- Edad: {str(edad) + ' anios' if edad is not None else '(no calculable)'}")
    L.append(f"- Sexo: {historia.get('paciente', {}).get('A', {}).get('sexo', '(no registrado)')}")
    L.append(f"- Ocupacion: {historia.get('paciente', {}).get('A', {}).get('ocupacion', '(no registrado)')}")
    L.append(f"- Escolaridad: {historia.get('paciente', {}).get('A', {}).get('escolaridad', '(no registrado)')}")
    L.append(f"- Estado civil: {historia.get('paciente', {}).get('A', {}).get('estado_civil', '(no registrado)')}")
    L.append(f"- Religion: {historia.get('paciente', {}).get('A', {}).get('religion', '(no registrado)')}")
    L.append("")
    L.append("## Alergias (veto absoluto)")
    L.append(f"- {datos.get('alergias') or '(ninguna registrada)'}")
    L.append("")
    L.append("## Identidad clinica")
    L.append(f"- Grupo sanguineo: {datos.get('grupo_sanguineo') or '(no registrado)'}")
    fum = (epi or {}).get("fum") or ""
    fpp = calcular_fpp(fum) if fum else None
    eg = calcular_eg(fum) if fum else None
    L.append(f"- FUM: {fecha_legible(fum) if fum else '(sin FUM)'}")
    L.append(f"- FPP: {fecha_legible(fpp) if fpp else '(no calculada)'}")
    L.append(f"- EG: {eg or '(no disponible)'}")
    L.append(f"- G-P-A: {(epi or {}).get('gpa') or '(no registrado)'}")
    L.append(f"- Tipo parto anterior: {(epi or {}).get('tipo_parto_anterior') or '(no registrado)'}")
    L.append("")
    # Antecedentes desde HISTORIA.json
    L.append("## Antecedentes (secciones A-M)")
    for sid in ("D", "E1", "E2", "E3", "E4", "E5", "F1", "F2", "F3", "F4", "F5", "F6", "F7"):
        secc = next((s for s in SECCIONES_HISTORIA if s["id"] == sid), None)
        if not secc:
            continue
        vals = (historia.get("paciente") or {}).get(sid, {})
        if not vals:
            continue
        L.append(f"### {secc['titulo']}")
        for it in secc["items"]:
            if it["t"] == "auto":
                continue
            v = vals.get(it["id"], "")
            if v not in ("", None):
                L.append(f"- {it['lbl']}: {_formato_valor_historia(it, v)}")
        L.append("")
    # Serie temporal de controles
    L.append("## Serie temporal de controles prenatales")
    if ctrls:
        L.append("| Fecha | Peso | TA | FC | AU | FCF | Proteinuria |")
        L.append("|---|---|---|---|---|---|---|")
        for c in sorted(ctrls, key=lambda x: str(x.get("fecha", ""))):
            ta = ""
            if c.get("ta_s") is not None and c.get("ta_d") is not None:
                ta = f"{c['ta_s']:g}/{c['ta_d']:g}"
            def _f(v):
                return "-" if v is None else f"{v:g}"
            L.append(f"| {c.get('fecha','')} | {_f(c.get('peso'))} | {ta or '-'} | "
                     f"{_f(c.get('fc'))} | {_f(c.get('au'))} | {_f(c.get('fcf'))} | "
                     f"{c.get('proteinuria','-')} |")
    else:
        L.append("(sin controles registrados)")
    L.append("")
    # Laboratorio
    L.append("## Laboratorio (Progreso Quimico)")
    if docs.get("progreso"):
        ok_l, _, enc_l, filas_l = leer_progreso_quimico(docs["progreso"])
        if ok_l and filas_l:
            L.append("| " + " | ".join(str(h) for h in enc_l) + " |")
            L.append("|" + "---|" * len(enc_l))
            for fila in filas_l[:20]:
                L.append("| " + " | ".join(str(v) if v is not None else "" for v in fila) + " |")
    else:
        L.append("(sin archivo de progreso)")
    L.append("")
    # Alertas
    alertas = historia.get("alertas") or []
    if alertas:
        L.append("## Alertas clinicas")
        for a in alertas:
            L.append(f"- [{a.get('estado','')}] {a.get('fecha','')} {a.get('origen','')}: {a.get('descripcion','')}")
        L.append("")
    # Medicamentos
    notas = historia.get("notas") or []
    medicamentos = []
    for n in notas:
        plan = (n.get("datos") or {}).get("J", {})
        if plan.get("terapeutica"):
            medicamentos.append(str(plan["terapeutica"]))
    if medicamentos:
        L.append("## Medicamentos actuales")
        for m in medicamentos:
            L.append(f"- {m}")
        L.append("")

    texto = "\n".join(L)
    # scrubbing de texto libre (segunda capa)
    texto = _scrub_texto_libre(texto, datos)
    # puerta anti-reidentificacion
    h_ident = escanear_identificadores(texto, datos)
    if h_ident:
        return False, (f"BLOQUEADO: se detectaron posibles identificadores: "
                       f"{', '.join(h_ident)}. No se genero el archivo."), None

    EXPORTACIONES_DIR.mkdir(parents=True, exist_ok=True)
    nombre = f"HISTORIA_ANONIMA_{iid}_{fecha_hoy().replace('-', '')}.md"
    destino = EXPORTACIONES_DIR / nombre
    escribir_atomico(destino, texto)
    bitacora(f"Export anonimizado generado: {nombre} (ID={iid})", autor=autor_actual())
    if incluir_csv and ctrls:
        csv_nombre = f"HISTORIA_ANONIMA_{iid}_{fecha_hoy().replace('-', '')}.csv"
        csv_path = EXPORTACIONES_DIR / csv_nombre
        filas_csv = [["fecha", "peso", "ta_s", "ta_d", "fc", "au", "fcf", "proteinuria"]]
        for c in sorted(ctrls, key=lambda x: str(x.get("fecha", ""))):
            filas_csv.append([c.get("fecha",""), c.get("peso",""), c.get("ta_s",""),
                              c.get("ta_d",""), c.get("fc",""), c.get("au",""),
                              c.get("fcf",""), c.get("proteinuria","")])
        with open(csv_path, "w", encoding="utf-8") as f:
            for fila in filas_csv:
                f.write(",".join(str(x) for x in fila) + "\n")
    return True, f"Exportado: {destino}", str(destino)


# ===== 08g LINEA DE TIEMPO, TENDENCIAS, IMPORT Y DICTADO (ADR-0037..0040) =====

def eventos_paciente(ruta, eid=None):
    """C-22/ADR-0037: eventos de un episodio agregados de todas las fuentes.
    Funcion pura; no escribe nada. -> lista [{fecha, hora, tipo, titulo, detalle}]."""
    eventos = []
    try:
        datos = cargar_paciente(ruta)
    except Exception:
        return eventos
    epi = episodio_contexto(datos)
    eid = eid or (epi or {}).get("id", "")
    if not eid:
        return eventos
    # controles (C-16)
    for c in listar_controles(ruta, eid):
        eventos.append({"fecha": str(c.get("fecha") or ""), "hora": "",
                        "tipo": "control", "titulo": f"Control {c.get('fecha','')}",
                        "detalle": linea_control(c)})
    # historia: notas, evoluciones, egreso, alertas, calculos, consentimientos (C-18/C-20/C-19/C-11)
    try:
        h = cargar_historia(ruta)
    except Exception:
        h = {}
    for a in h.get("alertas") or []:
        if a.get("episodio") and a["episodio"] != eid:
            continue
        eventos.append({"fecha": str(a.get("fecha") or ""), "hora": str(a.get("hora") or ""),
                        "tipo": "alerta",
                        "titulo": f"Alerta {a.get('origen','')}",
                        "detalle": f"[{a.get('estado','')}] {a.get('descripcion','')}"})
    for n in h.get("notas") or []:
        if n.get("episodio") and n["episodio"] != eid:
            continue
        eventos.append({"fecha": str(n.get("fecha") or ""), "hora": str(n.get("hora") or ""),
                        "tipo": "nota", "titulo": f"Nota {n.get('tipo','')}",
                        "detalle": f"{len((n.get('datos') or {}))} secciones capturadas"})
    for c in h.get("calculos") or []:
        if c.get("episodio") and c["episodio"] != eid:
            continue
        res = c.get("resultado") or {}
        eventos.append({"fecha": str(c.get("fecha") or ""), "hora": str(c.get("hora") or ""),
                        "tipo": "calculo", "titulo": f"Cálculo {c.get('id','')}",
                        "detalle": f"valor={res.get('valor','')} {res.get('etiqueta','')}"})
    # consentimientos (C-11)
    for c in datos.get("consentimientos") or []:
        eventos.append({"fecha": str(c.get("fecha") or ""), "hora": str(c.get("hora") or ""),
                        "tipo": "consentimiento",
                        "titulo": f"Consentimiento {c.get('estado','')}",
                        "detalle": str(c.get("archivo", ""))})
    # documentos generados en el episodio
    base = Path(ruta) / eid
    if base.is_dir():
        for f in sorted(base.iterdir()):
            if f.is_file() and f.suffix in (".txt", ".xlsx", ".docx"):
                st = f.stat()
                fecha_iso = datetime.datetime.fromtimestamp(st.st_mtime).date().isoformat()
                hora_iso = datetime.datetime.fromtimestamp(st.st_mtime).time().strftime("%H:%M")
                eventos.append({"fecha": fecha_iso, "hora": hora_iso, "tipo": "documento",
                                "titulo": f"Documento: {f.name}", "detalle": ""})
    # ordenar por fecha+hora desc
    eventos.sort(key=lambda e: (e.get("fecha", ""), e.get("hora", "")), reverse=True)
    return eventos


def tendencias_controles(ruta, eid=None):
    """C-22/ADR-0038: deltas de peso/TA/FC entre primera y ultima visita. Puro."""
    try:
        datos = cargar_paciente(ruta)
    except Exception:
        return {}
    epi = episodio_contexto(datos)
    eid = eid or (epi or {}).get("id", "")
    ctrls = sorted(listar_controles(ruta, eid), key=lambda c: str(c.get("fecha", "")))
    if len(ctrls) < 2:
        return {}
    primero, ultimo = ctrls[0], ctrls[-1]
    out = {"n_visitas": len(ctrls),
           "primera_fecha": primero.get("fecha", ""), "ultima_fecha": ultimo.get("fecha", "")}
    for campo in ("peso", "ta_s", "ta_d", "fc"):
        v1, v2 = primero.get(campo), ultimo.get(campo)
        if isinstance(v1, (int, float)) and isinstance(v2, (int, float)):
            out[f"delta_{campo}"] = round(v2 - v1, 2)
    # kg/semana si hay fechas
    try:
        f1 = datetime.date.fromisoformat(str(primero.get("fecha")))
        f2 = datetime.date.fromisoformat(str(ultimo.get("fecha")))
        dias = (f2 - f1).days
        if dias > 0 and isinstance(primero.get("peso"), (int, float)) and isinstance(ultimo.get("peso"), (int, float)):
            out["peso_kg_semana"] = round((ultimo["peso"] - primero["peso"]) / dias * 7, 2)
    except Exception:
        pass
    return out


def evaluar_tendencia_alerta(ruta, eid=None):
    """ADR-0038: si un delta de TA/FCF supera el rango configurado, devuelve
    hallazgo de alerta (para registrar via C-20). Rango en _CONFIG.RANGOS_TENDENCIA
    (vacio = sin alerta de tendencia)."""
    cfg = leer_config().get("RANGOS_TENDENCIA") or {}
    if not cfg:
        return None
    tend = tendencias_controles(ruta, eid)
    avisos = []
    d_ta = tend.get("delta_ta_s")
    r_ta = cfg.get("ta_s")
    if d_ta is not None and r_ta and abs(d_ta) >= abs(r_ta):
        avisos.append(f"delta TA sistólica {d_ta:+.0f} mmHg supera rango tendencia (±{r_ta})")
    d_fcf = tend.get("delta_fcf")
    r_fcf = cfg.get("fcf")
    if d_fcf is not None and r_fcf and abs(d_fcf) >= abs(r_fcf):
        avisos.append(f"delta FCF {d_fcf:+.0f} lpm supera rango tendencia (±{r_fcf})")
    if not avisos:
        return None
    return {"nivel": "alerta", "origen": "tendencia",
            "descripcion": " · ".join(avisos)}


def recordatorios_eg(fum, hoy_iso=None):
    """ADR-0038: faltan dias para termino, trimestre, dias de termino excedidos.
    Umbrales configurables en _CONFIG.recordatorios_eg."""
    cfg = leer_config().get("recordatorios_eg") or {"semanas_termino": 37, "semanas_postermino": 42}
    out = {"trimestre": "", "dias_termino": None, "dias_postermino": None}
    if not fum:
        return out
    eg = calcular_eg(fum, hoy_iso)
    if eg:
        out["trimestre"] = trimestre_desde_eg(eg)
    fpp = calcular_fpp(fum)
    if not fpp:
        return out
    hoy = datetime.date.fromisoformat(hoy_iso) if hoy_iso else datetime.date.today()
    fpp_d = datetime.date.fromisoformat(fpp)
    dias = (fpp_d - hoy).days
    term_d = fpp_d - datetime.timedelta(weeks=(52 - cfg.get("semanas_termino", 37)))
    # dias hasta termino = dias hasta FPP - (40 - semanas_termino)*7
    dias_term = dias - (40 - cfg.get("semanas_termino", 37)) * 7
    out["dias_termino"] = dias_term
    out["dias_postermino"] = dias - (40 - cfg.get("semanas_postermino", 42)) * 7
    return out


def detectar_mapeo_csv(encabezados):
    """ADR-0039: propone mapeo de columnas CSV a campos del paciente."""
    mapa = {}
    sin = [h.strip().lower().replace(" ", "_") for h in encabezados]
    candidatos = {
        "nombre": ["nombre", "nombres", "first_name"],
        "ap_paterno": ["apellido_paterno", "ap_paterno", "paterno", "last_name"],
        "ap_materno": ["apellido_materno", "ap_materno", "materno"],
        "fecha_nacimiento": ["fecha_nacimiento", "fecha nacimiento", "nacimiento", "dob", "birth_date"],
        "curp": ["curp", "curp "],
        "telefono": ["telefono", "telefono ", "tel", "phone"],
        "correo": ["correo", "email", "mail"],
        "medico_remitente": ["medico", "medico_remitente", "doctor", "physician"],
    }
    for campo, nombres in candidatos.items():
        for i, s in enumerate(sin):
            if s in nombres:
                mapa[campo] = encabezados[i]
                break
    return mapa


def importar_pacientes_csv(ruta_csv, mapeo):
    """ADR-0039: importa pacientes desde CSV validando con los validadores existentes.
    Nunca sobrescribe expedientes existentes. -> (ok, msg, resumen)."""
    import csv as _csv
    try:
        with open(ruta_csv, encoding="utf-8-sig") as f:
            lector = _csv.DictReader(f)
            filas = list(lector)
    except Exception as e:
        return False, f"No se pudo leer el CSV: {e}", {}
    resumen = {"leidas": len(filas), "importados": 0, "duplicados": 0, "invalidos": 0, "errores": []}
    for i, fila in enumerate(filas, start=2):
        try:
            ap_p = (fila.get(mapeo.get("ap_paterno", ""), "") or "").strip()
            ap_m = (fila.get(mapeo.get("ap_materno", ""), "") or "").strip()
            nombres = (fila.get(mapeo.get("nombre", ""), "") or "").strip()
            fecha = (fila.get(mapeo.get("fecha_nacimiento", ""), "") or "").strip()
            curp = (fila.get(mapeo.get("curp", ""), "") or "").strip()
            tel = (fila.get(mapeo.get("telefono", ""), "") or "").strip()
            correo = (fila.get(mapeo.get("correo", ""), "") or "").strip()
            medico = (fila.get(mapeo.get("medico_remitente", ""), "") or "").strip()
            if not (ap_p and ap_m and nombres and fecha):
                resumen["invalidos"] += 1
                resumen["errores"].append(f"Fila {i}: faltan campos obligatorios")
                continue
            if not parsear_fecha(fecha):
                resumen["invalidos"] += 1
                resumen["errores"].append(f"Fila {i}: fecha invalida '{fecha}'")
                continue
            ok, msg, ruta_n, _d = crear_paciente(ap_p, ap_m, nombres, fecha, curp)
            if ok:
                if tel or correo or medico:
                    guardar_datos_secundarios(ruta_n, tel, correo, curp, medico, "")
                resumen["importados"] += 1
            elif "Ya existe" in msg:
                resumen["duplicados"] += 1
            else:
                resumen["invalidos"] += 1
                resumen["errores"].append(f"Fila {i}: {msg}")
        except Exception as e:
            resumen["invalidos"] += 1
            resumen["errores"].append(f"Fila {i}: {e}")
    msg = (f"Importados: {resumen['importados']} · "
           f"Duplicados omitidos: {resumen['duplicados']} · "
           f"Inválidos: {resumen['invalidos']}")
    return True, msg, resumen


def dictado_disponible():
    """ADR-0040: detecta motor de dictado local (Vosk). -> (ok, msg)."""
    try:
        import vosk  # noqa: F401
        return True, "vosk disponible"
    except ImportError:
        pass
    return False, ("dictado del sistema operativo disponible "
                   "(Windows: Win+H · macOS: fn fn · Linux: Super+Ctrl+Espacio)")


# ===== 09 SERVICIO EXPORTACION =====

def _filas_pacientes():
    filas = [["Nombre completo", "Fecha nacimiento", "CURP", "Telefono", "Correo",
              "Medico remitente", "Episodios", "FUM", "EG", "Notas", "Expediente"]]
    for nombre, ruta in listar_pacientes():
        try:
            d = cargar_paciente(ruta)
        except Exception:
            d = {}
        epi = episodio_contexto(d)
        fum = (epi or {}).get("fum") or ""
        eg = calcular_eg(fum) if fum else None
        filas.append([
            d.get("nombre_completo", nombre),
            fecha_legible(d.get("fecha_nacimiento", "")),
            d.get("curp", ""), d.get("telefono", ""), d.get("correo", ""),
            d.get("medico_remitente", ""),
            str(len(d.get("episodios") or [])),
            fecha_legible(fum) if fum else "",
            eg or "",
            d.get("notas", ""),
            str(ruta),
        ])
    return filas


def exportar_pacientes_xlsx(destino):
    try:
        filas = _filas_pacientes()
        _crear_xlsx_minimo(Path(destino), filas)
        bitacora(f"Exportacion Excel: {destino}", autor=autor_actual())
        return True, f"Exportado: {len(filas) - 1} pacientes.", str(destino)
    except Exception as e:
        log_exc(e)
        return False, f"Error: {e}", None


def exportar_pacientes_csv(destino):
    try:
        import csv as _csv
        with open(destino, "w", newline="", encoding="utf-8-sig") as f:
            w = _csv.writer(f)
            for fila in _filas_pacientes():
                w.writerow(fila)
        bitacora(f"Exportacion CSV: {destino}", autor=autor_actual())
        return True, f"Exportado: {Path(destino).name}", str(destino)
    except Exception as e:
        log_exc(e)
        return False, f"Error: {e}", None


def _pdf_minimo(destino, lineas, titulo="Exportacion"):
    def esc(t):
        return t.replace("\\", r"\\").replace("(", r"\(").replace(")", r"\)")
    paginas = [lineas[i:i + 48] for i in range(0, len(lineas), 48)] or [[""]]
    objs = {}
    n_pag = len(paginas)
    kids = " ".join(f"{5 + 2 * i} 0 R" for i in range(n_pag))
    objs[1] = "<< /Type /Catalog /Pages 2 0 R >>"
    objs[2] = f"<< /Type /Pages /Kids [{kids}] /Count {n_pag} >>"
    objs[3] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
    objs[4] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"
    for i, pag in enumerate(paginas):
        num_pag = 5 + 2 * i
        num_cont = 6 + 2 * i
        y = 760
        partes = [f"BT /F2 11 Tf 50 {y} Td ({esc(titulo)} - Pagina {i + 1}/{n_pag}) Tj ET"]
        y -= 24
        partes.append("BT /F1 9 Tf")
        for ln in pag:
            try:
                txt = ln.encode("cp1252", "replace").decode("cp1252")
            except Exception:
                txt = ln
            partes.append(f"1 0 0 1 50 {y} Tm ({esc(txt)}) Tj")
            y -= 14
        partes.append("ET")
        stream = "\n".join(partes)
        objs[num_pag] = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
            f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents {num_cont} 0 R >>"
        )
        objs[num_cont] = {"stream": stream}
    salida = bytearray(b"%PDF-1.4\n")
    offsets = {}
    for num in sorted(objs):
        offsets[num] = len(salida)
        contenido = objs[num]
        if isinstance(contenido, dict):
            s = contenido["stream"].encode("latin-1", "replace")
            bloque = (f"{num} 0 obj\n<< /Length {len(s)} >>\nstream\n").encode() + s + b"\nendstream\nendobj\n"
        else:
            bloque = f"{num} 0 obj\n{contenido}\nendobj\n".encode("latin-1", "replace")
        salida.extend(bloque)
    xref_pos = len(salida)
    max_obj = max(objs)
    salida.extend(f"xref\n0 {max_obj + 1}\n".encode())
    salida.extend(b"0000000000 65535 f \n")
    for num in range(1, max_obj + 1):
        salida.extend(f"{offsets[num]:010d} 00000 n \n".encode())
    salida.extend(
        f"trailer\n<< /Size {max_obj + 1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode()
    )
    with open(destino, "wb") as f:
        f.write(salida)


def exportar_pacientes_pdf(destino):
    try:
        filas = _filas_pacientes()
        lineas = ["REPODE DE PACIENTES - MATERNIDAD", f"Fecha: {fecha_legible(fecha_hoy())}", ""]
        for fila in filas[1:]:
            lineas.append(f"{fila[0]} | Nac: {fila[1]} | CURP: {fila[2]} | Tel: {fila[3]}")
            lineas.append(f"    Medico: {fila[5]} | Episodios: {fila[6]}")
            lineas.append("")
        if not lineas[-1]:
            lineas = lineas[:-1]
        lineas.append("")
        lineas.append(f"Total: {len(filas) - 1} pacientes. Soporte normativo, no certificacion.")
        _pdf_minimo(Path(destino), lineas, "Listado de pacientes")
        bitacora(f"Exportacion PDF: {destino}", autor=autor_actual())
        return True, f"Exportado: {len(filas) - 1} pacientes.", str(destino)
    except Exception as e:
        log_exc(e)
        return False, f"Error: {e}", None


def exportar_expediente(ruta_paciente, destino):
    try:
        ruta_paciente = Path(ruta_paciente)
        with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
            for p in sorted(ruta_paciente.rglob("*")):
                if p.is_file():
                    z.write(p, p.relative_to(ruta_paciente.parent))
        bitacora(f"Expediente exportado: {ruta_paciente.name} -> {destino}",
                 autor=autor_actual())
        return True, f"Expediente exportado: {Path(destino).name}", str(destino)
    except Exception as e:
        log_exc(e)
        return False, f"Error: {e}", None


# ===== 10 INTERFAZ =====

if TK_DISPONIBLE:

    class CalendarioPopup(tk.Toplevel):
        """Selector de fecha: anio/mes/dia navegables, mes por nombre, teclado o clic."""

        def __init__(self, parent, al_elegir, formato="iso", iso_inicial=None):
            super().__init__(parent)
            self.al_elegir = al_elegir
            self.formato = formato
            self.title("Seleccionar fecha")
            self.resizable(False, False)
            self.configure(bg=CREMA)
            self.transient(parent)
            try:
                base = (datetime.date.fromisoformat(iso_inicial)
                        if iso_inicial else None)
            except Exception:
                base = None
            if base is None:
                base = datetime.date(2000, 1, 1)
            self.anio = base.year
            self.mes = base.month
            self.dia = base.day
            self.anio_max = datetime.date.today().year
            self.anio_min = 1900

            f_anio = tk.Frame(self, bg=CREMA)
            f_anio.pack(pady=(12, 2))
            tk.Button(f_anio, text="‹", width=3, command=lambda: self._mover(anio=-1),
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER
                      ).pack(side="left", padx=3)
            self.anio_var = tk.StringVar(value=str(self.anio))
            self.anio_cb = ttk.Combobox(f_anio, textvariable=self.anio_var, width=7,
                                        values=[str(a) for a in
                                                range(self.anio_min, self.anio_max + 1)],
                                        font=(FUENTE_SISTEMA, 13, "bold"))
            self.anio_cb.pack(side="left", padx=4)
            self.anio_cb.bind("<Return>", self._enter_anio)
            self.anio_cb.bind("<<ComboboxSelected>>", self._sel_anio)
            tk.Button(f_anio, text="›", width=3, command=lambda: self._mover(anio=1),
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER
                      ).pack(side="left", padx=3)

            f_mes = tk.Frame(self, bg=CREMA)
            f_mes.pack(pady=2)
            tk.Button(f_mes, text="‹", width=3, command=lambda: self._mover(mes=-1),
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER
                      ).pack(side="left", padx=3)
            self.mes_var = tk.StringVar(value=MESES_LARGOS[self.mes - 1])
            self.mes_cb = ttk.Combobox(f_mes, textvariable=self.mes_var, width=14,
                                       values=MESES_LARGOS, font=(FUENTE_SISTEMA, 13, "bold"))
            self.mes_cb.pack(side="left", padx=4)
            self.mes_cb.bind("<Return>", self._enter_mes)
            self.mes_cb.bind("<<ComboboxSelected>>", self._sel_mes)
            tk.Button(f_mes, text="›", width=3, command=lambda: self._mover(mes=1),
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER
                      ).pack(side="left", padx=3)

            f_dia = tk.Frame(self, bg=CREMA)
            f_dia.pack(pady=(6, 4))
            tk.Label(f_dia, text="Dia:", bg=CREMA, fg=TEXTO_OSCURO,
                     font=(FUENTE_SISTEMA, 11)).pack(side="left")
            self.dia_var = tk.StringVar(value=str(self.dia))
            self.dia_spin = tk.Spinbox(f_dia, from_=1, to=31, width=4, textvariable=self.dia_var,
                                       font=(FUENTE_SISTEMA, 13, "bold"), command=self._spin_dia)
            self.dia_spin.pack(side="left", padx=6)
            self.dia_spin.bind("<Return>", self._enter_dia)
            tk.Label(f_dia, text="(Enter selecciona y cierra)", bg=CREMA, fg=GRIS_TERCIARIO,
                     font=(FUENTE_SISTEMA, 8)).pack(side="left")

            f_semana = tk.Frame(self, bg=AZUL)
            f_semana.pack(fill="x", padx=8, pady=(4, 0))
            for sigla in ("L", "M", "M", "J", "V", "S", "D"):
                tk.Label(f_semana, text=sigla, width=5, bg=AZUL, fg=TEXTO_CLARO,
                         font=(FUENTE_SISTEMA, 10, "bold")).pack(side="left")

            self.frame_dias = tk.Frame(self, bg=CREMA)
            self.frame_dias.pack(fill="both", expand=True, padx=8, pady=4)

            f_pie = tk.Frame(self, bg=CREMA)
            f_pie.pack(fill="x", padx=8, pady=(0, 10))
            tk.Button(f_pie, text="Hoy", command=self._hoy, width=8,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER
                      ).pack(side="left", padx=4)
            self.btn_sel = tk.Button(f_pie, text="Seleccionar fecha",
                                     command=self._elegir_seleccion, width=16,
                                     bg=AZUL, fg=TEXTO_CLARO,
                                     activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 10, "bold"))
            self.btn_sel.pack(side="left", padx=4)
            tk.Button(f_pie, text="Cancelar", command=self.destroy, width=10,
                      bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO, activebackground=GRIS_APAGADO
                      ).pack(side="right", padx=4)

            self._pintar()
            self.anio_cb.focus_set()
            self.anio_cb.select_range(0, "end")
            self.update_idletasks()
            x = parent.winfo_rootx() + max(0, (parent.winfo_width() - self.winfo_width()) // 2)
            y = parent.winfo_rooty() + max(0, (parent.winfo_height() - self.winfo_height()) // 3)
            self.geometry(f"+{x}+{y}")
            self.grab_set()

        def _enter_anio(self, event=None):
            try:
                a = int(str(self.anio_var.get()).strip())
            except ValueError:
                a = self.anio
            self.anio = max(self.anio_min, min(a, self.anio_max))
            self._clamp_dia()
            self._pintar()
            return "break"

        def _sel_anio(self, event=None):
            # elegir el anio del desplegable actualiza la vista (auditoria B1)
            return self._enter_anio(event)

        def _enter_mes(self, event=None):
            q = normalizar(self.mes_var.get())
            if len(q) >= 3:
                for i, nombre in enumerate(MESES_LARGOS):
                    if normalizar(nombre).startswith(q):
                        self.mes = i + 1
                        break
            else:
                self.mes = max(1, min(self.mes, 12))
            self._clamp_dia()
            self._pintar()
            return "break"

        def _sel_mes(self, event=None):
            try:
                self.mes = MESES_LARGOS.index(self.mes_var.get()) + 1
            except ValueError:
                pass
            self._clamp_dia()
            self._pintar()

        def _spin_dia(self):
            try:
                self.dia = max(1, int(self.dia_var.get()))
            except ValueError:
                pass

        def _es_mes_actual(self):
            hoy = datetime.date.today()
            return (self.anio, self.mes) == (hoy.year, hoy.month)

        def _es_mes_futuro(self):
            hoy = datetime.date.today()
            return (self.anio, self.mes) > (hoy.year, hoy.month)

        def _max_dia(self):
            # dias permitidos: en el mes en curso, hasta hoy; si no, todo el mes
            tope = dias_en_mes(self.anio, self.mes)
            if self._es_mes_actual():
                return min(tope, datetime.date.today().day)
            return tope

        def _enter_dia(self, event=None):
            if self._es_mes_futuro():
                # dia tecnico imposible: no se puede seleccionar un mes futuro
                self._pintar()
                return "break"
            try:
                d = int(str(self.dia_var.get()).strip())
            except ValueError:
                d = self.dia
            d = max(1, min(d, self._max_dia()))
            try:
                elegida = datetime.date(self.anio, self.mes, d)
            except ValueError:
                return "break"
            self._elegir(elegida)
            return "break"

        def _elegir_seleccion(self):
            # R3: lee el spin (no self.dia) para respetar el dia tecleado
            try:
                d = int(str(self.dia_var.get()).strip())
            except ValueError:
                d = self.dia
            if self._es_mes_futuro():
                return
            d = max(1, min(d, self._max_dia()))
            try:
                self._elegir(datetime.date(self.anio, self.mes, d))
            except ValueError:
                pass

        def _mover(self, anio=0, mes=0):
            if mes:
                m, a = self.mes + mes, self.anio
                if m < 1:
                    m, a = 12, a - 1
                elif m > 12:
                    m, a = 1, a + 1
                self.mes, self.anio = m, max(self.anio_min, min(a, self.anio_max))
            if anio:
                self.anio = max(self.anio_min, min(self.anio + anio, self.anio_max))
            self._clamp_dia()
            self._pintar()

        def _clamp_dia(self):
            # D18: conserva el dia si existe; si no acota al ultimo valido del mes
            self.dia = min(max(1, self.dia), self._max_dia())

        def _pintar(self):
            self.anio_var.set(str(self.anio))
            self.mes_var.set(MESES_LARGOS[self.mes - 1])
            tope = dias_en_mes(self.anio, self.mes)
            self._clamp_dia()
            self.dia_spin.config(to=tope)
            self.dia_var.set(str(self.dia))
            futuro = self._es_mes_futuro()
            maximo = self._max_dia()
            for w in self.frame_dias.winfo_children():
                w.destroy()
            offset = calendar.monthrange(self.anio, self.mes)[0]  # 0 = lunes
            for d in range(1, tope + 1):
                fila, col = divmod(offset + d - 1, 7)
                permitido = (not futuro) and d <= maximo
                elegido = permitido and d == self.dia
                boton = tk.Button(self.frame_dias, text=str(d), width=5,
                                  bg=AZUL if elegido else BLANCO,
                                  fg=TEXTO_CLARO if elegido else TEXTO_OSCURO,
                                  activebackground=AZUL_HOVER, relief="flat",
                                  font=(FUENTE_SISTEMA, 10),
                                  command=lambda dd=d: self._elegir(
                                      datetime.date(self.anio, self.mes, dd)))
                if not permitido:
                    boton.config(state="disabled",
                                 disabledforeground="#999999", bg=GRIS_REJILLA)
                boton.grid(row=fila, column=col, padx=1, pady=1)
            # R5: Select inactivo si todo el mes es futuro
            if hasattr(self, "btn_sel"):
                self.btn_sel.config(state="disabled" if futuro else "normal")

        def _elegir(self, fecha_sel):
            iso = fecha_sel.isoformat()
            self.al_elegir(fecha_legible(iso) if self.formato == "legible" else iso)
            self.destroy()

        def _hoy(self):
            self._elegir(datetime.date.today())

    def boton_calendario(master, var, formato, popup_parent=None, on_change=None):
        """Boton [Cal] estandar: abre CalendarioPopup sobre el campo (F1)."""
        def _abrir():
            iso = parsear_fecha(var.get())

            def _escribir(valor):
                var.set(valor)
                if on_change is not None:
                    on_change()

            CalendarioPopup(popup_parent or master, _escribir,
                            formato=formato, iso_inicial=iso)

        return tk.Button(master, text="[Cal]", command=_abrir, bg=AZUL,
                         fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                         font=(FUENTE_SISTEMA, 9))

    def _label_calculo(parent, var_fum, texto_vacio="(FUM opcional: muestra FPP y edad gestacional)",
                       wrap=440):
        """Label FPP/EG en vivo compartido por los dialogos de embarazo (F1)."""
        lbl = tk.Label(parent, text="", bg=CREMA_CLARO, fg=AZUL,
                       font=(FUENTE_SISTEMA, 10, "bold"), wraplength=wrap, justify="left")

        def actualizar(*_args):
            fum_v = parsear_fecha(var_fum.get())
            if not fum_v:
                lbl.config(text=texto_vacio)
                return
            fpp = calcular_fpp(fum_v)
            eg = calcular_eg(fum_v)
            txt = f"FPP estimada: {fecha_legible(fpp)}" if fpp else "FPP: -"
            txt += f"  |  Edad gestacional hoy: {eg}" if eg else "  |  EG: - (FUM futura o sin calcular)"
            lbl.config(text=txt)

        return lbl, actualizar

    # --- Contexto de copiar/pegar global (seleccion + boton derecho) ---
    def _texto_ventana(win):
        """Recopila todo el texto visible de una ventana (para 'Copiar todo')."""
        partes = []
        for w in win.winfo_children():
            try:
                cls = w.winfo_class()
                if cls in ("Label", "TLabel", "Button", "TButton",
                           "Checkbutton", "Radiobutton"):
                    txt = str(w.cget("text"))
                    if txt and txt.strip():
                        partes.append(txt.strip())
                elif cls in ("Entry", "TEntry"):
                    v = str(w.get())
                    if v.strip():
                        partes.append(v.strip())
                elif cls == "Text":
                    v = w.get("1.0", "end-1c").strip()
                    if v:
                        partes.append(v)
                elif cls == "TCombobox":
                    v = str(w.get())
                    if v.strip():
                        partes.append(v.strip())
            except Exception:
                pass
            partes.extend(_texto_ventana(w))
        return partes

    def _texto_de_ventana(widget):
        """Recopila todo el texto visible de una ventana (Entry/Text/Label/Listbox)."""
        partes = []
        try:
            hijos = widget.winfo_children()
        except Exception:
            hijos = []
        for w in hijos:
            try:
                cls = w.winfo_class()
                if cls in ("Entry", "TEntry"):
                    partes.append(str(w.get()))
                elif cls == "Text":
                    partes.append(w.get("1.0", "end-1c"))
                elif cls in ("Label", "TLabel"):
                    try:
                        _t = str(w.cget("text"))
                    except Exception:
                        _t = ""
                    if _t and _t.strip():
                        partes.append(_t)
                elif cls == "Listbox":
                    partes.extend(str(w.get(i)) for i in range(w.size()))
            except Exception:
                pass
            partes.append(_texto_de_ventana(w))
        return "\n".join(p for p in partes if p and str(p).strip())

    def _menu_contexto(win, w):
        """Menu contextual UNIVERSAL: funciona en cualquier widget, no solo
        casilleros. Siempre devuelve un menu con al menos 'Copiar todo'."""
        menu = tk.Menu(win, tearoff=0)
        cls = w.winfo_class() if hasattr(w, "winfo_class") else ""

        def _copiar_todo_ventana():
            try:
                win.clipboard_clear()
                win.clipboard_append(_texto_de_ventana(win))
            except Exception:
                pass

        def _copiar_todo_widget():
            try:
                txt = _texto_de_ventana(w)
                win.clipboard_clear()
                win.clipboard_append(txt)
            except Exception:
                pass

        if cls in ("Entry", "TEntry"):
            menu.add_command(label="Cortar", command=lambda: w.event_generate("<Control-x>"))
            menu.add_command(label="Copiar", command=lambda: w.event_generate("<Control-c>"))
            menu.add_command(label="Pegar", command=lambda: w.event_generate("<Control-v>"))
            menu.add_separator()
            menu.add_command(label="Seleccionar todo",
                             command=lambda: (w.focus_set(), w.select_range(0, "end")))
        elif cls == "Text":
            try:
                hay_sel = bool(w.tag_ranges("sel"))
            except Exception:
                hay_sel = False
            def _copiar_sel():
                try:
                    sel = w.get("sel.first", "sel.last")
                except tk.TclError:
                    sel = None
                if sel:
                    win.clipboard_clear()
                    win.clipboard_append(sel)
            if hay_sel:
                menu.add_command(label="Copiar seleccion", command=_copiar_sel)
            menu.add_command(label="Copiar contenido",
                             command=lambda: (win.clipboard_clear(),
                                              win.clipboard_append(w.get("1.0", "end-1c"))))
            if str(w.cget("state")) != "disabled":
                menu.add_command(label="Pegar",
                                 command=lambda: w.event_generate("<Control-v>"))
                menu.add_command(label="Seleccionar todo",
                                 command=lambda: w.tag_add("sel", "1.0", "end"))
        elif cls in ("Label", "TLabel", "Button", "TButton",
                     "Checkbutton", "Radiobutton"):
            try:
                txt = str(w.cget("text"))
            except Exception:
                txt = ""
            if txt.strip():
                menu.add_command(label="Copiar texto",
                                 command=lambda: (win.clipboard_clear(),
                                                  win.clipboard_append(txt)))
        elif cls == "Listbox":
            sel = w.curselection()
            if sel:
                vals = "\n".join(w.get(i) for i in sel)
                menu.add_command(label="Copiar seleccion",
                                 command=lambda: (win.clipboard_clear(),
                                                  win.clipboard_append(vals)))
        elif cls == "Treeview":
            sel = w.selection()
            if sel:
                vals = w.item(sel[0], "values")
                txt = "\t".join(str(v) for v in vals)
                menu.add_command(label="Copiar fila",
                                 command=lambda: (win.clipboard_clear(),
                                                  win.clipboard_append(txt)))
        # SIEMPRE: copiar todo el texto de la ventana (fallback universal)
        menu.add_separator()
        menu.add_command(label="Copiar todo el texto de la ventana",
                         command=_copiar_todo_ventana)
        return menu

    def instalar_contexto(win):
        """Boton derecho UNIVERSAL + Ctrl+A. El binding en el toplevel captura
        a todos los descendientes (event.widget apunta al widget real)."""
        def _handler(event):
            try:
                menu = _menu_contexto(win, event.widget)
            except Exception:
                menu = None
            if menu:
                try:
                    menu.tk_popup(event.x_root, event.y_root)
                finally:
                    menu.grab_release()
        win.bind("<Button-3>", _handler, add="+")

        def _ctrl_a(event):
            w = event.widget
            cls = w.winfo_class()
            try:
                if cls in ("Entry", "TEntry"):
                    w.select_range(0, "end")
                    return "break"
                if cls == "Text" and str(w.cget("state")) != "disabled":
                    w.tag_add("sel", "1.0", "end")
                    return "break"
            except Exception:
                pass
        win.bind("<Control-a>", _ctrl_a, add="+")

        # v4.12: resaltado de FOCO en campos de datos (ensombrecimiento al escribir)
        def _foco_in(event):
            w = event.widget
            try:
                cls = w.winfo_class()
                if cls in ("Entry", "TEntry", "Text", "TCombobox"):
                    if not hasattr(w, "_bg_orig"):
                        w._bg_orig = w.cget("bg")
                    w.configure(bg="#E8F0FD")
            except Exception:
                pass

        def _foco_out(event):
            w = event.widget
            try:
                if hasattr(w, "_bg_orig"):
                    w.configure(bg=w._bg_orig)
            except Exception:
                pass
        win.bind("<FocusIn>", _foco_in, add="+")
        win.bind("<FocusOut>", _foco_out, add="+")

    def envoltorio_scroll(parent, bg=None):
        """Devuelve (interno, cont) con scrollbar vertical. El caller debe
        grid/pack 'cont' donde quiera la pagina; 'interno' va dentro del canvas.
        Para paginas/dialogos con muchos campos (ADR-0026/v4.16)."""
        bg = bg or CREMA_CLARO
        cont = tk.Frame(parent, bg=bg)
        cont.grid_rowconfigure(0, weight=1)
        cont.grid_columnconfigure(0, weight=1)
        canvas = tk.Canvas(cont, bg=bg, highlightthickness=0)
        scr = ttk.Scrollbar(cont, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scr.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scr.grid(row=0, column=1, sticky="ns")
        interno = tk.Frame(canvas, bg=bg)
        interno.grid_columnconfigure(0, weight=1)
        vid = canvas.create_window((0, 0), window=interno, anchor="nw")
        interno.bind("<Configure>",
                     lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(vid, width=e.width))
        return interno, cont

    def auto_tamano(win, parent=None):
        """Ventana a la medida de su contenido + centrada + modal (L3/D4)."""
        try:
            win.update_idletasks()
            ancho = max(win.winfo_reqwidth(), 240) + 8
            alto = max(win.winfo_reqheight(), 120) + 8
            x = y = None
            if parent is not None:
                win.transient(parent)
                if parent.winfo_exists():
                    x = parent.winfo_rootx() + max(0, (parent.winfo_width() - ancho) // 2)
                    y = parent.winfo_rooty() + max(0, (parent.winfo_height() - alto) // 3)
            geo = f"{ancho}x{alto}"
            if x is not None:
                geo += f"+{x}+{y}"
            win.geometry(geo)
            win.grab_set()
            instalar_contexto(win)
        except Exception:
            pass

    class Tooltip:
        """Aidita emergente tras 600 ms (ADR-0027). Sin modal ni grab."""

        def __init__(self, widget, texto):
            self.w = widget
            self.txt = texto
            self.top = None
            self._id = None
            widget.bind("<Enter>", self._programar, add="+")
            widget.bind("<Leave>", self._ocultar, add="+")
            widget.bind("<Destroy>", self._destruir, add="+")

        def _programar(self, _e=None):
            try:
                self._id = self.w.after(600, self._mostrar)
            except Exception:
                self._id = None

        def _mostrar(self):
            self._id = None
            try:
                if not self.w.winfo_exists():
                    return
                tw = tk.Toplevel(self.w)
                tw.wm_overrideredirect(True)
                x = self.w.winfo_rootx() + 6
                y = self.w.winfo_rooty() + self.w.winfo_height() + 3
                tw.wm_geometry(f"+{x}+{y}")
                tk.Label(tw, text=self.txt, bg="#FFFFE1", fg=TEXTO_OSCURO,
                         relief="solid", bd=1, font=(FUENTE_SISTEMA, 9),
                         padx=6, pady=3).pack()
                self.top = tw
            except Exception:
                self.top = None

        def _ocultar(self, _e=None):
            if self._id is not None:
                try:
                    self.w.after_cancel(self._id)
                except Exception:
                    pass
                self._id = None
            if self.top is not None:
                try:
                    self.top.destroy()
                except Exception:
                    pass
                self.top = None

        def _destruir(self, e):
            if getattr(e, "widget", None) is self.w:
                self._ocultar()

    class DocGenerator(tk.Toplevel):
        def __init__(self, parent_app, ruta):
            super().__init__(parent_app)
            self.app = parent_app
            self.ruta = ruta
            self.title("Generar documento")
            self.resizable(False, False)
            marco = tk.Frame(self, bg=CREMA_CLARO, padx=16, pady=16)
            marco.pack(fill="both", expand=True, padx=12, pady=12)
            tk.Label(marco, text="Seleccionar plantilla", font=(FUENTE_SISTEMA, 14, "bold"),
                     bg=CREMA_CLARO, fg=TEXTO_OSCURO).pack(anchor="w")
            self.tpls = listar_plantillas()
            self.tpl_var = tk.StringVar(value=self.tpls[0] if self.tpls else "")
            cb = ttk.Combobox(marco, textvariable=self.tpl_var, values=self.tpls,
                              state="readonly", font=(FUENTE_SISTEMA, 11))
            cb.pack(fill="x", pady=8)
            tk.Label(marco, text=f"Paciente: {nombre_mostrar(ruta)}",
                     bg=CREMA_CLARO, fg=TEXTO_OSCURO, wraplength=440,
                     justify="left").pack(anchor="w", pady=4)
            pie = tk.Frame(marco, bg=CREMA_CLARO)
            pie.pack(fill="x", pady=14)
            tk.Button(pie, text="Generar", command=self._generar, width=12,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER).pack(side="right", padx=4)
            tk.Button(pie, text="Cancelar", command=self.destroy, width=12,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER).pack(side="right")
            auto_tamano(self, parent_app)

        def _generar(self):
            tpl = self.tpl_var.get()
            if not tpl:
                messagebox.showwarning("Aviso", "Selecciona una plantilla.")
                return
            self.app.generar_con(self.ruta, tpl)
            self.destroy()

    class FichaPaciente(tk.Toplevel):
        """Ficha completa: ver/editar datos del paciente + episodios con FUM/FPP/EG (G1/G2)."""

        def __init__(self, parent_app, ruta):
            super().__init__(parent_app)
            self.app = parent_app
            self.ruta = ruta
            self.title("Ficha de paciente")
            self.configure(bg=CREMA)
            self.resizable(False, False)
            self.transient(parent_app)
            try:
                datos = cargar_paciente(ruta)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el expediente:\n{e}")
                self.destroy()
                return
            self._datos = datos
            marco = tk.Frame(self, bg=CREMA_CLARO, padx=16, pady=14)
            marco.pack(fill="both", expand=True, padx=10, pady=10)

            tk.Label(marco, text=datos.get("nombre_completo", ""),
                     font=(FUENTE_SISTEMA, 15, "bold"), bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).pack(anchor="w")
            tk.Label(marco,
                     text=f"Nacimiento: {fecha_legible(datos.get('fecha_nacimiento', ''))}"
                          f"   |   Expediente: {datos.get('clave', '')}",
                     bg=CREMA_CLARO, fg=GRIS_TEXTO).pack(anchor="w")
            tk.Label(marco, text="* nombre y fecha de nacimiento no se editan aqui",
                     bg=CREMA_CLARO, fg=GRIS_APAGADO,
                     font=(FUENTE_SISTEMA, 9)).pack(anchor="w", pady=(0, 6))

            tk.Label(marco, text="· DATOS DE LA PACIENTE ·", font=(FUENTE_SISTEMA, 9, "bold"),
                     bg=CREMA_CLARO, fg=AZUL).pack(anchor="w", pady=(4, 2))
            campos = tk.Frame(marco, bg=CREMA_CLARO)
            campos.pack(fill="x")
            campos.grid_columnconfigure(1, weight=1)
            self.vars = {}
            for i, (etiq, clave) in enumerate([("CURP", "curp"),
                                               ("Telefono", "telefono"),
                                               ("Correo", "correo"),
                                               ("Medico remitente", "medico_remitente"),
                                               ("Notas", "notas"),
                                               ("Alergias (Tier A)", "alergias"),
                                               ("Grupo sanguineo (Tier A)", "grupo_sanguineo")]):
                tk.Label(campos, text=etiq, bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                         font=(FUENTE_SISTEMA, 10)).grid(row=i, column=0, sticky="w",
                                                     pady=2, padx=(0, 8))
                var = tk.StringVar(value=datos.get(clave, "") or "")
                e = tk.Entry(campos, textvariable=var, font=(FUENTE_SISTEMA, 11),
                             bg=BLANCO, state="disabled", relief="solid", bd=1)
                e.grid(row=i, column=1, sticky="we", pady=2)
                self.vars[clave] = var

            tk.Label(marco, text="· EPISODIOS DE EMBARAZO Y CONTROLES ·",
                     font=(FUENTE_SISTEMA, 9, "bold"), bg=CREMA_CLARO,
                     fg=AZUL).pack(anchor="w", pady=(10, 2))
            marco_tabla = tk.Frame(marco, bg=BLANCO, highlightthickness=1,
                                   highlightbackground=GRIS_BORDE)
            marco_tabla.pack(fill="both", expand=True)
            self.tabla = ttk.Treeview(
                marco_tabla, columns=("epi", "ref", "fum", "fpp", "eg", "ctrls"),
                show="headings", height=min(6, max(2, len(datos.get("episodios", [])))))
            for col, texto, ancho in (("epi", "Episodio", 90), ("ref", "Fecha ref.", 90),
                                      ("fum", "FUM", 90), ("fpp", "FPP", 90),
                                      ("eg", "EG", 75), ("ctrls", "Ctrls", 45)):
                self.tabla.heading(col, text=texto)
                self.tabla.column(col, width=ancho, anchor="center")
            self.tabla.pack(fill="both", expand=True)
            for epi in datos.get("episodios", []):
                fum = epi.get("fum") or ""
                fpp = calcular_fpp(fum) if fum else None
                eg = calcular_eg(fum) if fum else None
                self.tabla.insert("", "end", values=(
                    epi.get("id", ""),
                    fecha_legible(epi.get("fecha_referencia", "")) if epi.get("fecha_referencia") else "-",
                    fecha_legible(fum) if fum else "-",
                    fecha_legible(fpp) if fpp else "-",
                    eg or "-",
                    str(len(epi.get("controles") or []))))

            pie = tk.Frame(marco, bg=CREMA_CLARO)
            pie.pack(fill="x", pady=(12, 0))
            self.btn_editar = tk.Button(pie, text="Editar datos", command=self._editar,
                                        bg=AZUL, fg=TEXTO_CLARO,
                                        activebackground=AZUL_HOVER, padx=8, pady=5)
            self.btn_editar.pack(side="left", padx=3)
            self.btn_guardar = tk.Button(pie, text="Guardar", command=self._guardar,
                                         bg=VERDE_OK, fg="white", state="disabled",
                                         padx=8, pady=5)
            self.btn_guardar.pack(side="left", padx=3)
            tk.Button(pie, text="Abrir carpeta", command=lambda: abrir_carpeta(self.ruta),
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      padx=8, pady=5).pack(side="left", padx=3)
            tk.Button(pie, text="Ver historial", command=self._historial,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      padx=8, pady=5).pack(side="left", padx=3)
            tk.Button(pie, text="Export anonimo",
                      command=lambda: self.app._exportar_anonimo(),
                      bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                      activebackground=GRIS_APAGADO, padx=8, pady=5
                      ).pack(side="left", padx=3)
            # C-22: timeline en 1 clic desde la Ficha
            tk.Button(pie, text="Timeline",
                      command=lambda: self.app._abrir_timeline(),
                      bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                      activebackground=GRIS_APAGADO, padx=8, pady=5
                      ).pack(side="left", padx=3)
            tk.Button(pie, text="Cerrar", command=self.destroy, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO,
                      padx=8, pady=5).pack(side="right", padx=3)
            auto_tamano(self, parent_app)

        def _editar(self):
            for hijo in self._walk_entries(self):
                hijo.config(state="normal")
            self.btn_guardar.config(state="normal")
            self.btn_editar.config(state="disabled")

        def _walk_entries(self, widget):
            encontrados = []
            for hijo in widget.winfo_children():
                if isinstance(hijo, tk.Entry):
                    encontrados.append(hijo)
                encontrados.extend(self._walk_entries(hijo))
            return encontrados

        def _historial(self):
            ruta_hist = Path(self.ruta) / "HISTORIAL.txt"
            if ruta_hist.exists():
                abrir_archivo(ruta_hist)
            else:
                messagebox.showinfo("Aviso", "Esta paciente aun no tiene historial.")

        def _guardar(self):
            curp = self.vars["curp"].get().strip()
            if curp:
                ok, msg = validar_curp(curp)
                if not ok and not messagebox.askyesno(
                        "CURP invalida", f"{msg}\n¿Deseas guardar sin CURP valida?",
                        parent=self):
                    return
            ok, msg = guardar_datos_secundarios(
                self.ruta,
                self.vars["telefono"].get(),
                self.vars["correo"].get(),
                curp,
                self.vars["medico_remitente"].get(),
                self.vars["notas"].get(),
                alergias=self.vars["alergias"].get(),
                grupo_sanguineo=self.vars["grupo_sanguineo"].get())
            if ok:
                self.app.datos_activos = cargar_paciente(self.ruta)
                self.app.refresh_home()
                self.app.refrescar_recientes()
                self.app._actualizar_etiqueta()
                self.app._actualizar_episodios()
                self.app._actualizar_admin_sel()
                self.app.status.set("Ficha actualizada")
                self.app.toast("Datos actualizados ✓")
                for hijo in self._walk_entries(self):
                    hijo.config(state="disabled")
                self.btn_guardar.config(state="disabled")
                self.btn_editar.config(state="normal")
            else:
                messagebox.showerror("Error", msg, parent=self)

    class ControlesPrenatales(tk.Toplevel):
        """C-16 (v2.1): controles prenatales del episodio — lista, alta y edicion.
        Sin borrado (riesgo clínico, C-16 §5)."""

        def __init__(self, parent_app):
            super().__init__(parent_app)
            self.app = parent_app
            self.title("Controles prenatales")
            self.configure(bg=CREMA)
            marco = tk.Frame(self, bg=CREMA_CLARO, padx=14, pady=12)
            marco.pack(fill="both", expand=True, padx=10, pady=10)
            self.lbl_ctx = tk.Label(marco, text="", font=(FUENTE_SISTEMA, 11, "bold"),
                                    bg=CREMA_CLARO, fg=TEXTO_OSCURO, anchor="w")
            self.lbl_ctx.pack(fill="x")
            tabla_frame = tk.Frame(marco, bg=BLANCO, highlightthickness=1,
                                   highlightbackground=GRIS_BORDE)
            tabla_frame.pack(fill="x", pady=6)
            self.tabla = ttk.Treeview(
                tabla_frame, columns=("fecha", "peso", "ta", "fc", "au", "fcf"),
                show="headings", height=6)
            for col, texto, ancho in (("fecha", "Fecha", 95), ("peso", "Peso", 65),
                                      ("ta", "TA", 70), ("fc", "FC", 55),
                                      ("au", "AU", 55), ("fcf", "FCF", 55)):
                self.tabla.heading(col, text=texto)
                self.tabla.column(col, width=ancho, anchor="center")
            self.tabla.pack(fill="x")
            self.tabla.bind("<<TreeviewSelect>>", self._sel_fila)
            self._idx_map = {}
            self._editando = None

            campos_frame = tk.Frame(marco, bg=CREMA_CLARO)
            campos_frame.pack(fill="x", pady=4)
            campos_frame.grid_columnconfigure(1, weight=1)
            self.vars = {}
            filas = [("Fecha (hoy por defecto)", "fecha"),
                     ("Peso (kg)", "peso"),
                     ("TA sistolica", "ta_s"),
                     ("TA diastolica", "ta_d"),
                     ("FC (lpm)", "fc"),
                     ("Altura uterina (cm)", "au"),
                     ("FCF (lpm)", "fcf"),
                     ("Temperatura (C)", "temp"),
                     ("FR (rpm)", "fr"),
                     ("SpO2 (%)", "spo2"),
                     ("Proteinuria", "proteinuria"),
                     ("Notas", "notas")]
            for i, (lab, key) in enumerate(filas):
                tk.Label(campos_frame, text=lab, bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                         font=(FUENTE_SISTEMA, 10), anchor="w"
                         ).grid(row=i, column=0, sticky="w", pady=2, padx=(0, 8))
                var = tk.StringVar()
                e = tk.Entry(campos_frame, textvariable=var, font=(FUENTE_SISTEMA, 11),
                             bg=BLANCO, relief="solid", bd=1)
                e.grid(row=i, column=1, sticky="we", pady=2)
                self.vars[key] = var
                if key == "fecha":
                    boton_calendario(campos_frame, var, "iso", popup_parent=self
                                     ).grid(row=i, column=2, padx=(6, 0))
            self.vars["fecha"].set(fecha_hoy())

            # C-20: lista de alertas del episodio (circuito cerrado)
            tk.Label(marco, text="ALERTAS DEL EPISODIO", bg=CREMA_CLARO, fg=AZUL,
                     font=(FUENTE_SISTEMA, 9, "bold"), anchor="w"
                     ).pack(fill="x", pady=(8, 2))
            marco_alertas = tk.Frame(marco, bg=BLANCO, highlightthickness=1,
                                     highlightbackground=GRIS_BORDE)
            marco_alertas.pack(fill="x")
            self.tabla_alertas = ttk.Treeview(
                marco_alertas, columns=("fecha", "origen", "estado", "desc"),
                show="headings", height=4)
            for col, texto, ancho in (("fecha", "Fecha", 85), ("origen", "Origen", 95),
                                      ("estado", "Estado", 70), ("desc", "Descripcion", 320)):
                self.tabla_alertas.heading(col, text=texto)
                self.tabla_alertas.column(col, width=ancho)
            self.tabla_alertas.pack(fill="x")
            self._alertas_map = {}

            pie = tk.Frame(marco, bg=CREMA_CLARO)
            pie.pack(fill="x", pady=(8, 0))
            for txt, cmd, color in [
                    ("Nueva visita", self._nueva, AZUL),
                    ("Guardar", self._guardar, VERDE_OK),
                    ("Atender alerta", self._atender_seleccionada, GRIS_SECUNDARIO),
                    ("Revisar Progreso", self._revisar_progreso, GRIS_SECUNDARIO),
                    ("Refrescar", self._refrescar, AZUL),
                    ("Dictar (voz)", self._dictar, GRIS_SECUNDARIO),
                    ("Cerrar", self.destroy, GRIS_SECUNDARIO)]:
                tk.Button(pie, text=txt, command=cmd, bg=color, fg=TEXTO_CLARO,
                          activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 10),
                          padx=8, pady=BTN_PADY).pack(side="left", padx=3)
            self._refrescar()
            self._refrescar_alertas()
            auto_tamano(self, parent_app)

        def _contexto(self):
            ruta = self.app.ruta_activa
            eid = (self.app.datos_activos or {}).get("episodio_activo") if ruta else None
            return ruta, eid

        def _dictar(self):
            """ADR-0040: dictado para notas del control."""
            ok, msg = dictado_disponible()
            win = tk.Toplevel(self)
            win.title("Dictado de voz")
            win.configure(bg=CREMA)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=14, pady=12)
            marco.pack(fill="both", expand=True)
            tk.Label(marco, text="Dictado para notas del control",
                     font=(FUENTE_SISTEMA, 12, "bold"), bg=CREMA_CLARO, fg=AZUL
                     ).pack(anchor="w")
            if ok:
                tk.Label(marco, text="Motor Vosk disponible. Usa el dictado del SO "
                                     "(Win+H / fn fn) en el campo de notas.",
                         bg=CREMA_CLARO, fg=VERDE_OK, wraplength=400,
                         justify="left", font=(FUENTE_SISTEMA, 10)).pack(anchor="w", pady=6)
            else:
                tk.Label(marco, text=f"Dictado disponible:\n{msg}\n\n"
                                     "En Windows: presiona Win+H y dicta.\n"
                                     "En macOS: presiona fn dos veces y dicta.\n"
                                     "En Linux: Super+Ctrl+Espacio.",
                         bg=CREMA_CLARO, fg=TEXTO_OSCURO, wraplength=400,
                         justify="left", font=(FUENTE_SISTEMA, 10)).pack(anchor="w", pady=6)
            tk.Button(marco, text="Cerrar", command=win.destroy, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO, padx=10, pady=6
                      ).pack(anchor="e", pady=8)
            auto_tamano(win, self)

        def _refrescar_alertas(self):
            ruta, eid = self._contexto()
            self.tabla_alertas.delete(*self.tabla_alertas.get_children())
            self._alertas_map = {}
            if not ruta:
                return
            for a in cargar_historia(ruta).get("alertas") or []:
                if eid and a.get("episodio") and a["episodio"] != eid:
                    continue
                estado = "● pendiente" if a.get("estado") == "pendiente" else "✓ atendida"
                iid = self.tabla_alertas.insert("", "end", values=(
                    a.get("fecha", ""), a.get("origen", ""), estado,
                    str(a.get("descripcion", ""))[:90]))
                self._alertas_map[iid] = a

        def _atender_seleccionada(self):
            sel = self.tabla_alertas.selection()
            if not sel or sel[0] not in self._alertas_map:
                messagebox.showinfo("Aviso", "Selecciona una alerta de la lista.",
                                    parent=self)
                return
            alerta = self._alertas_map[sel[0]]
            if alerta.get("estado") != "pendiente":
                messagebox.showinfo("Aviso", "Esa alerta ya fue atendida.",
                                    parent=self)
                return
            ruta, _eid = self._contexto()

            def _listo(respuesta):
                ok, msg = atender_alerta(ruta, alerta.get("id"), respuesta)
                if ok:
                    self.app.toast("Alerta atendida ✓")
                    self.app.status.set(msg)
                    self._refrescar_alertas()
                    self.app._actualizar_etiqueta()
                else:
                    messagebox.showerror("Error", msg, parent=self)
            AtenderAlerta(self, alerta, _listo)

        def _revisar_progreso(self):
            ruta, eid = self._contexto()
            if not ruta:
                messagebox.showinfo("Aviso", "Selecciona una paciente.", parent=self)
                return
            ok, msg, hall = revisar_lab_progreso(ruta, eid)
            vw = tk.Toplevel(self)
            vw.title("Revisar Progreso Quimico")
            vw.configure(bg=CREMA)
            cuerpo = tk.Frame(vw, bg=CREMA_CLARO, padx=12, pady=10)
            cuerpo.pack(fill="both", expand=True)
            tk.Label(cuerpo, text=msg, font=(FUENTE_SISTEMA, 11, "bold"),
                     bg=CREMA_CLARO, fg=ROJO_AVISO if hall else VERDE_OK,
                     anchor="w", wraplength=560, justify="left").pack(fill="x")
            for h in hall:
                tk.Label(cuerpo, text="• " + h.get("mensaje", ""), bg=CREMA_CLARO,
                         fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 9), anchor="w",
                         wraplength=560, justify="left").pack(fill="x", pady=2)
            if not ok and not hall:
                tk.Label(cuerpo, text="(solo lectura: el archivo del medico no se toca)",
                         bg=CREMA_CLARO, fg=GRIS_TEXTO, font=(FUENTE_SISTEMA, 8)
                         ).pack(anchor="w", pady=(4, 0))
            pie2 = tk.Frame(cuerpo, bg=CREMA_CLARO)
            pie2.pack(fill="x", pady=(8, 0))
            tk.Button(pie2, text="Cerrar", command=vw.destroy, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO, padx=10, pady=6
                      ).pack(side="right")
            criticos = [h for h in hall if h.get("nivel") == "critico"]
            if criticos and ok:
                def _altas():
                    n = 0
                    for h in criticos:
                        ok_a, _id = registrar_alerta(ruta, {
                            "origen": f"lab:{h.get('campo', '')}",
                            "descripcion": h.get("mensaje", ""),
                            "episodio": eid or ""})
                        n += 1 if ok_a else 0
                    self.app.toast(f"{n} alerta(s) de laboratorio ⚠")
                    self._refrescar_alertas()
                    self.app._actualizar_etiqueta()
                    vw.destroy()
                tk.Button(pie2, text=f"Registrar {len(criticos)} alerta(s)",
                          command=_altas, bg=ROJO_AVISO, fg=TEXTO_CLARO,
                          activebackground="#C43A4B", font=(FUENTE_SISTEMA, 10, "bold"),
                          padx=10, pady=6).pack(side="right", padx=6)
            auto_tamano(vw, self)

        def _nueva(self):
            self._editando = None
            for k in self.vars:
                self.vars[k].set(fecha_hoy() if k == "fecha" else "")
            self.lbl_ctx.config(
                text=(self.lbl_ctx.cget("text").split("  |  Nueva")[0]
                      + "  |  Nueva visita"))
            self.vars["fecha"].focus_set()

        def _sel_fila(self, event=None):
            sel = self.tabla.selection()
            if not sel:
                return
            idx = self._idx_map.get(sel[0])
            if idx is None:
                return
            ruta, eid = self._contexto()
            datos = cargar_paciente(ruta)
            epi = next((e for e in datos.get("episodios", [])
                        if e.get("id") == eid), None)
            if not epi:
                return
            c = (epi.get("controles") or [])[idx]
            self._editando = idx
            self.vars["fecha"].set(str(c.get("fecha") or ""))
            for k in _CAMPOS_CONTROL:
                v = c.get(k)
                self.vars[k].set("" if v is None else f"{v:g}")
            self.vars["proteinuria"].set(str(c.get("proteinuria") or ""))
            self.vars["notas"].set(str(c.get("notas") or ""))

        def _refrescar(self):
            ruta, eid = self._contexto()
            self.tabla.delete(*self.tabla.get_children())
            self._idx_map = {}
            if not ruta or not eid:
                self.lbl_ctx.config(text="Sin paciente/episodio en contexto",
                                    fg=ROJO_AVISO)
                return
            datos = cargar_paciente(ruta)
            epi = next((e for e in datos.get("episodios", [])
                        if e.get("id") == eid), None)
            almacenados = (epi or {}).get("controles") or []
            nombre = datos.get("nombre_completo", "")
            self.lbl_ctx.config(
                text=f"{nombre}  |  Episodio: {eid}  |  "
                     f"{len(almacenados)} control(es)", fg=VERDE_OK)
            orden = sorted(range(len(almacenados)),
                           key=lambda i: str(almacenados[i].get("fecha") or ""),
                           reverse=True)
            for i in orden:
                c = almacenados[i]
                ta = ""
                if c.get("ta_s") is not None and c.get("ta_d") is not None:
                    ta = f"{c['ta_s']:g}/{c['ta_d']:g}"

                def fmt(v):
                    return "-" if v is None else f"{v:g}"
                iid = self.tabla.insert("", "end", values=(
                    fecha_legible(str(c.get("fecha") or "")) or "-",
                    fmt(c.get("peso")), ta or "-", fmt(c.get("fc")),
                    fmt(c.get("au")), fmt(c.get("fcf"))))
                self._idx_map[iid] = i

        def _guardar(self):
            ruta, eid = self._contexto()
            if not ruta or not eid:
                messagebox.showinfo("Aviso", "Selecciona una paciente y episodio.",
                                    parent=self)
                return
            campos = {k: self.vars[k].get() for k in self.vars}
            # Rangos blandos propios: confirmar, nunca bloquear
            avisos = []
            for k in _CAMPOS_CONTROL:
                msg = rangos_control_fuera(k, _a_numero(campos.get(k)))
                if msg:
                    avisos.append(msg)
            if avisos and not messagebox.askyesno(
                    "Fuera de rango blando",
                    "\n".join(avisos) + "\n\n¿Registrar de todos modos?", parent=self):
                return
            # C-20: catalogos del medico -> panel de hallazgos (confirma, no bloquea)
            hall, pts_q, qmsg, perfil = hallazgos_control(ruta, campos)
            marcar_qsofa = False
            if hall or (pts_q is not None and pts_q >= 2):
                panel_h = HallazgosControl(self, hall, pts_q, qmsg, perfil)
                self.wait_window(panel_h)
                if not panel_h.confirmado:
                    return
                marcar_qsofa = panel_h.marcar_qsofa
            if self._editando is None:
                ok, msg, _ = agregar_control(ruta, eid, campos)
            else:
                ok, msg = actualizar_control(ruta, eid, self._editando, campos)
            if ok:
                self.app.datos_activos = cargar_paciente(ruta)
                self.app.status.set(msg)
                self.app.toast("Control registrado ✓")
                n_altas = 0
                for h in hall:
                    if h.get("nivel") == "critico":
                        ok_a, _id_a = registrar_alerta(ruta, {
                            "origen": str(h.get("campo", "constante")),
                            "descripcion": str(h.get("mensaje", "")),
                            "episodio": eid or ""})
                        if ok_a:
                            n_altas += 1
                if marcar_qsofa and pts_q is not None and pts_q >= 2:
                    ok_a, _id_a = registrar_alerta(ruta, {
                        "origen": "qsofa", "descripcion": qmsg,
                        "episodio": eid or ""})
                    if ok_a:
                        n_altas += 1
                # ADR-0038: alerta de tendencia si el medico configura RANGOS_TENDENCIA
                try:
                    tend_hall = evaluar_tendencia_alerta(ruta, eid)
                    if tend_hall:
                        ok_a, _id_a = registrar_alerta(ruta, {
                            "origen": tend_hall["origen"],
                            "descripcion": tend_hall["descripcion"],
                            "episodio": eid or ""})
                        if ok_a:
                            n_altas += 1
                except Exception:
                    pass
                if n_altas:
                    self.app.toast(f"{n_altas} alerta(s) pendientes ⚠")
                self._nueva()
                self._refrescar()
                self._refrescar_alertas()
            else:
                messagebox.showerror("Error", msg, parent=self)

    class HistoriaCompleta(tk.Toplevel):
        """C-18 (ADR-0030/T16): historia NOM-004 completa A-M con progreso vivo."""

        def __init__(self, parent_app):
            super().__init__(parent_app)
            self.app = parent_app
            self.ruta = parent_app.ruta_activa
            self.title("Historia clinica A-M")
            self.configure(bg=CREMA)
            self.h = cargar_historia(self.ruta)
            self._prefill_f7()
            self._destino = None
            self._widgets = {}
            self._secc = None
            self.nota_idx = -1

            marco = tk.Frame(self, bg=CREMA_CLARO, padx=12, pady=10)
            marco.pack(fill="both", expand=True)

            # --- encabezado + progreso global ---
            cab = tk.Frame(marco, bg=CREMA_CLARO)
            cab.pack(fill="x")
            d = self.h  # solo para titulo
            try:
                nom = cargar_paciente(self.ruta).get("nombre_completo", "")
            except Exception:
                nom = "?"
            tk.Label(cab, text=f"Historia clinica A-M · {nom}",
                     font=(FUENTE_SISTEMA, 13, "bold"), bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).pack(side="left")
            self.barra = ttk.Progressbar(cab, length=180, mode="determinate")
            self.barra.pack(side="right", padx=(8, 0))
            self.lbl_prog = tk.Label(cab, text="0%", bg=CREMA_CLARO, fg=AZUL,
                                     font=(FUENTE_SISTEMA, 10, "bold"))
            self.lbl_prog.pack(side="right")

            # --- notas: selector + altas (sin borrado, C-16 precedent) ---
            fila_notas = tk.Frame(marco, bg=CREMA_CLARO)
            fila_notas.pack(fill="x", pady=(6, 2))
            tk.Label(fila_notas, text="Nota:", bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                     font=(FUENTE_SISTEMA, 10)).pack(side="left")
            self.combo_notas = ttk.Combobox(fila_notas, state="readonly",
                                             width=34, font=(FUENTE_SISTEMA, 10))
            self.combo_notas.pack(side="left", padx=6)
            self.combo_notas.bind("<<ComboboxSelected>>", self._cambiar_nota)
            for txt, tipo in [("+ Nota ingreso", "ingreso"),
                              ("+ Evolucion", "evolucion"),
                              ("+ Egreso", "egreso")]:
                tk.Button(fila_notas, text=txt, command=lambda tp=tipo:
                          self._crear_nota(tp), bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                          activebackground=GRIS_APAGADO, font=(FUENTE_SISTEMA, 9),
                          padx=6, pady=2).pack(side="left", padx=3)
            tk.Button(fila_notas, text="Exportar checklist (.xlsx)",
                      command=self._exportar_checklist, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO,
                      font=(FUENTE_SISTEMA, 9), padx=8, pady=3
                      ).pack(side="right", padx=3)
            tk.Button(fila_notas, text="Vista previa A-M",
                      command=self._vista_previa_v2, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO,
                      font=(FUENTE_SISTEMA, 9), padx=8, pady=3
                      ).pack(side="right", padx=3)
            tk.Button(fila_notas, text="Generar .txt A-M",
                      command=self._generar_v2, bg=AZUL, fg=TEXTO_CLARO,
                      activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 10, "bold"),
                      padx=10, pady=3).pack(side="right", padx=3)
            tk.Button(fila_notas, text="Guardar seccion", command=self._guardar,
                      bg=VERDE_OK, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 10, "bold"), padx=10, pady=3
                      ).pack(side="right", padx=3)
            tk.Button(fila_notas, text="Cerrar", command=self.destroy,
                      bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                      activebackground=GRIS_APAGADO, font=(FUENTE_SISTEMA, 10),
                      padx=10, pady=3).pack(side="right", padx=3)

            # --- notebook A-M ---
            # Solo tira de pestanas: el contenido vive en el canvas con scroll
            self.notebook = ttk.Notebook(marco)
            self.notebook.pack(fill="x", pady=(6, 0))
            self.paneles = {}
            for s in SECCIONES_HISTORIA:
                fr = tk.Frame(self.notebook, bg=CREMA_CLARO)
                self.paneles[s["id"]] = fr
                self.notebook.add(fr, text=f" {s['id']} ")
            self.notebook.bind("<<NotebookTabChanged>>", self._tab_cambiada)

            # --- panel de la seccion activa ---
            self.frame_cab_sec = tk.Frame(marco, bg=CREMA_CLARO)
            self.frame_cab_sec.pack(fill="x", pady=(6, 0))
            self.lbl_sec = tk.Label(self.frame_cab_sec, text="", bg=CREMA_CLARO,
                                    fg=AZUL, font=(FUENTE_SISTEMA, 11, "bold"),
                                    anchor="w", justify="left")
            self.lbl_sec.pack(fill="x")

            self.canvas = tk.Canvas(marco, bg=CREMA_CLARO, highlightthickness=0)
            scr = ttk.Scrollbar(marco, orient="vertical", command=self.canvas.yview)
            self.canvas.configure(yscrollcommand=scr.set)
            self.canvas.pack(side="left", fill="both", expand=True, pady=4)
            scr.pack(side="right", fill="y", pady=4)
            self.contenido = tk.Frame(self.canvas, bg=CREMA_CLARO)
            self._vent = self.canvas.create_window((0, 0), window=self.contenido,
                                                   anchor="nw")
            self.contenido.bind("<Configure>", lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")))
            self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(
                self._vent, width=e.width))

            self._refrescar_notas()
            self._aplicar_seccion(SECCIONES_HISTORIA[0]["id"])
            self.bind("<Control-z>", self._deshacer)
            auto_tamano(self, parent_app)

        def _deshacer(self, _e=None):
            """A13: deshacer la ultima edicion de la seccion paciente activa."""
            sid = self._secc
            if not sid:
                return
            secc = next((s for s in SECCIONES_HISTORIA if s["id"] == sid), None)
            if not secc or secc["nivel"] != "paciente":
                self.app.toast("Deshacer disponible solo en secciones de paciente")
                return
            ok, msg = deshacer_ultima_edicion(self.ruta, sid, secc["nivel"])
            if ok:
                self.h = cargar_historia(self.ruta)
                self.app.toast("Deshacer ✓")
                self._aplicar_seccion(sid)
            else:
                self.app.toast(msg)

        def _prefill_f7(self):
            """C-18 §5: F7 gestaciones/partos/abortos desde el G-P-A (si vacios)."""
            epi = episodio_contexto(self.app.datos_activos or {})
            gpa = (epi or {}).get("gpa") or ""
            m = re.match(r"^[Gg](\d{1,2})[Pp](\d{1,2})[Aa](\d{1,2})$", gpa.strip())
            if not m:
                return
            f7 = self.h.setdefault("paciente", {}).setdefault("F7", {})
            for clave, grupo in (("gestaciones", 1), ("partos", 2), ("abortos", 3)):
                if not str(f7.get(clave, "")).strip():
                    f7[clave] = m.group(grupo)

        # ---------- notas ----------
        def _etiquetas_notas(self):
            out = []
            for i, n in enumerate(self.h.get("notas") or []):
                out.append(f"{i + 1}. {n.get('tipo', 'nota')} "
                           f"{n.get('fecha', '')} {n.get('hora', '')}".strip())
            return out

        def _refrescar_notas(self, seleccionar=None):
            etiquetas = self._etiquetas_notas()
            self.combo_notas["values"] = etiquetas or ["(sin notas)"]
            if seleccionar is not None:
                self.nota_idx = seleccionar
            if etiquetas:
                if not (0 <= self.nota_idx < len(etiquetas)):
                    self.nota_idx = len(etiquetas) - 1
                self.combo_notas.current(self.nota_idx)
            else:
                self.nota_idx = -1
                self.combo_notas.set("(sin notas)")

        def _crear_nota(self, tipo):
            idx = nueva_nota_historia(self.h, tipo,
                                      (self.app.datos_activos or {}).get("episodio_activo", ""))
            self._refrescar_notas(seleccionar=idx)
            self.toast = getattr(self.app, "toast", None)
            if self.toast:
                self.app.toast(f"Nota {tipo} creada ✓")
            sc = next((s for s in SECCIONES_HISTORIA if s["id"] == self._secc), None)
            if sc and sc["nivel"] != "paciente":
                self._aplicar_seccion(self._secc)

        def _cambiar_nota(self, _e=None):
            self._levantar()
            sel = self.combo_notas.current()
            self.nota_idx = sel if sel >= 0 else -1
            if self._secc is not None:
                self._aplicar_seccion(self._secc)

        def _nota_objetivo(self, secc, crear=True):
            if secc["nivel"] == "paciente":
                return None
            notas = self.h.setdefault("notas", [])
            eid = (self.app.datos_activos or {}).get("episodio_activo", "")
            tipo_pref = {"L": "evolucion", "M": "egreso"}.get(secc["id"])
            if tipo_pref:
                if (0 <= self.nota_idx < len(notas)
                        and notas[self.nota_idx].get("tipo") == tipo_pref):
                    return notas[self.nota_idx]
                compatibles = [i for i, n in enumerate(notas)
                               if n.get("tipo") == tipo_pref]
                if compatibles:
                    self.nota_idx = compatibles[-1]
                    self.combo_notas.current(self.nota_idx)
                    return notas[self.nota_idx]
                if crear:
                    self.nota_idx = nueva_nota_historia(self.h, tipo_pref, eid)
                    self._refrescar_notas(seleccionar=self.nota_idx)
                    return notas[self.nota_idx]
                return None
            if 0 <= self.nota_idx < len(notas):
                return notas[self.nota_idx]
            if notas:
                self.nota_idx = len(notas) - 1
                self.combo_notas.current(self.nota_idx)
                return notas[self.nota_idx]
            if crear:
                self.nota_idx = nueva_nota_historia(self.h, "ingreso", eid)
                self._refrescar_notas(seleccionar=self.nota_idx)
                return notas[self.nota_idx]
            return None

        # ---------- navegacion de secciones ----------
        def _tab_cambiada(self, _e=None):
            try:
                fr = self.notebook.select()
                sid = None
                for k, v in self.paneles.items():
                    if str(v) == str(fr):
                        sid = k
                        break
                if sid is not None and sid != self._secc:
                    self._aplicar_seccion(sid)
            except Exception:
                pass

        def _aplicar_seccion(self, sid):
            self._levantar()
            self._secc = sid
            secc = next(s for s in SECCIONES_HISTORIA if s["id"] == sid)
            # construir la seccion activa en el contenedor con scroll
            for w in self.contenido.winfo_children():
                w.destroy()
            marco_sec = tk.Frame(self.contenido, bg=CREMA_CLARO)
            marco_sec.pack(fill="both", expand=True, padx=8, pady=6)
            inv, tot = progreso_historia(self.h).get(sid, (0, 0))
            nivel_txt = {"paciente": "paciente (acumulativa)",
                         "nota": "nota", "evolucion": "evolucion (por visita)",
                         "egreso": "egreso (por episodio)"}[secc["nivel"]]
            self.lbl_sec.config(
                text=f"{secc['id']} · {secc['titulo']}   [{nivel_txt}]   "
                     f"investigados {inv}/{tot}")
            self._render(secc, marco_sec)

        def _levantar(self):
            """Vuelve los widgets de la seccion anterior al diccionario de trabajo."""
            if not self._widgets or self._destino is None:
                return
            for item, lectura in self._widgets.items():
                try:
                    self._destino[item] = lectura()
                except Exception:
                    pass
            self._widgets = {}
            self._destino = None
            self._actualizar_progreso()

        def _actualizar_progreso(self):
            pr = progreso_historia(self.h)
            inv = sum(a for a, _ in pr.values())
            tot = sum(b for _, b in pr.values()) or 1
            pct = int(round(inv * 100 / tot))
            self.barra["value"] = pct
            self.lbl_prog.config(text=f"{pct}%")

        # ---------- render de items ----------
        def _render(self, secc, contenedor):
            nota = self._nota_objetivo(secc)
            if secc["nivel"] == "paciente":
                self._destino = self.h.setdefault("paciente", {}).setdefault(
                    secc["id"], {})
            else:
                self._destino = (nota or {}).setdefault("datos", {}).setdefault(
                    secc["id"], {})
            ctx = contexto_historia(self.ruta)
            ctx["nota"] = nota or {}
            filas = tk.Frame(contenedor, bg=CREMA_CLARO)
            filas.pack(fill="both", expand=True)
            filas.grid_columnconfigure(1, weight=1)
            for i, it in enumerate(secc["items"]):
                tk.Label(filas, text=it["lbl"], bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                         font=(FUENTE_SISTEMA, 9), anchor="nw", justify="left",
                         wraplength=240).grid(row=i, column=0, sticky="nw",
                                              pady=2, padx=(0, 8))
                valor = self._destino.get(it["id"], "")
                w = self._render_item(filas, secc, it, valor, ctx)
                w.grid(row=i, column=1, sticky="we", pady=2)

        def _render_item(self, parent, secc, it, valor, ctx):
            t = it["t"]
            if t == "auto":
                txt = auto_historia(secc["id"], it["id"], ctx)
                e = tk.Entry(parent, font=(FUENTE_SISTEMA, 10), bg="#ECECEC",
                             fg=GRIS_TEXTO, state="readonly")
                e.insert(0, txt)
                return e
            if t in ("texto", "numero", "fecha"):
                var = tk.StringVar(value="" if valor in (None, "") else str(valor))
                if t == "fecha":
                    caja = tk.Frame(parent, bg=CREMA_CLARO)
                    e = tk.Entry(caja, textvariable=var, font=(FUENTE_SISTEMA, 10),
                                 bg=BLANCO)
                    e.pack(side="left", fill="x", expand=True)
                    boton_calendario(caja, var, "iso", popup_parent=self
                                     ).pack(side="left", padx=(4, 0))
                    self._widgets[it["id"]] = lambda v=var: v.get().strip()
                    return caja
                e = tk.Entry(parent, textvariable=var, font=(FUENTE_SISTEMA, 10),
                             bg=BLANCO)
                self._widgets[it["id"]] = lambda v=var: v.get().strip()
                return e
            if t == "enum":
                ops = list(it.get("ops") or [])
                var = tk.StringVar(value="" if valor in (None, "") else str(valor))
                cb = ttk.Combobox(parent, textvariable=var, values=ops,
                                  font=(FUENTE_SISTEMA, 10))
                self._widgets[it["id"]] = lambda v=var: v.get().strip()
                return cb
            if t == "libre":
                box = tk.Text(parent, height=3, font=(FUENTE_SISTEMA, 10), bg=BLANCO,
                              relief="solid", bd=1)
                box.insert("1.0", "" if valor in (None, "") else str(valor))
                self._widgets[it["id"]] = (lambda b=box:
                                           b.get("1.0", "end").strip())
                return box
            # radios tri-state (+ extras)
            caja = tk.Frame(parent, bg=CREMA_CLARO)
            opciones = [("", "(sin investigar)"), ("si", "Si"), ("no", "No")]
            if t == "snns":
                opciones.append(("no_sabe", "No sabe"))
            if t == "snna":
                opciones.append(("no_aplica", "No aplica"))
            estado = ""
            extras = {}
            if t in ("snv", "snc", "snh") and isinstance(valor, dict):
                estado = str(valor.get("e", ""))
                extras = {k: str(v) for k, v in valor.items() if k != "e"}
            elif isinstance(valor, str):
                estado = valor
            var = tk.StringVar(value=estado)
            for val, lab in opciones:
                tk.Radiobutton(caja, text=lab, variable=var, value=val,
                               bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                               selectcolor=BLANCO, font=(FUENTE_SISTEMA, 9)
                               ).pack(side="left", padx=(0, 4))
            ents = {}
            if t == "snv":
                e1 = tk.Entry(caja, font=(FUENTE_SISTEMA, 9), width=18, bg=BLANCO)
                e1.insert(0, extras.get("v", ""))
                e1.pack(side="left", padx=4)
                ents["v"] = e1
            elif t == "snc":
                for k, lab, ancho in (("d1", "Edad/dx", 8), ("d2", "Tratam.", 14)):
                    tk.Label(caja, text=lab, bg=CREMA_CLARO, fg=GRIS_TEXTO,
                             font=(FUENTE_SISTEMA, 8)).pack(side="left")
                    e1 = tk.Entry(caja, font=(FUENTE_SISTEMA, 9), width=ancho, bg=BLANCO)
                    e1.insert(0, extras.get(k, ""))
                    e1.pack(side="left", padx=(1, 4))
                    ents[k] = e1
            elif t == "snh":
                e1 = tk.Entry(caja, font=(FUENTE_SISTEMA, 9), width=30, bg=BLANCO)
                e1.insert(0, extras.get("h", ""))
                e1.pack(side="left", padx=4)
                ents["h"] = e1

            def leer(v=var, tipo=t, es=ents):
                estado = v.get()
                if tipo in ("sn", "snns", "snna"):
                    return estado
                salida = {"e": estado}
                for k, e in es.items():
                    salida[k] = e.get().strip()
                return salida
            self._widgets[it["id"]] = leer
            return caja

        def _guardar(self):
            self._levantar()
            ok, msg = guardar_historia(self.ruta, self.h)
            if ok:
                self.app.toast("Historia guardada ✓")
                self.app.status.set(msg)
                self._actualizar_progreso()
            else:
                messagebox.showerror("Error", msg, parent=self)

        def _vista_previa_v2(self):
            self._levantar()
            ok_g, msg_g = guardar_historia(self.ruta, self.h)
            if not ok_g:
                messagebox.showerror("Error", msg_g, parent=self)
                return
            texto = render_historia_v2(self.ruta)
            vw = tk.Toplevel(self)
            vw.title("Vista previa A-M (aun no generada)")
            vw.configure(bg=CREMA)
            caja = tk.Frame(vw, bg=CREMA_CLARO, padx=10, pady=8)
            caja.pack(fill="both", expand=True)
            tk.Label(caja, text="Historia completa A-M · vista previa",
                     font=(FUENTE_SISTEMA, 11, "bold"), bg=CREMA_CLARO, fg=AZUL
                     ).pack(anchor="w")
            box = tk.Text(caja, width=92, height=26, font=("Courier", 10),
                          bg=BLANCO, relief="solid", bd=1, wrap="none")
            box.insert("1.0", texto)
            box.config(state="disabled")
            box.pack(fill="both", expand=True, pady=6)
            tk.Button(caja, text="Cerrar", command=vw.destroy, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO, padx=10, pady=6
                      ).pack(anchor="e")
            auto_tamano(vw, self)

        def _generar_v2(self):
            self._levantar()
            ok_g, msg_g = guardar_historia(self.ruta, self.h)
            if not ok_g:
                messagebox.showerror("Error", msg_g, parent=self)
                return
            ok, msg, salida = generar_historia_v2(self.ruta)
            if ok:
                self.app.toast("Historia completa generada ✓")
                self.destroy()
                ir_a_ruta(salida)
                messagebox.showinfo("Historia clinica A-M",
                                    f"{msg}\nEl formato requiere firma profesional "
                                    "para validez oficial.")
            else:
                messagebox.showerror("Error", msg, parent=self)

        def _exportar_checklist(self):
            self._levantar()
            ok_g, msg_g = guardar_historia(self.ruta, self.h)
            if not ok_g:
                messagebox.showerror("Error", msg_g, parent=self)
                return
            ok, msg, salida = exportar_checklist_xlsx(self.ruta)
            if ok:
                self.app.toast("Checklist exportado ✓")
                if messagebox.askyesno("Checklist exportado",
                                       f"{msg}\n\n¿Abrir el archivo?", parent=self):
                    abrir_archivo(salida)
            else:
                messagebox.showerror("Error", msg, parent=self)

    class HallazgosControl(tk.Toplevel):
        """C-20: panel de hallazgos al guardar — CONFIRMAR, NUNCA BLOQUEAR."""

        def __init__(self, parent, hallazgos, qsofa_pts, qmsg, perfil):
            super().__init__(parent)
            self.title("Revisar constantes antes de guardar")
            self.configure(bg=CREMA)
            self.confirmado = False
            self.marcar_qsofa = False
            marco = tk.Frame(self, bg=CREMA_CLARO, padx=14, pady=12)
            marco.pack(fill="both", expand=True)
            tk.Label(marco, text=f"Perfil: {perfil} · hallazgos: {len(hallazgos)}",
                     font=(FUENTE_SISTEMA, 12, "bold"), bg=CREMA_CLARO, fg=AZUL
                     ).pack(anchor="w")
            colores = {"critico": "#F8C9C4", "alerta": "#FFE9B8",
                       "esperado_embarazo": AZUL_CHIP}
            if not hallazgos:
                tk.Label(marco, text="Todas las constantes dentro de rango.",
                         bg=CREMA_CLARO, fg=VERDE_OK,
                         font=(FUENTE_SISTEMA, 10)).pack(anchor="w", pady=4)
            for h in hallazgos:
                nivel = h.get("nivel", "alerta")
                fila = tk.Frame(marco, bg=colores.get(nivel, BLANCO),
                                highlightthickness=1,
                                highlightbackground=GRIS_BORDE)
                fila.pack(fill="x", pady=3)
                tk.Label(fila, text=nivel.upper(), bg=colores.get(nivel, BLANCO),
                         fg=ROJO_AVISO if nivel == "critico" else TEXTO_OSCURO,
                         font=(FUENTE_SISTEMA, 9, "bold"), padx=6, pady=4
                         ).pack(side="left")
                tk.Label(fila, text=h.get("mensaje", ""), bg=colores.get(nivel, BLANCO),
                         fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 9), wraplength=520,
                         justify="left", padx=6, pady=4).pack(side="left", fill="x")
            # qSOFA
            if qsofa_pts is not None:
                crit = qsofa_pts >= 2
                tk.Label(marco, text=qmsg, bg=CREMA_CLARO,
                         fg=ROJO_AVISO if crit else TEXTO_OSCURO,
                         font=(FUENTE_SISTEMA, 10, "bold"), anchor="w", wraplength=540,
                         justify="left").pack(fill="x", pady=(6, 0))
                if crit:
                    tk.Button(marco, text="Marcar alerta (qSOFA)",
                              command=self._marcar_qsofa, bg=ROJO_AVISO,
                              fg=TEXTO_CLARO, activebackground="#C43A4B",
                              font=(FUENTE_SISTEMA, 10, "bold"), padx=10, pady=6
                              ).pack(anchor="w", pady=4)
            else:
                tk.Label(marco, text=qmsg, bg=CREMA_CLARO, fg=GRIS_TEXTO,
                         font=(FUENTE_SISTEMA, 9), anchor="w", wraplength=540,
                         justify="left").pack(fill="x", pady=(6, 0))
            pie = tk.Frame(marco, bg=CREMA_CLARO)
            pie.pack(fill="x", pady=(10, 0))
            tk.Button(pie, text="Cancelar", command=self.destroy,
                      bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                      activebackground=GRIS_APAGADO, padx=10, pady=6
                      ).pack(side="left")
            tk.Button(pie, text="Registrar de todos modos", command=self._aceptar,
                      bg=VERDE_OK, fg=TEXTO_CLARO, activebackground="#14602F",
                      font=(FUENTE_SISTEMA, 11, "bold"), padx=12, pady=6
                      ).pack(side="right")
            auto_tamano(self, parent)

        def _aceptar(self):
            self.confirmado = True
            self.destroy()

        def _marcar_qsofa(self):
            self.marcar_qsofa = True

        def on_close(self):
            self.destroy()

    class AtenderAlerta(tk.Toplevel):
        """C-20: respuesta registrada (texto + medico + cedula) — circuito cerrado."""

        def __init__(self, parent, alerta, on_listo):
            super().__init__(parent)
            self.title(f"Atender alerta {alerta.get('id', '')}")
            self.configure(bg=CREMA)
            self.on_listo = on_listo
            marco = tk.Frame(self, bg=CREMA_CLARO, padx=14, pady=12)
            marco.pack(fill="both", expand=True)
            tk.Label(marco, text=f"{alerta.get('origen', '')} · "
                                 f"{alerta.get('fecha', '')}",
                     font=(FUENTE_SISTEMA, 11, "bold"), bg=CREMA_CLARO, fg=ROJO_AVISO,
                     anchor="w", wraplength=520, justify="left").pack(fill="x")
            tk.Label(marco, text=alerta.get("descripcion", ""), bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 9), anchor="w",
                     wraplength=520, justify="left").pack(fill="x", pady=(4, 8))
            tk.Label(marco, text="Respuesta / evolucion:", bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 10)).pack(anchor="w")
            self.txt = tk.Text(marco, height=4, width=64, font=(FUENTE_SISTEMA, 10),
                               bg=BLANCO, relief="solid", bd=1)
            self.txt.pack(fill="x", pady=3)
            fila = tk.Frame(marco, bg=CREMA_CLARO)
            fila.pack(fill="x", pady=(4, 0))
            tk.Label(fila, text=f"Medico: {autor_actual()}", bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 9)).pack(side="left")
            tk.Label(fila, text="Cedula:", bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                     font=(FUENTE_SISTEMA, 9)).pack(side="right")
            self.e_cedula = tk.Entry(fila, font=(FUENTE_SISTEMA, 10), width=16, bg=BLANCO)
            self.e_cedula.pack(side="right", padx=4)
            pie = tk.Frame(marco, bg=CREMA_CLARO)
            pie.pack(fill="x", pady=(10, 0))
            tk.Button(pie, text="Cancelar", command=self.destroy, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO, padx=10, pady=6
                      ).pack(side="left")
            tk.Button(pie, text="Registrar respuesta", command=self._guardar,
                      bg=VERDE_OK, fg=TEXTO_CLARO, activebackground="#14602F",
                      font=(FUENTE_SISTEMA, 11, "bold"), padx=12, pady=6
                      ).pack(side="right")
            auto_tamano(self, parent)

        def _guardar(self):
            texto = self.txt.get("1.0", "end").strip()
            if not texto:
                messagebox.showwarning("Aviso",
                                       "Escribe la respuesta/evolucion.", parent=self)
                return
            self.on_listo({"texto": texto, "medico": autor_actual(),
                           "cedula": self.e_cedula.get().strip()})
            self.destroy()


    class Calculadoras(tk.Frame):
        """C-19 (ADR-0032/T19 + ADR-0042/v5.0): pagina embebida de calculadoras
        del medico en el area principal (sin ventana nueva). Estratifica, no
        diagnostica (ADR-0013)."""

        def __init__(self, parent_app, container=None):
            super().__init__(container or parent_app, bg=CREMA)
            self.app = parent_app
            marco = tk.Frame(self, bg=CREMA_CLARO, padx=12, pady=10)
            marco.pack(fill="both", expand=True)
            # ADR-0042: banner de contexto (reemplaza al messagebox de apertura)
            self.lbl_ctx = tk.Label(marco, text="", font=(FUENTE_SISTEMA, 10, "bold"),
                                    bg=CREMA_CLARO, anchor="w")
            self.lbl_ctx.pack(fill="x", pady=(0, 6))
            # ADR-0044: casillero PERFILES — embarazada/adulto, rapido y visible.
            # Arranca auto (perfil_control = EG valida del episodio); el usuario
            # puede sobreescribirlo hasta que cambie de paciente.
            fila_perfil = tk.Frame(marco, bg="#EEF3FA", highlightthickness=1,
                                   highlightbackground=AZUL)
            fila_perfil.pack(fill="x", pady=(0, 4))
            self.var_embarazada = tk.BooleanVar(value=False)
            self._perfil_manual = False
            tk.Checkbutton(
                fila_perfil, text="PACIENTE EMBARAZADA", variable=self.var_embarazada,
                command=self._cambiar_perfil, onvalue=True, offvalue=False,
                bg="#EEF3FA", fg=AZUL, activebackground="#EEF3FA",
                selectcolor=BLANCO, font=(FUENTE_SISTEMA, 11, "bold"),
                padx=8, pady=4).pack(side="left")
            self.lbl_perfil = tk.Label(fila_perfil, text="", bg="#EEF3FA",
                                       fg=GRIS_TEXTO, font=(FUENTE_SISTEMA, 9),
                                       anchor="w", justify="left", wraplength=480)
            self.lbl_perfil.pack(side="left", fill="x", expand=True, padx=(4, 8),
                                 pady=4)
            self._actualizar_lbl_perfil()
            # grupo + lista
            fila = tk.Frame(marco, bg=CREMA_CLARO)
            fila.pack(fill="x")
            tk.Label(fila, text="Grupo:", bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                     font=(FUENTE_SISTEMA, 10)).pack(side="left")
            # ADR-0043: ★ Prioritarias encabeza; defecto = composicion corporal
            self._grupos_ui = [GRUPO_PRIORITARIAS] + list(GRUPOS_CALCULADORAS)
            self.var_grupo = tk.StringVar(value=GRUPO_PRIORITARIAS)
            cb = ttk.Combobox(fila, textvariable=self.var_grupo,
                              values=self._grupos_ui, state="readonly",
                              width=32, font=(FUENTE_SISTEMA, 10))
            cb.pack(side="left", padx=6)
            cb.bind("<<ComboboxSelected>>", lambda e: self._poblar_lista())
            self.lista = tk.Listbox(marco, height=7, font=(FUENTE_SISTEMA, 10),
                                    bg=BLANCO, fg=TEXTO_OSCURO, relief="solid",
                                    bd=1, selectbackground=AZUL,
                                    selectforeground=TEXTO_CLARO)
            self.lista.pack(fill="x", pady=(6, 0))
            self.lista.bind("<<ListboxSelect>>", lambda e: self._cargar_calc())
            # ADR-0043: fijar/quitar la calculadora actual de ★ Prioritarias
            self.btn_star = tk.Button(marco, text="☆ Prioritaria", width=14,
                                      command=self._alternar_prioritaria,
                                      bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                                      activebackground=GRIS_APAGADO,
                                      font=(FUENTE_SISTEMA, 9), pady=2)
            self.btn_star.pack(side="right", pady=(8, 2))
            Tooltip(self.btn_star,
                    "Fija esta calculadora en ★ Prioritarias (flujo rapido); "
                    "el consultorio define su lista en _CONFIG.json")
            self.lbl_titulo = tk.Label(marco, text="", font=(FUENTE_SISTEMA, 12, "bold"),
                                       bg=CREMA_CLARO, fg=AZUL, anchor="w")
            self.lbl_titulo.pack(side="left", fill="x", expand=True, pady=(8, 2))
            self.campos_frame = tk.Frame(marco, bg=CREMA_CLARO)
            self.campos_frame.pack(fill="x")
            self.lbl_formula = tk.Label(marco, text="", bg=CREMA_CLARO,
                                        fg=GRIS_TEXTO, font=(FUENTE_SISTEMA, 9),
                                        anchor="w", justify="left", wraplength=560)
            self.lbl_formula.pack(fill="x", pady=(4, 0))
            # resultado
            self.caja_res = tk.Frame(marco, bg="#F4F7FB", highlightthickness=1,
                                     highlightbackground=GRIS_BORDE)
            self.caja_res.pack(fill="x", pady=(8, 0))
            self.lbl_res_valor = tk.Label(self.caja_res, text="—",
                                          font=(FUENTE_SISTEMA, 22, "bold"),
                                          bg="#F4F7FB", fg=GRIS_TEXTO, anchor="w")
            self.lbl_res_valor.pack(fill="x", padx=12, pady=(8, 0))
            self.lbl_res_banda = tk.Label(self.caja_res, text="Introduce los datos y pulsa Calcular",
                                          font=(FUENTE_SISTEMA, 10, "bold"),
                                          bg="#F4F7FB", fg=GRIS_TEXTO, anchor="w")
            self.lbl_res_banda.pack(fill="x", padx=12, pady=(0, 4))
            self.lbl_res_msg = tk.Label(self.caja_res, text="", bg="#F4F7FB",
                                        fg=GRIS_TEXTO, font=(FUENTE_SISTEMA, 9),
                                        anchor="w", wraplength=560, justify="left")
            self.lbl_res_msg.pack(fill="x", padx=12, pady=(0, 8))
            pie = tk.Frame(marco, bg=CREMA_CLARO)
            pie.pack(fill="x", pady=(8, 0))
            tk.Button(pie, text="Guardar en la nota", command=self._guardar_nota,
                      bg=VERDE_OK, fg=TEXTO_CLARO, activebackground="#14602F",
                      font=(FUENTE_SISTEMA, 10, "bold"), padx=10, pady=6
                      ).pack(side="left")
            # ADR-0034: sin boton Calcular; recalculo en vivo via trace
            tk.Label(pie, text="Resultado en vivo — sin boton",
                     bg=CREMA_CLARO, fg=GRIS_TERCIARIO,
                     font=(FUENTE_SISTEMA, 8)).pack(side="right", padx=6)
            tk.Button(pie, text="Ver pacientes",
                      command=lambda: self.app.show_page("home"),
                      bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                      activebackground=GRIS_APAGADO, padx=10, pady=6
                      ).pack(side="right", padx=6)
            tk.Label(marco, text="Estratifica riesgo; no sustituye el juicio clinico "
                                 "(ADR-0013). Cifras del catalogo del medico.",
                     bg=CREMA_CLARO, fg=GRIS_TERCIARIO, font=(FUENTE_SISTEMA, 8),
                     anchor="w", wraplength=560, justify="left").pack(fill="x",
                                                                      pady=(6, 0))
            self._calc_actual = None
            self._vars = {}
            self._ultimo = None
            self._poblar_lista()
            self.refrescar_ctx()

        def refrescar_ctx(self):
            """ADR-0042: banner de paciente/episodio en contexto (pagina)."""
            d = self.app.datos_activos if self.app.ruta_activa else None
            if d:
                epi = episodio_contexto(d) or {}
                eid = d.get("episodio_activo") or epi.get("id", "")
                self.lbl_ctx.config(
                    text=(f"{d.get('nombre_completo', '')}   |   Episodio: "
                          f"{eid or '(sin episodio)'}"),
                    fg=VERDE_OK)
            else:
                self.lbl_ctx.config(
                    text="Sin paciente en contexto — puedes calcular; "
                         "'Guardar en la nota' lo exigira.",
                    fg=ROJO_AVISO)
            # ADR-0044: al entrar (o cambiar de paciente) el perfil vuelve a
            # auto-sincronizarse con el episodio (FUM/EG valida -> embarazada)
            if not self._perfil_manual:
                auto = bool(d) and perfil_control(self.app.ruta_activa) == "embarazada"
                self.var_embarazada.set(auto)
                self._actualizar_lbl_perfil()

        def _info_perfil(self):
            """(texto_hint, perfil) segun el casillero de perfil embarazada."""
            if self.var_embarazada.get():
                r = RANGOS_SIGNOS_VITALES["embarazada"]
                return (f"Perfil EMBARAZADA — referencias: "
                        f"TA {r['ta_s'][0]}-{r['ta_s'][1]}/{r['ta_d'][0]}-{r['ta_d'][1]} · "
                        f"FC {r['fc'][0]}-{r['fc'][1]} · FR {r['fr'][0]}-{r['fr'][1]} · "
                        f"Temp {r['temp'][0]}-{r['temp'][1]} · SpO2 >={r['spo2'][0]}%",
                        "embarazada")
            r = RANGOS_SIGNOS_VITALES["adulto"]
            return (f"Perfil adulto — referencias: "
                    f"TA {r['ta_s'][0]}-{r['ta_s'][1]}/{r['ta_d'][0]}-{r['ta_d'][1]} · "
                    f"FC {r['fc'][0]}-{r['fc'][1]} · FR {r['fr'][0]}-{r['fr'][1]} · "
                    f"Temp {r['temp'][0]}-{r['temp'][1]} · SpO2 >={r['spo2'][0]}%",
                    "adulto")

        def _actualizar_lbl_perfil(self):
            """Muestra el perfil activo y sus referencias (catálogo del médico)."""
            texto, perfil = self._info_perfil()
            origen = "auto (según episodio)" if not self._perfil_manual else "manual"
            if perfil == "embarazada":
                self.lbl_perfil.config(text=f"{texto}  ·  {origen}",
                                       fg=ROJO_AVISO)
            else:
                self.lbl_perfil.config(text=f"{texto}  ·  {origen}",
                                       fg=GRIS_TEXTO)

        def _cambiar_perfil(self):
            """ADR-0044: casillero embarazada/adulto; recalcula con el nuevo perfil."""
            self._perfil_manual = True
            self._actualizar_lbl_perfil()
            self.app.status.set(
                "Perfil: " + ("EMBARAZADA" if self.var_embarazada.get()
                              else "adulto") + " (catálogo del médico)")
            if self._calc_actual:
                self._calcular()  # reevalua SV del nuevo perfil en vivo

        def _poblar_lista(self):
            self.lista.delete(0, "end")
            grupo = self.var_grupo.get()
            if grupo == GRUPO_PRIORITARIAS:
                # ADR-0043: orden definido por el consultorio (no el del catalogo)
                self.ids_grupo = list(calculadoras_prioritarias())
            else:
                self.ids_grupo = [cid for cid, c in CALCULADORAS.items()
                                  if c["grupo"] == grupo]
            if not self.ids_grupo:
                self.lista.insert(
                    "end", " (Vacio: fija calculadoras con el boton ★) ")
                self._calc_actual = None
                self._actualizar_star()
                return
            for cid in self.ids_grupo:
                c = CALCULADORAS[cid]
                suf = "" if c["estado"] == "activa" else "  [pendiente]"
                pre = "★ " if grupo == GRUPO_PRIORITARIAS else ""
                self.lista.insert("end", f"{pre}{c['titulo']}{suf}")
            if self.ids_grupo:
                self.lista.selection_set(0)
                self._cargar_calc()

        def _actualizar_star(self):
            """ADR-0043: estado del boton ☆/★ segun la calculadora cargada."""
            en = self._calc_actual in calculadoras_prioritarias()
            self.btn_star.config(
                text="★ Prioritaria" if en else "☆ Prioritaria",
                bg=ACENTO if en else GRIS_SECUNDARIO,
                fg=NAV_BG if en else TEXTO_CLARO)

        def _alternar_prioritaria(self):
            """Fija/quita la calculadora actual en ★ Prioritarias (persiste en config)."""
            if not self._calc_actual:
                return
            quedo = alternar_prioritaria(self._calc_actual)
            self.app.status.set(
                f"{CALCULADORAS[self._calc_actual]['titulo']}: "
                + ("añadida a" if quedo else "quitada de") + " ★ Prioritarias")
            self.app.toast("★ Guardada en Prioritarias" if quedo
                           else "☆ Quitada de Prioritarias")
            self._actualizar_star()
            if self.var_grupo.get() == GRUPO_PRIORITARIAS:
                # refresca la lista conservando la calculadora visible
                self._poblar_lista()

        def _cargar_calc(self):
            sel = self.lista.curselection()
            if not sel or sel[0] >= len(self.ids_grupo):
                return
            self._levantar_cfg()
            cid = self.ids_grupo[sel[0]]
            c = CALCULADORAS[cid]
            self._calc_actual = cid
            self.lbl_titulo.config(text=c["titulo"])
            self.lbl_formula.config(text=f"Fórmula: {c.get('formula', '')}")
            self._actualizar_star()
            for w in self.campos_frame.winfo_children():
                w.destroy()
            self._vars = {}
            self.campos_frame.grid_columnconfigure(1, weight=1)
            for i, (fid, lbl, tipo) in enumerate(c["campos"]):
                if tipo == "bool":
                    var = tk.BooleanVar(value=False)
                    tk.Checkbutton(self.campos_frame, text=lbl, variable=var,
                                   bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                                   selectcolor=BLANCO, font=(FUENTE_SISTEMA, 10),
                                   anchor="w").grid(row=i, column=0, columnspan=2,
                                                    sticky="w", pady=1)
                else:
                    tk.Label(self.campos_frame, text=lbl, bg=CREMA_CLARO,
                             fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 10), anchor="w"
                             ).grid(row=i, column=0, sticky="w", pady=1,
                                    padx=(0, 8))
                    var = tk.StringVar()
                    if tipo == "txt" and fid == "sexo":
                        w = ttk.Combobox(self.campos_frame, textvariable=var,
                                         values=("M", "F"), width=8,
                                         font=(FUENTE_SISTEMA, 10))
                    else:
                        w = tk.Entry(self.campos_frame, textvariable=var,
                                     font=(FUENTE_SISTEMA, 10), bg=BLANCO, width=12)
                    w.grid(row=i, column=1, sticky="w", pady=1)
                    w.bind("<Return>", lambda e: self._calcular())
                self._vars[fid] = (tipo, var)
                if tipo != "bool":
                    var.trace_add("write", lambda *_: self._recalc_en_vivo())
            self._prefill_control()
            # limpiar resultado previo
            self.lbl_res_valor.config(text="—", fg=GRIS_TEXTO)
            self.lbl_res_banda.config(text="Introduce los datos y pulsa Calcular",
                                      fg=GRIS_TEXTO)
            self.lbl_res_banda.config(bg="#F4F7FB")
            self.lbl_res_valor.config(bg="#F4F7FB")
            self.lbl_res_msg.config(text="", bg="#F4F7FB")

        def _prefill_control(self):
            """ADR-0034/v4.16: precarga los campos desde el ultimo control de la
            paciente en contexto (carril A: calculadora no registra, solo lee)."""
            ruta = self.app.ruta_activa
            if not ruta or not self._calc_actual:
                return
            try:
                epi = episodio_contexto(self.app.datos_activos or {})
                eid = (epi or {}).get("id", "")
                ctrls = listar_controles(ruta, eid)
                if not ctrls:
                    return
                ultimo = max(ctrls, key=lambda c: str(c.get("fecha", "")))
            except Exception:
                return
            for fid, (tipo, var) in self._vars.items():
                if tipo != "num":
                    continue
                try:
                    if str(var.get()).strip():
                        continue  # no pisar lo que el usuario ya tecleo
                except Exception:
                    pass
                v = ultimo.get(fid)
                if isinstance(v, (int, float)):
                    try:
                        var.set(f"{v:g}")
                    except Exception:
                        var.set(str(v))

        def _levantar_cfg(self):
            self._vars = {}

        def _entradas(self):
            out = {}
            for fid, (tipo, var) in self._vars.items():
                if tipo == "bool":
                    out[fid] = bool(var.get())
                elif tipo == "num":
                    txt = str(var.get()).strip().replace(",", ".")
                    if txt == "":
                        continue
                    try:
                        out[fid] = float(txt)
                    except ValueError:
                        out[fid] = txt  # la formula decidira si es valido
                else:
                    out[fid] = str(var.get()).strip()
            return out

        def _recalc_en_vivo(self):
            """ADR-0034: recalculo al teclear (debounce 150ms, sin boton)."""
            if hasattr(self, "_after_recalc"):
                try:
                    self.after_cancel(self._after_recalc)
                except Exception:
                    pass
            self._after_recalc = self.after(150, self._calcular)

        def _plausibilidad_en_vivo(self):
            """ADR-0035/A3: validacion blando en vivo en C-16 (nunca bloquea)."""
            if not hasattr(self, "_lbl_plaus"):
                return
            valores = {k: self.vars[k].get() for k in self.vars}
            h = validar_plausibilidad(valores)
            if h:
                self._lbl_plaus.config(text=h[0]["mensaje"], fg=ROJO_AVISO)
            else:
                self._lbl_plaus.config(text="Valores plausibles", fg=VERDE_OK)

        def _calcular(self):
            if not self._calc_actual:
                return
            entradas = self._entradas()
            ok, res, msg = calcular(self._calc_actual, entradas)
            # ADR-0044: referencia de SV del perfil activo (informativo, P2)
            _txt_perfil, perfil = self._info_perfil()
            hall_perfil = hallazgos_perfil_entradas(entradas, perfil)
            nota_perfil = ""
            if hall_perfil:
                nota_perfil = ("  ·  ⚠ " + hall_perfil[0]["mensaje"]
                               + (" (esperado en embarazo: evaluar con síntomas)"
                                  if hall_perfil[0]["nivel"] == "esperado_embarazo"
                                  else ""))
            colores_fondo = {"ok": "#DFF3E4", "alerta": "#FFE9B8",
                             "critico": "#F8C9C4"}
            if ok:
                nivel = (res or {}).get("nivel", "")
                self.lbl_res_valor.config(text=f"{(res or {}).get('valor', ''):g}"
                                          if isinstance((res or {}).get("valor"),
                                                        (int, float))
                                          else str((res or {}).get("valor", "")),
                                          fg=AZUL, bg=colores_fondo.get(nivel, "#F4F7FB"))
                self.lbl_res_banda.config(
                    text=(res or {}).get("etiqueta", "") or "Sin banda",
                    fg=ROJO_AVISO if nivel == "critico" else TEXTO_OSCURO,
                    bg=colores_fondo.get(nivel, "#F4F7FB"))
                self.lbl_res_msg.config(text=msg + nota_perfil,
                                        bg=colores_fondo.get(nivel, "#F4F7FB"))
            else:
                self.lbl_res_valor.config(text="—", fg=ROJO_AVISO, bg="#F8E8E8")
                self.lbl_res_banda.config(text="No puntua", fg=ROJO_AVISO,
                                          bg="#F8E8E8")
                self.lbl_res_msg.config(text=msg + nota_perfil, bg="#F8E8E8")
                self._ultimo = None
                return
            self._ultimo = {"id": self._calc_actual, "entradas": entradas,
                            "resultado": res, "mensaje": msg,
                            "perfil": perfil}

        def _guardar_nota(self):
            if not getattr(self, "_ultimo", None):
                messagebox.showinfo("Aviso", "Calcula primero un resultado.",
                                    parent=self)
                return
            if not self.app.ruta_activa:
                messagebox.showinfo("Aviso",
                                    "Selecciona una paciente para guardar en su nota.",
                                    parent=self)
                return
            try:
                h = cargar_historia(self.app.ruta_activa)
                registro = {"fecha": fecha_hoy(), "hora": hora_ahora(),
                            "episodio": (self.app.datos_activos or {}).get(
                                "episodio_activo", ""),
                            **self._ultimo}
                h.setdefault("calculos", []).append(registro)
                _escribir_historia(self.app.ruta_activa, h)
                bitacora(f"Calculadora guardada: {self._ultimo['id']} -> "
                         f"{self._ultimo['resultado'].get('etiqueta', '')}",
                         autor=autor_actual())
                self.app.toast("Calculo guardado en la nota ✓")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar: {e}",
                                     parent=self)


    class LineaTiempo(tk.Toplevel):
        """C-22/ADR-0037: linea de tiempo clinica navegable (1 clic desde cualquier pantalla)."""

        def __init__(self, parent_app):
            super().__init__(parent_app)
            self.app = parent_app
            self.title("Linea de tiempo clinica")
            self.configure(bg=CREMA)
            marco = tk.Frame(self, bg=CREMA_CLARO, padx=12, pady=10)
            marco.pack(fill="both", expand=True)
            d = self.app.datos_activos or {}
            tk.Label(marco, text=f"Linea de tiempo — {d.get('nombre_completo','?')}",
                     font=(FUENTE_SISTEMA, 13, "bold"), bg=CREMA_CLARO, fg=AZUL
                     ).pack(anchor="w")
            # filtros por tipo
            fila_f = tk.Frame(marco, bg=CREMA_CLARO)
            fila_f.pack(fill="x", pady=(6, 0))
            self.var_filtro = tk.StringVar(value="todos")
            for txt in ["todos", "control", "alerta", "nota", "documento", "calculo", "consentimiento"]:
                tk.Radiobutton(fila_f, text=txt.capitalize(), variable=self.var_filtro,
                               value=txt, command=self._refrescar, bg=CREMA_CLARO,
                               fg=TEXTO_OSCURO, selectcolor=BLANCO,
                               font=(FUENTE_SISTEMA, 9)).pack(side="left", padx=2)
            self.lista = tk.Listbox(marco, height=16, font=(FUENTE_SISTEMA, 9),
                                    bg=BLANCO, fg=TEXTO_OSCURO, relief="solid", bd=1,
                                    selectbackground=AZUL, selectforeground=TEXTO_CLARO)
            self.lista.pack(fill="both", expand=True, pady=(6, 0))
            self._eventos = []
            self._refrescar()
            auto_tamano(self, parent_app)

        def _refrescar(self):
            self.lista.delete(0, "end")
            eid = (self.app.datos_activos or {}).get("episodio_activo", "")
            evs = eventos_paciente(self.app.ruta_activa, eid)
            f = self.var_filtro.get()
            if f != "todos":
                evs = [e for e in evs if e["tipo"] == f]
            self._eventos = evs
            for e in evs:
                linea = f"{e['fecha']} {e.get('hora','')}  [{e['tipo'].upper()}]  {e['titulo']}"
                if e.get("detalle"):
                    linea += f" — {e['detalle'][:60]}"
                self.lista.insert("end", linea)


    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title(APP_TITLE)
            self.geometry("1080x760")
            self.minsize(880, 620)
            self.configure(bg=CREMA)
            self.ruta_activa = None
            self.datos_activos = None
            self.recientes_rutas = []
            self.busq_rutas = []
            self.tpl_lista = []
            self._after_id = None
            self._cola_respaldo = None
            self._respaldo_activo = False
            self._poblando_tree = False
            self._filtros_activos = {"medico": "", "desde": "", "hasta": "",
                                     "curp": "", "telefono": "", "correo": "",
                                     "ocupacion": "", "alergias": "", "grupo": "",
                                     "eg_min": "", "eg_max": "", "episodios_min": ""}
            sys.excepthook = self._on_error
            self._build()
            self.protocol("WM_DELETE_WINDOW", self._cerrar_seguro)
            self.refresh_home()
            self._binds()
            instalar_contexto(self)
            _zoom = leer_config().get("zoom")
            if _zoom and float(_zoom) != 1.0:
                self.aplicar_zoom(_zoom)
            if not leer_config().get("bienvenida"):
                self.after(400, self._bienvenida)
            # A8: chequeo diario de respaldo programado
            self.after(60000, self._chequear_respaldo_programado)

        def _cerrar_seguro(self):
            """D11: no cerrar la app mientras un respaldo esta en curso."""
            if self._respaldo_activo:
                messagebox.showwarning(
                    "Respaldo en curso",
                    "Espera a que termine el respaldo antes de cerrar el sistema.")
                return
            self.destroy()

        def destroy(self):
            """P7 (T9c): salida limpia -> elimina el lock de instancia."""
            limpiar_bloqueo()
            super().destroy()

        def _bienvenida(self):
            """D13: bienvenida de primera ejecucion apuntando a _PLANTILLAS."""
            try:
                guardar_config({"bienvenida": True})
            except Exception:
                return
            messagebox.showinfo(
                "Bienvenida",
                "Sistema de Expedientes Maternidad listo.\n\n"
                "1. Coloca tus plantillas (.docx, .xlsx, .txt) en la carpeta _PLANTILLAS.\n"
                "2. Registra tu primera paciente (Ctrl+N).\n"
                "3. Genera documentos desde la pestana Documentos.\n\n"
                "Todo funciona sin internet; tus datos viven en esta carpeta.",
                parent=self)

        def _on_error(self, tipo, valor, tb):
            log_exc(valor)
            try:
                messagebox.showerror("Error inesperado", f"Ocurrio un error:\n\n{valor}")
                self.status.set("Error registrado en _ERRORES.log")
            except Exception:
                pass

        # ---------- CONSTRUCCION ----------

        def _build(self):
            # ADR-0034: store reactivo; los 7 refresh manuales se suscriben aqui
            self.estado = Estado()
            self._metricas = {"teclas": 0, "recalculos": 0, "ultimo_recalc_ms": 0.0}
            self.estado.suscribir("recientes", self.refrescar_recientes)
            self.estado.suscribir("etiquetas", self._actualizar_etiqueta)
            self.estado.suscribir("episodios", self._actualizar_episodios)
            self.estado.suscribir("admin", self._actualizar_admin_sel)
            self.estado.suscribir("home", self.refresh_home)
            self.estado.suscribir("hoy", self._refrescar_hoy)
            # Estética global (ADR-0022): estilos ttk unificados
            try:
                self._scaling_base = float(self.tk.call("tk", "scaling"))
            except Exception:
                self._scaling_base = 1.0
            try:
                estilo = ttk.Style(self)
                estilo.configure("Treeview", rowheight=26,
                                 font=(FUENTE_SISTEMA, FUENTE_S))
                estilo.configure("Treeview.Heading",
                                 font=(FUENTE_SISTEMA, FUENTE_S, "bold"))
                estilo.map("Treeview",
                           background=[("selected", AZUL)],
                           foreground=[("selected", TEXTO_CLARO)])
                estilo.configure("Treeview", fieldbackground=BLANCO)
                estilo.map("TCombobox",
                           fieldbackground=[("readonly", BLANCO)])
                estilo.configure("TCombobox", padding=ESP_XS)
                estilo.configure("TProgressbar", troughcolor=GRIS_REJILLA)
            except Exception:
                pass
            self.option_add("*Listbox.selectBackground", AZUL)
            self.option_add("*Listbox.selectForeground", TEXTO_CLARO)
            self.option_add("*Entry.selectBackground", AZUL)
            self.option_add("*Entry.selectForeground", TEXTO_CLARO)
            self.option_add("*Text.selectBackground", AZUL)
            self.option_add("*Text.selectForeground", TEXTO_CLARO)
            self.grid_columnconfigure(1, weight=1)
            self.grid_rowconfigure(0, weight=1)
            # ===== SIDEBAR v2.4 (ADR-0026): marca, contexto, píldoras y pie =====
            sidebar = tk.Frame(self, bg=NAV_BG, width=215)
            sidebar.grid(row=0, column=0, sticky="nsew")
            sidebar.grid_propagate(False)
            tk.Label(sidebar, text="MATERNIDAD", font=(FUENTE_SISTEMA, 17, "bold"),
                     bg=NAV_BG, fg=BLANCO).pack(pady=(16, 0))
            tk.Frame(sidebar, bg=ACENTO, height=3).pack(fill="x", padx=18, pady=(4, 2))
            tk.Label(sidebar, text=f"Expedientes · v{VERSION}", font=(FUENTE_SISTEMA, 8),
                     bg=NAV_BG, fg=NAV_PIE).pack(anchor="w", padx=18, pady=(0, 8))
            # Tarjeta de contexto: el medico siempre ve con quién trabaja
            tarjeta = tk.Frame(sidebar, bg=NAV_TARJETA, highlightthickness=1,
                               highlightbackground="#1A3F6A")
            tarjeta.pack(fill="x", padx=10, pady=(0, 6))
            tk.Label(tarjeta, text="EN CONTEXTO", font=(FUENTE_SISTEMA, 7, "bold"),
                     bg=NAV_TARJETA, fg=ACENTO, anchor="w"
                     ).pack(fill="x", padx=10, pady=(6, 0))
            self.nav_ctx_nombre = tk.Label(
                tarjeta, text="Sin paciente seleccionada",
                font=(FUENTE_SISTEMA, 10, "bold"), bg=NAV_TARJETA, fg=NAV_TEXTO,
                anchor="w", wraplength=180, justify="left")
            self.nav_ctx_nombre.pack(fill="x", padx=10)
            self.nav_ctx_epi = tk.Label(
                tarjeta, text="", font=(FUENTE_SISTEMA, 8), bg=NAV_TARJETA,
                fg=NAV_LEYENDA, anchor="w", wraplength=180)
            self.nav_ctx_epi.pack(fill="x", padx=10, pady=(0, 8))
            self.lbl_sidebar_ctx = self.nav_ctx_nombre  # alias (smoke/contratos)

            def _sep(txt):
                tk.Label(sidebar, text=txt, bg=NAV_BG, fg=NAV_LEYENDA,
                         font=(FUENTE_SISTEMA, 8, "bold"), anchor="w", padx=16
                         ).pack(fill="x", pady=(10, 2))
                tk.Frame(sidebar, bg="#16385F", height=1).pack(fill="x", padx=16)

            # Nav: workflow del medico (ADR-0023) + píldora de página activa
            self._nav_actual = ""
            self.nav_paginas = {"register": "Registrar paciente",
                                "home": "Pacientes",
                                "docs": "Documentos", "graficos": "Graficos",
                                "calculadoras": "Calculadoras",
                                "admin": "Administracion"}
            self.nav = {}
            self.nav_barras = {}
            grupos = [
                (None, [
                    ("Registrar paciente", "✚", lambda: self.show_page("register")),
                ]),
                ("PACIENTES", [
                    ("Pacientes", "☰", lambda: (self.show_page("home"),
                                                self.entry_busqueda_home.focus_set())),
                ]),
                ("ATENCION", [
                    ("Documentos", "▤", lambda: self.show_page("docs")),
                    ("Graficos", "▦", lambda: self.show_page("graficos")),
                    ("Calculadoras", "∑", lambda: self.show_page("calculadoras")),
                ]),
                ("SISTEMA", [
                    ("Administracion", "⚙", lambda: self.show_page("admin")),
                    ("Respaldar", "⇩", self.dialogo_respaldo),
                    ("Salir", "✕", self.destroy),
                ]),
            ]
            for grupo, items in grupos:
                if grupo:
                    _sep(grupo)
                for texto, glifo, accion in items:
                    cont = tk.Frame(sidebar, bg=NAV_BG)
                    cont.pack(fill="x", padx=8, pady=2)
                    barra = tk.Frame(cont, bg=NAV_BG, width=4)
                    barra.pack(side="left", fill="y")
                    b = tk.Button(cont, text=f"  {glifo}  {texto}", command=accion,
                                  relief="flat", bg=NAV_BG, fg=NAV_TEXTO,
                                  activebackground=NAV_TARJETA, activeforeground=BLANCO,
                                  font=(FUENTE_SISTEMA, 11), anchor="w", padx=10,
                                  pady=6, cursor="hand2")
                    b.pack(side="left", fill="x", expand=True, padx=(6, 0))
                    b.bind("<Enter>", lambda e, bb=b, t=texto: bb.config(
                        bg=BLANCO if t == self._nav_actual else NAV_TARJETA))
                    b.bind("<Leave>", lambda e, bb=b, t=texto: bb.config(
                        bg=BLANCO if t == self._nav_actual else NAV_BG))
                    self.nav[texto] = b
                    self.nav_barras[texto] = barra
                    if texto == "Salir":
                        b.config(fg=ROJO_SALIR, activeforeground=ROJO_SALIR)
                    tips = {
                        "Registrar paciente": "Alta de paciente (nueva pagina de inicio)",
                        "Pacientes": "Listado con tabla, filtros y busqueda (Ctrl+F)",
                        "Documentos": "Generar documentos del dia (Ctrl+D)",
                        "Graficos": "Peso, quimica y controles del embarazo",
                        "Calculadoras": "Estratificacion de riesgo (C-19)",
                        "Administracion": "Carpetas, exportaciones, bitacora, zoom",
                        "Respaldar": "Copiar, verificar y restaurar datos",
                        "Salir": "Cierra el sistema",
                    }
                    if texto in tips:
                        Tooltip(b, tips[texto])
            tk.Frame(sidebar, bg="#16385F", height=1).pack(
                side="bottom", fill="x", padx=14, pady=(8, 2))
            tk.Label(sidebar, text="Datos 100% locales · sin internet",
                     bg=NAV_BG, fg=NAV_PIE, font=(FUENTE_SISTEMA, 8), anchor="w"
                     ).pack(side="bottom", fill="x", padx=16, pady=(0, 12))
            self.lbl_nav_autor = tk.Label(sidebar, text=f"Autor: {autor_actual()}",
                                          bg=NAV_BG, fg=NAV_PIE, font=(FUENTE_SISTEMA, 8),
                                          anchor="w")
            self.lbl_nav_autor.pack(side="bottom", fill="x", padx=16)

            main = tk.Frame(self, bg=CREMA)
            main.grid(row=0, column=1, sticky="nsew")
            main.grid_rowconfigure(1, weight=1)
            main.grid_columnconfigure(0, weight=1)
            hdr = tk.Frame(main, bg=CREMA)
            hdr.grid(row=0, column=0, sticky="we", padx=14, pady=(10, 4))
            hdr.grid_columnconfigure(1, weight=1)
            col_tit = tk.Frame(hdr, bg=CREMA)
            col_tit.grid(row=0, column=0, sticky="w")
            self.header = tk.Label(col_tit, text="Inicio", font=(FUENTE_SISTEMA, 18, "bold"),
                                   bg=CREMA, fg=TEXTO_OSCURO)
            self.header.pack(anchor="w")
            # Subtítulo por pantalla/pestaña: identidad de cada contrato
            self.header_sub = tk.Label(col_tit, text="", font=(FUENTE_SISTEMA, 9),
                                       bg=CREMA, fg=GRIS_TEXTO)
            self.header_sub.pack(anchor="w")
            # Acciones del encabezado: la acción primaria vive arriba a la derecha
            self.header_acc = tk.Frame(hdr, bg=CREMA)
            self.header_acc.grid(row=0, column=2, sticky="e")
            self.lbl_autor = tk.Label(hdr, text=f"Autor: {autor_actual()}",
                                      font=(FUENTE_SISTEMA, 9), bg=CREMA, fg=GRIS_SECUNDARIO)
            self.lbl_autor.grid(row=0, column=1, sticky="ne")
            self.pages_box = tk.Frame(main, bg=CREMA_CLARO)
            self.pages_box.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 8))
            self.pages_box.grid_rowconfigure(0, weight=1)
            self.pages_box.grid_columnconfigure(0, weight=1)
            self.status = tk.StringVar(value="Listo")
            tk.Label(self, textvariable=self.status, anchor="w", bg=AZUL,
                     fg=TEXTO_CLARO, font=(FUENTE_SISTEMA, 9), height=1
                     ).grid(row=1, column=0, columnspan=2, sticky="we")
            self.progreso = ttk.Progressbar(self, mode="indeterminate", length=240)
            self.progreso.grid(row=2, column=0, columnspan=2, sticky="we", padx=14, pady=3)
            self.progreso.grid_remove()

            self.pages = {}
            self.pages["home"] = self._pagina_home()
            self.pages["register"] = self._pagina_registro()
            self.pages["docs"] = self._pagina_documentos()
            self.pages["graficos"] = self._pagina_graficos()
            self.pages["calculadoras"] = self._pagina_calculadoras()
            self.pages["admin"] = self._pagina_admin()
            self.show_page("register")

        def _binds(self):
            self.bind("<Control-n>", lambda e: self.show_page("register"))
            self.bind("<Control-d>", lambda e: self.show_page("docs"))
            self.bind("<Control-f>", lambda e: (self.show_page("home"),
                                                self.entry_busqueda_home.focus_set()))
            self.bind("<Control-h>", lambda e: self._historia_rapida())
            self.bind("<F1>", lambda e: self._ayuda())
            # ADR-0027: acciones de un tecla por pantalla
            self.bind("<F2>", lambda e: self._generar_docs())
            self.bind("<F3>", lambda e: self._abrir_ficha(self.ruta_activa))
            self.bind("<F4>", lambda e: self._dlg_nom004())
            self.bind("<F5>", lambda e: self._abrir_controles())

        def _ayuda(self):
            messagebox.showinfo(
                "Ayuda",
                "ATAJOS:\n"
                "  Ctrl+N  Registro\n  Ctrl+F  Buscar\n"
                "  Ctrl+D  Documentos\n  Ctrl+O  Abrir expediente\n"
                "  Ctrl+H  Historia clinica (con paciente seleccionada)\n"
                "  F2  Generar documento   F3  Ficha\n"
                "  F4  Historia clinica    F5  Controles prenatales\n"
                "  F1  Esta ayuda\n\n"
                "REGISTRO: Enter avanza entre campos y guarda en el ultimo.\n"
                "El boton [Cal] abre el calendario: anio por teclado o lista,\n"
                "mes por nombre, dia en la rejilla o con 'Seleccionar fecha'.\n"
                "El mes actual no permite dias futuros.\n\n"
                "EMBARAZO: un embarazo = un episodio (boton 'Nuevo embarazo').\n"
                "La FUM nace preseleccionada con HOY; cambiala si la real es otra\n"
                "y veras FPP y edad gestacional calculadas al instante.\n"
                "'Datos del embarazo' edita FUM y fecha de referencia.\n"
                "Progreso Quimico es un archivo vivo POR EPISODIO.\n\n"
                "DOCUMENTOS: plantillas en _PLANTILLAS con marcadores\n"
                "{{PACIENTE}} {{FUM}} {{FPP}} {{EG}} {{DOCTOR}}...\n"
                "Ficha: ver/editar datos de la paciente y sus episodios.\n\n"
                "GRAFICOS: dibuja las columnas numericas de tu\n"
                "PROGRESO_QUIMICO.xlsx del episodio (peso, quimica, ...).\n\n"
                "RESPALDO: Verificar (SHA-256) y Restaurar.\n"
                "Contraseña opcional: cifra el respaldo (Fernet+PBKDF2)\n"
                "si esta instalada cryptography; si no, solo integridad.\n"
                "PIN de apertura y monousuario: en Administracion.\n"
                "Busca por nombre, CURP, telefono o correo.\n"
                "Soporte normativo, no certificacion oficial."
            )

        def show_page(self, nombre):
            for n, fr in self.pages.items():
                if n == nombre:
                    fr.grid(row=0, column=0, sticky="nsew")
                    fr.tkraise()
                else:
                    fr.grid_remove()
            titulos = {"home": "Pacientes", "register": "Registrar paciente",
                       "docs": "Documentos", "graficos": "Graficos de progreso",
                       "calculadoras": "Calculadoras medicas",
                       "admin": "Administracion"}
            # ADR-0026: cada pantalla tiene su propio subtítulo (identidad)
            subs = {
                "home": "Listado de pacientes · ¿quién viene hoy?",
                "register": "Alta de paciente en 2 pasos",
                "docs": "El taller · documentos del día con la paciente en contexto",
                "graficos": "Progreso del embarazo · peso, química y controles",
                "calculadoras": "Estratificación de riesgo · catálogo del médico (C-19)",
                "admin": "Sistema · carpetas, exportaciones, bitácora y respaldos",
            }
            self.header.config(text=titulos.get(nombre, nombre))
            self.header_sub.config(text=subs.get(nombre, ""))
            # Píldora de la página activa en el sidebar (barra dorada + negrita)
            self._nav_actual = self.nav_paginas.get(nombre, "")
            self._nav_repintar()
            self._header_acciones(nombre)
            if nombre == "docs":
                self.refrescar_recientes()
                self._actualizar_episodios()
                self._subtitulo_pestana()
                # U131/v4.17: selector abierto solo si no hay paciente en contexto
                self._docs_selector(self.ruta_activa is None)
            if nombre == "home":
                self.refresh_home()
            if nombre == "graficos":
                self._dibujar_graficos()
            if nombre == "calculadoras":
                self.calc_panel.refrescar_ctx()

        def _nav_repintar(self):
            for texto, b in self.nav.items():
                activo = texto == self._nav_actual
                b.config(bg=BLANCO if activo else NAV_BG,
                         fg=NAV_BG if activo else (ROJO_SALIR if texto == "Salir"
                                                   else NAV_TEXTO),
                         font=(FUENTE_SISTEMA, 11, "bold" if activo else "normal"))
                barra = self.nav_barras.get(texto)
                if barra is not None:
                    barra.config(bg=ACENTO if activo else NAV_BG)

        def _header_acciones(self, nombre):
            """Acción primaria del encabezado por pantalla (ADR-0026)."""
            for w in self.header_acc.winfo_children():
                w.destroy()
            tabla = {
                "home": [("Nueva paciente", lambda: self.show_page("register"), True),
                         ("Actualizar", self.refresh_home, False)],
                "graficos": [("Refrescar", self._dibujar_graficos, True)],
                "calculadoras": [("Ver pacientes", lambda: self.show_page("home"), False)],
                "admin": [("Respaldar", self.dialogo_respaldo, True)],
                "register": [("Ver pacientes", lambda: self.show_page("home"), False)],
                "docs": [("Ayuda (F1)", self._ayuda, False)],
            }
            for txt, cmd, primario in tabla.get(nombre, []):
                if primario:
                    tk.Button(self.header_acc, text=f"▸ {txt}", command=cmd, bg=AZUL,
                              fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                              font=(FUENTE_SISTEMA, 10, "bold"), padx=12, pady=6,
                              cursor="hand2").pack(side="right", padx=4)
                else:
                    tk.Button(self.header_acc, text=txt, command=cmd, bg=BLANCO,
                              fg=AZUL, activebackground="#E8EEF6", relief="flat",
                              highlightthickness=1, highlightbackground=GRIS_BORDE,
                              font=(FUENTE_SISTEMA, 10), padx=12, pady=6,
                              cursor="hand2").pack(side="right", padx=4)

        def _subtitulo_pestana(self):
            """El subtítulo del encabezado sigue a la pestaña activa de Documentos."""
            try:
                idx = self.tabs.index(self.tabs.select())
            except Exception:
                return
            subs = [
                "Con paciente · generar desde plantillas con la paciente en contexto",
                "Sin paciente · documentos generales del consultorio",
                "Plantillas · cargar/guardar .txt y crear nuevas",
            ]
            self.header_sub.config(text=subs[idx] if 0 <= idx < len(subs) else "")

        # ---------- SELECCION DE PACIENTE (fuente unica de verdad: ruta) ----------

        def _on_state_change(self, dirty):
            """ADR-0034: punto unico de notificacion del store."""
            self.estado.dispatch("estado:cambiado", dirty)
            self.estado.flush()

        def seleccionar_paciente(self, ruta, origen=None):
            try:
                datos = cargar_paciente(ruta)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el expediente:\n{e}")
                return
            self.ruta_activa = ruta
            self.datos_activos = datos
            # ADR-0034: un solo dispatch reemplaza los 7 refresh manuales
            self.estado.dispatch("paciente:cambiada", {
                "recientes", "etiquetas", "episodios", "admin", "home", "hoy"})
            self.estado.flush()
            self._resaltar(ruta, origen)
            if origen != "home":
                self._resaltar_en_tree(ruta)
            registrar_lectura(ruta, "seleccion")  # A5 hito
            self.status.set(f"Seleccionada: {datos.get('nombre_completo', '')}")
            self._docs_selector(False)  # U131/v4.17: elegir paciente pliega el selector

        def _resaltar_en_tree(self, ruta):
            """Sincroniza la seleccion del tree de Inicio con ruta_activa (R2)."""
            if not hasattr(self, "tree"):
                return
            self._poblando_tree = True
            try:
                hijos = self.tree.get_children()
                idx = self.rutas_tree.index(ruta) if ruta in self.rutas_tree else None
                if idx is not None and idx < len(hijos):
                    self.tree.selection_set(hijos[idx])
                else:
                    self.tree.selection_remove(self.tree.selection())
            finally:
                self._poblando_tree = False

        def _actualizar_etiqueta(self):
            if self.ruta_activa and self.datos_activos:
                txt = f"Paciente seleccionada: {self.datos_activos.get('nombre_completo', '')}"
                color = VERDE_OK
                epi = self.datos_activos.get("episodio_activo", "")
                if epi:
                    txt += f"   |   Episodio: {epi}"
            else:
                txt = "Ninguna paciente seleccionada"
                color = ROJO_AVISO
            self.lbl_sel.config(text=txt, fg=color)
            self._refrescar_ctx_home()
            self._refrescar_tira()
            self._actualizar_ctx_sidebar()

        # ---------- SIDEBAR: CONTEXTO PERSISTENTE (ADR-0023/v1.3) ----------

        def _actualizar_ctx_sidebar(self):
            if not hasattr(self, "nav_ctx_nombre"):
                return
            if self.ruta_activa and self.datos_activos:
                self.nav_ctx_nombre.config(
                    text=self.datos_activos.get("nombre_completo", ""), fg=BLANCO)
                self.nav_ctx_epi.config(
                    text=self.datos_activos.get("episodio_activo", ""))
            else:
                self.nav_ctx_nombre.config(text="Sin paciente seleccionada",
                                           fg=NAV_TEXTO)
                self.nav_ctx_epi.config(text="")

        # ---------- U-PROP-6: TIRA FUM|FPP|EG (C-05, T6b) ----------

        def _refrescar_tira(self):
            if not hasattr(self, "lbl_tira"):
                return
            d = self.datos_activos if self.ruta_activa else None
            if not d:
                self.lbl_tira.config(text="Resumen del embarazo: (sin paciente)",
                                     fg=GRIS_TEXTO)
                return
            epi = episodio_contexto(d)
            fum = (epi or {}).get("fum") or ""
            if not fum:
                self.lbl_tira.config(
                    text="FUM: (sin FUM)  |  FPP: -  |  EG: -  (captura la FUM en "
                         "el episodio para ver FPP y EG)", fg=GRIS_TERCIARIO)
                return
            fpp = calcular_fpp(fum)
            eg = calcular_eg(fum)
            txt = (f"FUM {fecha_legible(fum)}  |  "
                   f"FPP {fecha_legible(fpp) if fpp else '-'}  |  "
                   f"EG {eg if eg else '- (FUM futura)'}"
                   + self._sufijo_gestacional(fum))
            # ADR-0038: tendencia de controles en la tira
            try:
                tend = tendencias_controles(self.ruta_activa)
                if tend.get("n_visitas"):
                    partes = [f"Tendencia ({tend['n_visitas']} visitas)"]
                    if tend.get("delta_peso") is not None:
                        partes.append(f"Δpeso {tend['delta_peso']:+.1f} kg")
                    if tend.get("delta_ta_s") is not None:
                        partes.append(f"ΔTA {tend['delta_ta_s']:+.0f}")
                    txt += "\n" + " · ".join(partes)
            except Exception:
                pass
            self.lbl_tira.config(text=txt, fg=AZUL)

        def _sufijo_gestacional(self, fum):
            """ADR-0027: 'trimestre II  ·  faltan 34 dias' (tira y panel)."""
            if not fum:
                return ""
            fpp = calcular_fpp(fum)
            partes = []
            tri = trimestre_desde_eg(calcular_eg(fum))
            if tri:
                partes.append(f"trimestre {tri}")
            dpf = dias_para_fpp(fpp) if fpp else None
            if dpf is not None:
                partes.append(f"faltan {dpf} dias" if dpf >= 0
                              else f"FPP hace {-dpf} dias")
            return "   ·   " + "   ·   ".join(partes) if partes else ""

        def _resaltar(self, ruta, origen):
            idx_r = self.recientes_rutas.index(ruta) if ruta in self.recientes_rutas else None
            idx_b = self.busq_rutas.index(ruta) if ruta in self.busq_rutas else None
            if idx_r is not None and origen != "busqueda":
                self.lb_recientes.selection_clear(0, "end")
                self.lb_recientes.selection_set(idx_r)
                self.lb_recientes.see(idx_r)
            elif idx_r is None:
                self.lb_recientes.selection_clear(0, "end")
            if idx_b is not None and origen == "busqueda":
                self.lb_busqueda.selection_clear(0, "end")
                self.lb_busqueda.selection_set(idx_b)
                self.lb_busqueda.see(idx_b)
            elif idx_b is None:
                self.lb_busqueda.selection_clear(0, "end")

        def refrescar_recientes(self):
            self.lb_recientes.delete(0, "end")
            self.recientes_rutas = []
            for clave, ruta in listar_recientes(10):
                self.lb_recientes.insert("end", nombre_mostrar(ruta))
                self.recientes_rutas.append(ruta)
            if self.ruta_activa in self.recientes_rutas:
                i = self.recientes_rutas.index(self.ruta_activa)
                self.lb_recientes.selection_set(i)
                self.lb_recientes.see(i)
            else:
                self.lb_recientes.selection_clear(0, "end")
            self._actualizar_etiqueta()

        def ir_a_documentos(self):
            self.show_page("docs")
            try:
                self.tabs.select(0)
            except Exception:
                pass

        # ---------- PAGINA: INICIO ----------

        def _pagina_home(self):
            home = tk.Frame(self.pages_box, bg=CREMA_CLARO)
            home.grid_rowconfigure(5, weight=1)
            home.grid_columnconfigure(0, weight=1)

            # Panel "Hoy" (ADR-0027/T13)
            self.frame_hoy = tk.Frame(home, bg=BLANCO, highlightthickness=1,
                                      highlightbackground=GRIS_BORDE)
            self.frame_hoy.grid(row=1, column=0, sticky="we", padx=16, pady=(0, 4))
            self.lbl_hoy = tk.Label(self.frame_hoy, text="HOY", bg=BLANCO, fg=AZUL,
                                    font=(FUENTE_SISTEMA, 9, "bold"), anchor="w", padx=10)
            self.lbl_hoy.pack(fill="x", padx=6, pady=(6, 0))
            self.caja_hoy = tk.Frame(self.frame_hoy, bg=BLANCO)
            self.caja_hoy.pack(fill="x", padx=8, pady=(2, 8))

            tk.Label(home, text="Busca por nombre, CURP, telefono o correo",
                     font=(FUENTE_SISTEMA, 10), bg=CREMA_CLARO,
                     fg=GRIS_TEXTO).grid(row=2, column=0, sticky="w", padx=18)
            marco_busq = tk.Frame(home, bg=CREMA_CLARO)
            marco_busq.grid(row=3, column=0, sticky="we", padx=16, pady=8)
            marco_busq.grid_columnconfigure(0, weight=1)
            self.search_var = tk.StringVar()
            self.entry_busqueda_home = tk.Entry(marco_busq, textvariable=self.search_var,
                                                font=(FUENTE_SISTEMA, 12), bg=BLANCO,
                                                fg=TEXTO_OSCURO, relief="solid", bd=1)
            self.entry_busqueda_home.grid(row=0, column=0, sticky="we", padx=(0, 6))
            self.search_var.trace_add("write", self._programar_busqueda)
            tk.Button(marco_busq, text="Buscar", command=self._ejecutar_busqueda,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 10)).grid(row=0, column=1, padx=3)
            tk.Button(marco_busq, text="Filtros", command=self._dialogo_filtros,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 10)).grid(row=0, column=2, padx=3)
            self.fila_chips = tk.Frame(home, bg=CREMA_CLARO)
            self.fila_chips.grid(row=4, column=0, sticky="we", padx=16, pady=(0, 6))
            self.fila_chips.grid_remove()

            # ===== ZONA PRINCIPAL: tabla (izq) + columna de acciones (der) =====
            zona = tk.Frame(home, bg=CREMA_CLARO)
            zona.grid(row=5, column=0, sticky="nsew", padx=16, pady=4)
            zona.grid_columnconfigure(0, weight=1)
            zona.grid_rowconfigure(0, weight=1)

            marco_arbol = tk.Frame(zona, bg=BLANCO, highlightthickness=1,
                                   highlightbackground=GRIS_BORDE)
            marco_arbol.grid(row=0, column=0, sticky="nsew")
            marco_arbol.grid_rowconfigure(0, weight=1)
            marco_arbol.grid_columnconfigure(0, weight=1)
            self._orden_col = "ult"
            self._orden_desc = True
            cols = ("nombre", "fecha", "curp", "tel", "ocup", "alerg",
                    "eg", "epi", "ult")
            self.tree = ttk.Treeview(marco_arbol, columns=cols, show="headings",
                                     height=10)
            _cols_def = [
                ("nombre", "Nombre completo", 180, True),
                ("fecha", "FdN", 90, False),
                ("curp", "CURP", 110, False),
                ("tel", "Telefono", 80, False),
                ("ocup", "Ocupacion", 100, False),
                ("alerg", "Alergias", 70, False),
                ("eg", "EG", 75, False),
                ("epi", "Epis", 45, False),
                ("ult", "Ult.mov", 95, False),
            ]
            for col, texto, ancho, stretch in _cols_def:
                self.tree.heading(col, text=texto,
                                  command=lambda c=col: self._ordenar_por(c))
                self.tree.column(col, width=ancho, stretch=stretch)
            self.tree.grid(row=0, column=0, sticky="nsew")
            scr_v = ttk.Scrollbar(marco_arbol, orient="vertical",
                                  command=self.tree.yview)
            scr_v.grid(row=0, column=1, sticky="ns")
            scr_h = ttk.Scrollbar(marco_arbol, orient="horizontal",
                                  command=self.tree.xview)
            scr_h.grid(row=1, column=0, sticky="we")
            self.tree.configure(yscrollcommand=scr_v.set, xscrollcommand=scr_h.set)
            self.tree.bind("<Double-Button-1>", lambda e: self._abrir_de_tree())
            self.tree.bind("<Return>", lambda e: self._abrir_de_tree())
            self.tree.bind("<<TreeviewSelect>>", self._sel_tree_home)
            self.rutas_tree = []

            # ===== COLUMNA DE ACCIONES A LA DERECHA DE LA TABLA =====
            acc_col = tk.Frame(zona, bg=BLANCO, highlightthickness=1,
                               highlightbackground=GRIS_BORDE, width=150)
            acc_col.grid(row=0, column=1, sticky="ns", padx=(6, 0))
            acc_col.grid_propagate(False)
            self.lbl_ctx_home = tk.Label(
                acc_col, text="Sin paciente",
                font=(FUENTE_SISTEMA, 9, "bold"), bg=BLANCO, fg=ROJO_AVISO,
                anchor="w", wraplength=140, justify="left", padx=8)
            self.lbl_ctx_home.pack(fill="x", pady=(8, 4))
            tk.Frame(acc_col, bg=GRIS_BORDE, height=1).pack(fill="x", padx=6, pady=2)
            for txt, cmd, prim in [
                    ("Generar", self._ctx_generar, True),
                    ("Ficha", self._ctx_ficha, False),
                    ("Historia A-M", self._ctx_historia_am, False),
                    ("Progreso", self._ctx_progreso, False),
                    ("Consentim.", self._ctx_consentimiento, False),
                    ("Resumen", self._ctx_resumen, False),
                    ("Graficos", lambda: self.show_page("graficos"), False),
                    ("Timeline", self._abrir_timeline, False)]:
                tk.Button(acc_col, text=txt, command=cmd,
                          bg=AZUL if prim else GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                          activebackground=AZUL_HOVER if prim else GRIS_APAGADO,
                          font=(FUENTE_SISTEMA, 9, "bold" if prim else "normal"),
                          anchor="w", padx=8, pady=3).pack(fill="x", padx=6, pady=2)

            self.frame_vacio_home = tk.Frame(home, bg=CREMA_CLARO)
            self.frame_vacio_home.grid(row=6, column=0, sticky="we", padx=16, pady=4)
            tk.Label(self.frame_vacio_home, text="Aun no hay pacientes registradas.",
                     font=(FUENTE_SISTEMA, 12), bg=CREMA_CLARO, fg=TEXTO_OSCURO
                     ).pack(anchor="w")
            tk.Button(self.frame_vacio_home, text=" Registrar primera paciente",
                      command=lambda: self.show_page("register"),
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 10, "bold"), padx=10, pady=4
                      ).pack(anchor="w", pady=4)
            self.frame_vacio_home.grid_remove()
            marco_btns = tk.Frame(home, bg=CREMA_CLARO)
            marco_btns.grid(row=7, column=0, sticky="we", padx=16, pady=10)
            tips_home = {
                "Abrir expediente": "Abre la Ficha (expediente en el software)",
                "Limpiar filtros": "Quita filtros y chips, lista completa",
            }
            for txt, cmd in [
                ("Abrir expediente", self._abrir_de_tree),
                ("Limpiar filtros", self._limpiar_filtros),
            ]:
                bb = tk.Button(marco_btns, text=txt, command=cmd, bg=GRIS_SECUNDARIO,
                               fg=TEXTO_CLARO, activebackground=GRIS_APAGADO,
                               font=(FUENTE_SISTEMA, 10), padx=12, pady=BTN_PADY)
                bb.pack(side="left", padx=4)
                if txt in tips_home:
                    Tooltip(bb, tips_home[txt])
            return home

        # ---------- UX-P1: PANEL DE CONTEXTO (U-PROP-1) ----------

        def _ordenar_por(self, col):
            """Ordena el tree por columna (asc/desc) — ADR-0026."""
            if self._orden_col == col:
                self._orden_desc = not self._orden_desc
            else:
                self._orden_col = col
                self._orden_desc = True
            self._ejecutar_busqueda()

        def _ctx_ruta(self):
            return self._ruta_de_tree() or self.ruta_activa

        def _ctx_generar(self):
            ruta = self._ctx_ruta()
            if not ruta:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            self.seleccionar_paciente(ruta, origen="home")
            DocGenerator(self, ruta)

        def _ctx_ficha(self):
            self._abrir_ficha(self._ctx_ruta())

        def _abrir_timeline(self):
            """C-22/ADR-0037: linea de tiempo en 1 clic."""
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            LineaTiempo(self)

        def _ctx_historia_am(self):
            """C-18: historia completa A-M (checklist rutina de vida) en 1 clic."""
            ruta = self._ctx_ruta()
            if not ruta:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            self.seleccionar_paciente(ruta, origen="home")
            HistoriaCompleta(self)

        def _ctx_progreso(self):
            """Progreso Quimico del episodio en 1 clic desde el panel."""
            ruta = self._ctx_ruta()
            if not ruta:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            self.seleccionar_paciente(ruta, origen="home")
            ok, msg, salida = progreso_quimico(self.ruta_activa)
            if ok:
                self.toast("Progreso Quimico ✓")
                ir_a_ruta(salida)
            else:
                messagebox.showerror("Error", msg)

        def _ctx_consentimiento(self):
            """Consentimiento informado en 1 clic desde el panel."""
            ruta = self._ctx_ruta()
            if not ruta:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            self.seleccionar_paciente(ruta, origen="home")
            self._dlg_consentimiento()

        def _refrescar_ctx_home(self):
            if not hasattr(self, "lbl_ctx_home"):
                return
            d = self.datos_activos if self.ruta_activa else None
            if not d:
                self.lbl_ctx_home.config(
                    text="Ninguna paciente en contexto — selecciona una fila de la lista",
                    fg=ROJO_AVISO)
                return
            epi = episodio_contexto(d)
            eid = d.get("episodio_activo") or (epi or {}).get("id", "") or "(sin episodio)"
            fum = (epi or {}).get("fum") or ""
            eg = calcular_eg(fum) if fum else None
            txt = f"{d.get('nombre_completo', '')}   |   Episodio: {eid}"
            if not fum:
                txt += "   |   EG: (sin FUM)"
            elif eg:
                txt += f"   |   EG: {eg}"
            else:
                txt += "   |   EG: - (FUM futura)"
            txt += self._sufijo_gestacional(fum)
            # ADR-0038: recordatorio de termino y tendencia de controles
            try:
                rec = recordatorios_eg(fum)
                if rec.get("dias_termino") is not None:
                    dt = rec["dias_termino"]
                    if dt > 0:
                        txt += f"   ·   Término en {dt} d"
                    elif rec.get("dias_postermino", 0) > 0:
                        txt += f"   ·   +{abs(dt)} d post-término"
            except Exception:
                pass
            try:
                tend = tendencias_controles(self.ruta_activa, eid)
                if tend.get("peso_kg_semana") is not None:
                    txt += f"   ·   Δpeso {tend['peso_kg_semana']:+.2f} kg/sem"
            except Exception:
                pass
            try:
                pend = alertas_pendientes(self.ruta_activa, eid)
                if pend:
                    txt += f"   ·   ⚠ {len(pend)} alerta(s) pendiente(s)"
            except Exception:
                pass
            self.lbl_ctx_home.config(text=txt, fg=VERDE_OK)

        def _refrescar_hoy(self):
            """Panel Hoy (ADR-0027/T13): controles recientes y umbral neutro."""
            if not hasattr(self, "caja_hoy"):
                return
            for w in self.caja_hoy.winfo_children():
                w.destroy()
            try:
                r = resumen_hoy()
            except Exception:
                r = {"hoy": fecha_hoy(), "umbral": DIAS_SIN_CONTROL_AVISO,
                     "recientes": [], "vencidos": []}
            self.lbl_hoy.config(text=f"HOY · {fecha_legible(r['hoy'])}")
            col_izq = tk.Frame(self.caja_hoy, bg=BLANCO)
            col_izq.pack(side="left", fill="both", expand=True, padx=(0, 6))
            col_der = tk.Frame(self.caja_hoy, bg=BLANCO)
            col_der.pack(side="left", fill="both", expand=True, padx=(6, 0))

            def _bloque(col, titulo, items, izq):
                tk.Label(col, text=titulo, bg=BLANCO, fg=GRIS_TEXTO,
                         font=(FUENTE_SISTEMA, 9, "bold"), anchor="w"
                         ).pack(fill="x", pady=(0, 2))
                if not items:
                    tk.Label(col, text="(ninguno)", bg=BLANCO, fg=GRIS_APAGADO,
                             font=(FUENTE_SISTEMA, 9), anchor="w").pack(fill="x")
                    return
                for nombre, ruta, ult, delta in items[:5]:
                    if izq:
                        txt = f"● {nombre} · {fecha_legible(ult)}"
                        b = tk.Button(col, text=txt, command=lambda rr=ruta:
                                      self.seleccionar_paciente(rr, origen="home"),
                                      bg=BLANCO, fg=AZUL, relief="flat",
                                      activebackground="#E8EEF6", anchor="w",
                                      font=(FUENTE_SISTEMA, 9), cursor="hand2")
                    else:
                        txt = f"● {nombre} · hace {delta} d (ult. {fecha_legible(ult)})"
                        b = tk.Label(col, text=txt, bg=BLANCO, fg=ROJO_AVISO,
                                     font=(FUENTE_SISTEMA, 9), anchor="w")
                    b.pack(fill="x")
                if len(items) > 5:
                    tk.Label(col, text=f"+{len(items) - 5} mas…", bg=BLANCO,
                             fg=GRIS_APAGADO, font=(FUENTE_SISTEMA, 8),
                             anchor="w").pack(fill="x")

            _bloque(col_izq, f"CONTROLES DE LOS ULTIMOS 7 DIAS",
                    r["recientes"], izq=True)
            _bloque(col_der,
                    f"SIN CONTROL DESDE HACE {r['umbral']}+ DIAS (umbral configurable)",
                    r["vencidos"], izq=False)

        def _ctx_resumen(self):
            ruta = self._ctx_ruta()
            if not ruta:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            self._abrir_resumen(ruta)

        def _abrir_resumen(self, ruta):
            """Resumen de consulta de 1 pagina (ADR-0027/T13)."""
            try:
                texto = resumen_consulta(ruta)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo armar el resumen:\n{e}")
                return
            win = tk.Toplevel(self)
            win.title("Resumen de consulta")
            win.configure(bg=CREMA)
            cuerpo = tk.Frame(win, bg=CREMA_CLARO, padx=12, pady=10)
            cuerpo.pack(fill="both", expand=True)
            tk.Label(cuerpo, text="Resumen de consulta (1 pagina)",
                     font=(FUENTE_SISTEMA, 12, "bold"), bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).pack(anchor="w")
            caja = tk.Text(cuerpo, width=86, height=26, font=("Courier", 10),
                           bg=BLANCO, fg=TEXTO_OSCURO, relief="solid", bd=1,
                           wrap="none")
            caja.insert("1.0", texto)
            caja.config(state="disabled")
            caja.pack(fill="both", expand=True, pady=6)

            def _archivo():
                try:
                    nombre = (f"RESUMEN_{Path(ruta).name}_{timestamp()}.txt")
                    salida = EXPORTACIONES_DIR / nombre
                    EXPORTACIONES_DIR.mkdir(parents=True, exist_ok=True)
                    salida.write_text(texto, encoding="utf-8")
                    self.toast("Resumen generado ✓")
                    ir_a_ruta(salida)
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo generar: {e}",
                                         parent=win)

            pie = tk.Frame(cuerpo, bg=CREMA_CLARO)
            pie.pack(fill="x")
            tk.Button(pie, text="Generar archivo y abrir", command=_archivo,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 10, "bold"), padx=10, pady=6
                      ).pack(side="left")
            tk.Button(pie, text="Cerrar", command=win.destroy, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO,
                      padx=10, pady=6).pack(side="right")
            auto_tamano(win, self)

        def _vacio_home(self, vacio):
            """UX-P5: CTA de estado vacio visible solo con 0 pacientes totales."""
            if not hasattr(self, "frame_vacio_home"):
                return
            if vacio:
                self.frame_vacio_home.grid()
            else:
                self.frame_vacio_home.grid_remove()

        # ---------- UX-P4: CHIPS DE FILTRO (U-PROP-4, T6b) ----------

        def _refrescar_chips(self):
            if not hasattr(self, "fila_chips"):
                return
            for w in self.fila_chips.winfo_children():
                w.destroy()
            f = self._filtros_activos
            activos = []
            _etq = {"medico": "Médico", "desde": "FdN desde", "hasta": "FdN hasta",
                    "curp": "CURP", "telefono": "Tel", "correo": "Correo",
                    "ocupacion": "Ocupación", "alergias": "Alergias",
                    "grupo": "Grupo", "eg_min": "EG ≥", "eg_max": "EG ≤",
                    "episodios_min": "Epis ≥"}
            for k in f:
                if not f.get(k):
                    continue
                v = f[k]
                if k in ("desde", "hasta"):
                    v = fecha_legible(v)
                activos.append((k, f"{_etq.get(k, k)}: {v}"))
            if not activos:
                self.fila_chips.grid_remove()
                return
            tk.Label(self.fila_chips, text="Filtros:", bg=CREMA_CLARO,
                     fg=GRIS_TEXTO, font=(FUENTE_SISTEMA, FUENTE_XS)
                     ).pack(side="left", padx=(0, ESP_XS))
            for clave, texto in activos:
                chip = tk.Frame(self.fila_chips, bg=AZUL_CHIP, highlightthickness=1,
                                highlightbackground=GRIS_BORDE)
                chip.pack(side="left", padx=(0, ESP_S))
                tk.Label(chip, text=f" {texto} ", bg=AZUL_CHIP, fg=TEXTO_OSCURO,
                         font=(FUENTE_SISTEMA, FUENTE_XS)).pack(side="left")
                tk.Button(chip, text="×", command=lambda k=clave: self._quitar_chip(k),
                          bg=AZUL_CHIP, fg=ROJO_AVISO, activebackground=AZUL_CHIP,
                          activeforeground=ROJO_AVISO, relief="flat", bd=0,
                          font=(FUENTE_SISTEMA, FUENTE_M, "bold"), cursor="hand2"
                          ).pack(side="right", padx=(0, ESP_XS))
            self.fila_chips.grid()

        def _quitar_chip(self, clave):
            """Quitar un chip actualiza _filtros_activos y rebusca (C-02 §8)."""
            self._filtros_activos[clave] = ""
            self._ejecutar_busqueda()

        def _limpiar_filtros(self):
            self._filtros_activos = {"medico": "", "desde": "", "hasta": "",
                                     "curp": "", "telefono": "", "correo": "",
                                     "ocupacion": "", "alergias": "", "grupo": "",
                                     "eg_min": "", "eg_max": "", "episodios_min": ""}
            self._ejecutar_busqueda()
            self.toast("Filtros limpiados ✓")

        def _programar_busqueda(self, *args):
            if self._after_id:
                try:
                    self.after_cancel(self._after_id)
                except Exception:
                    pass
            self._after_id = self.after(300, self._ejecutar_busqueda)

        def _sel_tree_home(self, event=None):
            # R1: sincroniza seleccion Inicio -> ruta_activa sin bucles
            if self._poblando_tree:
                return
            ruta = self._ruta_de_tree()
            if not ruta or ruta == self.ruta_activa:
                return
            self.seleccionar_paciente(ruta, origen="home")

        def _ejecutar_busqueda(self):
            self._after_id = None
            q = self.search_var.get().strip()
            resultados = buscar(q) if q else listar_pacientes()
            f = self._filtros_activos
            hay_filtro = any(f.get(k) for k in f)
            if hay_filtro:
                resultados = filtrar(
                    resultados, medico=f.get("medico", ""),
                    desde=f.get("desde", ""), hasta=f.get("hasta", ""),
                    curp=f.get("curp", ""), telefono=f.get("telefono", ""),
                    correo=f.get("correo", ""), ocupacion=f.get("ocupacion", ""),
                    alergias=f.get("alergias", ""), grupo=f.get("grupo", ""),
                    eg_min=f.get("eg_min", ""), eg_max=f.get("eg_max", ""),
                    episodios_min=f.get("episodios_min", ""))
            self._poblar_tree(resultados)
            self._vacio_home(not listar_pacientes())  # UX-P5: 0 pacientes totales
            activos = "; ".join(
                p for p in (f"{k} {v}" for k, v in f.items() if v) if p)
            self.status.set(f"{len(resultados)} pacientes"
                            + (f"  [Filtros: {activos}]" if activos else ""))
            self._refrescar_chips()  # UX-P4: chips de los filtros activos

        def _poblar_tree(self, resultados):
            self._poblando_tree = True  # R1: suppress <<TreeviewSelect>> during refill
            try:
                self.tree.delete(*self.tree.get_children())
                self.rutas_tree = []
                filas = []
                for nombre, ruta in resultados:
                    try:
                        d = cargar_paciente(ruta)
                        fecha = fecha_legible(d.get("fecha_nacimiento", ""))
                        curp = d.get("curp", "")
                        tel = d.get("telefono", "")
                        ocup = d.get("ocupacion", "") or ""
                        alerg = "Si" if (d.get("alergias") or "").strip() else "No"
                        mostrar = d.get("nombre_completo", nombre)
                        epi = episodio_contexto(d)
                        fum = (epi or {}).get("fum") or ""
                        eg = calcular_eg(fum) if fum else None
                        if eg:
                            eg_txt = eg
                        elif fum:
                            eg_txt = "-"
                        else:
                            eg_txt = "(sin FUM)"
                        try:
                            m = os.path.getmtime(Path(ruta) / "datos.json")
                        except OSError:
                            m = 0.0
                        try:
                            ult = fecha_legible(
                                datetime.datetime.fromtimestamp(m).date().isoformat())
                        except Exception:
                            ult = "—"
                        n_epi = len(d.get("episodios") or [])
                    except Exception:
                        mostrar, fecha, curp, tel = nombre, "", "", ""
                        ocup, alerg = "", ""
                        eg_txt, ult, n_epi, m = "(sin datos)", "—", 0, 0.0
                    valores = (mostrar, fecha, curp, tel, ocup, alerg,
                               eg_txt, n_epi, ult)
                    filas.append((m, valores, ruta))
                # orden por columna activa (default: Ult. movimiento desc)
                _orden = getattr(self, "_orden_col", "ult")
                _desc = getattr(self, "_orden_desc", True)
                _idx = {"nombre": 0, "fecha": 1, "curp": 2, "tel": 3,
                        "ocup": 4, "alerg": 5, "eg": 6, "epi": 7, "ult": 8}.get(_orden, 8)
                if _orden == "ult":
                    filas.sort(key=lambda f: f[0], reverse=_desc)
                else:
                    filas.sort(key=lambda f: str(f[1][_idx]).lower(), reverse=_desc)
                for _, valores, ruta in filas:
                    self.tree.insert("", "end", values=valores)
                    self.rutas_tree.append(ruta)
            finally:
                self._poblando_tree = False

        def refresh_home(self, *args):
            self._ejecutar_busqueda()
            self._refrescar_hoy()

        def _ruta_de_tree(self):
            sel = self.tree.selection()
            if not sel:
                return None
            return self.rutas_tree[self.tree.index(sel[0])]

        def _abrir_de_tree(self):
            """Doble clic / Enter / 'Abrir expediente': abre el expediente EN EL
            SOFTWARE (Ficha C-08). La carpeta del sistema vive en el boton
            'Carpeta' del panel, en la Ficha y en Administracion (2026-10-16)."""
            ruta = self._ruta_de_tree()
            if not ruta:
                messagebox.showinfo("Aviso", "Selecciona un paciente de la lista.")
                return
            self.seleccionar_paciente(ruta, origen="home")
            FichaPaciente(self, ruta)

        def _generar_de_tree(self):
            ruta = self._ruta_de_tree()
            if not ruta:
                messagebox.showinfo("Aviso", "Selecciona un paciente de la lista.")
                return
            self.seleccionar_paciente(ruta)
            DocGenerator(self, ruta)

        def _dialogo_filtros(self):
            win = tk.Toplevel(self)
            win.title("Filtros avanzados")
            win.configure(bg=CREMA)
            win.resizable(False, False)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=14, pady=14)
            marco.pack(fill="both", expand=True, padx=10, pady=10)
            tk.Label(marco, text="Filtros", font=(FUENTE_SISTEMA, 14, "bold"),
                     bg=CREMA_CLARO, fg=TEXTO_OSCURO).grid(row=0, column=0, columnspan=2, pady=(0, 8))
            campos = {}
            vars_campos = {}
            for i, (lab, ph) in enumerate([("Medico remitente:", ""),
                                           ("Nacimiento desde:", "AAAA-MM-DD"),
                                           ("Nacimiento hasta:", "AAAA-MM-DD")], start=1):
                tk.Label(marco, text=lab, bg=CREMA_CLARO, fg=TEXTO_OSCURO
                         ).grid(row=i, column=0, sticky="w", pady=5, padx=(0, 8))
                var = tk.StringVar()
                e = tk.Entry(marco, font=(FUENTE_SISTEMA, 11), width=24, bg=BLANCO,
                             textvariable=var)
                e.grid(row=i, column=1, sticky="we", pady=5)
                campos[lab] = e
                vars_campos[lab] = var
                if ph:
                    boton_calendario(marco, var, "iso", popup_parent=win
                                     ).grid(row=i, column=2, padx=(6, 0), pady=5)

            # prefill con el estado activo
            vars_campos["Medico remitente:"].set(self._filtros_activos["medico"])
            vars_campos["Nacimiento desde:"].set(self._filtros_activos["desde"])
            vars_campos["Nacimiento hasta:"].set(self._filtros_activos["hasta"])

            def aplicar():
                desde_txt = vars_campos["Nacimiento desde:"].get().strip()
                hasta_txt = vars_campos["Nacimiento hasta:"].get().strip()
                desde = parsear_fecha(desde_txt) if desde_txt else ""
                hasta = parsear_fecha(hasta_txt) if hasta_txt else ""
                if (desde_txt and not desde) or (hasta_txt and not hasta):
                    messagebox.showwarning(
                        "Fecha invalida",
                        "Usa AAAA-MM-DD, dd/mm/aaaa o el boton [Cal].", parent=win)
                    return
                self._filtros_activos = {
                    "medico": vars_campos["Medico remitente:"].get().strip(),
                    "desde": desde or "",
                    "hasta": hasta or "",
                }
                self._ejecutar_busqueda()
                win.destroy()

            def limpiar():
                self._filtros_activos = {"medico": "", "desde": "", "hasta": "",
                                     "curp": "", "telefono": "", "correo": "",
                                     "ocupacion": "", "alergias": "", "grupo": "",
                                     "eg_min": "", "eg_max": "", "episodios_min": ""}
                self._ejecutar_busqueda()
                win.destroy()

            tk.Button(marco, text="Limpiar filtros", command=limpiar,
                      bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO, activebackground=GRIS_APAGADO,
                      font=(FUENTE_SISTEMA, 11)).grid(row=4, column=0, pady=14, sticky="w")
            tk.Button(marco, text="Aplicar", command=aplicar, bg=AZUL, fg=TEXTO_CLARO,
                      activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 11)
                      ).grid(row=4, column=1, pady=14, sticky="e")
            win.bind("<Escape>", lambda ev: win.destroy())
            auto_tamano(win, self)

        # ---------- PAGINA: GRAFICOS DE PROGRESO (C-17, ADR-0023) ----------

        def _pagina_graficos(self):
            """ADR-0045/v6.0: 3 zonas — contexto (tira+fuentes) · stat-cards ·
            mini-gráficas; barra de acciones con Refrescar primario."""
            pag = tk.Frame(self.pages_box, bg=CREMA_CLARO)
            pag.grid_columnconfigure(0, weight=1)
            pag.grid_rowconfigure(2, weight=1)
            ctx_card = tk.Frame(pag, bg=BLANCO, highlightthickness=1,
                                highlightbackground=GRIS_BORDE)
            ctx_card.grid(row=0, column=0, sticky="we", padx=16, pady=(8, 2))
            self.lbl_graf_ctx = tk.Label(ctx_card, text="Ninguna paciente en contexto",
                                         font=(FUENTE_SISTEMA, 11, "bold"), bg=BLANCO,
                                         fg=ROJO_AVISO, anchor="w", padx=10)
            self.lbl_graf_ctx.pack(fill="x", pady=(7, 2))
            # U134: tira FUM|FPP|EG (mismo dato de cabecera que Inicio/Documentos)
            self.lbl_graf_tira = tk.Label(ctx_card, text="", font=(FUENTE_SISTEMA, 9),
                                          bg=BLANCO, fg=GRIS_TEXTO, anchor="w",
                                          padx=10)
            self.lbl_graf_tira.pack(fill="x", pady=(0, 2))
            # U133: línea de fuentes honesta (se calcula DESPUÉS de leer)
            self.lbl_graf_fuentes = tk.Label(ctx_card, text="", font=(FUENTE_SISTEMA, 9, "bold"),
                                             bg=BLANCO, fg=GRIS_TEXTO, anchor="w",
                                             padx=10)
            self.lbl_graf_fuentes.pack(fill="x", pady=(0, 7))
            acciones = tk.Frame(pag, bg=CREMA_CLARO)
            acciones.grid(row=1, column=0, sticky="we", padx=16, pady=6)
            for txt, cmd, prim in [
                    ("Refrescar", self._graf_refrescar, True),
                    ("Abrir/crear Progreso Quimico", self._graf_progreso, False),
                    ("Controles", self._graf_registrar_control, False),
                    ("Ver pacientes", lambda: (self.show_page("home"),
                                               self.entry_busqueda_home.focus_set()),
                     False)]:
                if prim:
                    tk.Button(acciones, text=f"▸ {txt}", command=cmd, bg=AZUL,
                              fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                              font=(FUENTE_SISTEMA, 10, "bold"), padx=10,
                              pady=BTN_PADY).pack(side="left", padx=4)
                else:
                    tk.Button(acciones, text=txt, command=cmd, bg=BLANCO, fg=AZUL,
                              activebackground="#E8EEF6", relief="flat",
                              highlightthickness=1, highlightbackground=GRIS_BORDE,
                              font=(FUENTE_SISTEMA, 10), padx=10,
                              pady=BTN_PADY).pack(side="left", padx=4)
            marco = tk.Frame(pag, bg=BLANCO, highlightthickness=1,
                             highlightbackground=GRIS_BORDE)
            marco.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 10))
            marco.grid_rowconfigure(0, weight=1)
            marco.grid_columnconfigure(0, weight=1)
            self.canvas_graf = tk.Canvas(marco, bg=BLANCO, highlightthickness=0)
            scr = ttk.Scrollbar(marco, orient="vertical", command=self.canvas_graf.yview)
            self.canvas_graf.configure(yscrollcommand=scr.set)
            self.canvas_graf.grid(row=0, column=0, sticky="nsew")
            scr.grid(row=0, column=1, sticky="ns")
            self._graf_caja = tk.Frame(self.canvas_graf, bg=BLANCO)
            self._graf_vent = self.canvas_graf.create_window(
                (0, 0), window=self._graf_caja, anchor="nw")
            self._graf_caja.bind(
                "<Configure>",
                lambda e: self.canvas_graf.configure(scrollregion=self.canvas_graf.bbox("all")))
            self.canvas_graf.bind("<Configure>", self._graf_ancho_cambio)
            # U141: rueda del ratón sobre el área de gráficas (además de la barra)
            self.canvas_graf.bind(
                "<MouseWheel>",
                lambda e: self.canvas_graf.yview_scroll(-1 * (e.delta // 120), "units"))
            self.canvas_graf.bind(
                "<Button-4>", lambda e: self.canvas_graf.yview_scroll(-1, "units"))
            self.canvas_graf.bind(
                "<Button-5>", lambda e: self.canvas_graf.yview_scroll(1, "units"))
            self._graf_ancho_prev = 0
            return pag

        def _graf_ancho_cambio(self, e):
            """U137/v6.0: las gráficas se dibujan al ancho disponible; el
            redibujo es debounced y solo si el ancho cambió de verdad (R1)."""
            self.canvas_graf.itemconfigure(self._graf_vent, width=e.width)
            if abs(e.width - getattr(self, "_graf_ancho_prev", 0)) <= 24:
                return
            self._graf_ancho_prev = e.width
            if getattr(self, "_graf_after_ancho", None):
                try:
                    self.after_cancel(self._graf_after_ancho)
                except Exception:
                    pass
            self._graf_after_ancho = self.after(300, self._dibujar_graficos)

        def _pagina_calculadoras(self):
            """ADR-0042/v5.0: C-19 embebida en el area principal (sin ventana)."""
            pag = tk.Frame(self.pages_box, bg=CREMA_CLARO)
            pag.grid_rowconfigure(0, weight=1)
            pag.grid_columnconfigure(0, weight=1)
            self.calc_panel = Calculadoras(self, container=pag)
            self.calc_panel.grid(row=0, column=0, sticky="nsew")
            return pag

        def _graf_progreso(self):
            """U135/v6.0: abre (o crea en 1.er uso) el Progreso Quimico con el
            sistema — antes el boton decia 'abriendo' y no abria."""
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            ok, msg, salida = progreso_quimico(self.ruta_activa)
            (messagebox.showinfo if ok else messagebox.showerror)("Progreso Quimico", msg)
            if ok:
                if salida:
                    ir_a_ruta(salida)
                self._dibujar_graficos()

        def _graf_refrescar(self):
            """U143/v6.0: relee el Excel Y recarga datos_activos del disco
            (un control registrado en C-16 aparece sin reiniciar)."""
            if self.ruta_activa:
                try:
                    self.datos_activos = cargar_paciente(self.ruta_activa)
                except Exception:
                    pass
            self._dibujar_graficos()

        def _graf_registrar_control(self):
            """U136/v6.0: abre C-16 (Controles) sin pasar por Documentos; al
            cerrarla la pagina se repinta sola."""
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            win = ControlesPrenatales(self)
            win.wait_window()
            self._dibujar_graficos()

        @staticmethod
        def _clave_serie(titulo):
            """Clave de casado entre fuentes (U138): 'Peso (control)' == 'Peso'.

            El sufijo se quita ANTES de normalizar (normalizar convierte
            '(control)' en '_CONTROL_' y el paréntesis ya no existiría)."""
            t = str(titulo or "")
            t = re.sub(r"\(control\)", " ", t, flags=re.I)
            t = normalizar(t).lower()
            t = re.sub(r"[^a-z0-9]+", " ", t)
            return t.strip()

        def _graf_tira(self):
            """U134/v6.0: tira FUM|FPP|EG + trimestre (dato de cabecera U-PROP-6)."""
            d = self.datos_activos if self.ruta_activa else None
            epi = episodio_contexto(d) if d else None
            fum = (epi or {}).get("fum") or ""
            if not fum:
                self.lbl_graf_tira.config(
                    text="FUM: (sin FUM)  |  FPP: -  |  EG: -  (captura la FUM en el "
                         "episodio para ver FPP y EG)", fg=GRIS_TERCIARIO)
                return
            fpp = calcular_fpp(fum)
            eg = calcular_eg(fum)
            self.lbl_graf_tira.config(
                text=(f"FUM {fecha_legible(fum)}  |  "
                      f"FPP {fecha_legible(fpp) if fpp else '-'}  |  "
                      f"EG {eg if eg else '- (FUM futura)'}"
                      + self._sufijo_gestacional(fum)),
                fg=GRIS_TEXTO)

        def _dibujar_graficos(self):
            """C-17 v2.0 (ADR-0045): contexto honesto + stat-cards + series con
            eje temporal y superposición de fuentes. Solo visualiza lo que el
            médico ya registró; SIN interpretación clínica (ADR-0013)."""
            if not hasattr(self, "_graf_caja"):
                return
            for w in self._graf_caja.winfo_children():
                w.destroy()
            d = self.datos_activos if self.ruta_activa else None
            if not d:
                self.lbl_graf_ctx.config(
                    text="Ninguna paciente en contexto — selecciona una en Inicio",
                    fg=ROJO_AVISO)
                self.lbl_graf_tira.config(text="")
                self.lbl_graf_fuentes.config(text="")
                self._graf_mensaje(
                    "Sin paciente en contexto.",
                    (("Ver pacientes", lambda: self.show_page("home")),))
                return
            epi = episodio_contexto(d)
            eid = d.get("episodio_activo") or (epi or {}).get("id", "")
            self.lbl_graf_ctx.config(
                text=(f"{d.get('nombre_completo', '')}   |   Episodio: "
                      f"{eid or '(sin episodio)'}"),
                fg=VERDE_OK)
            self._graf_tira()
            if not eid:
                self.lbl_graf_fuentes.config(text="")
                self._graf_mensaje(
                    "Esta paciente no tiene episodio activo.",
                    (("Nuevo embarazo", self._nuevo_episodio),))
                return
            # --- lectura de ambas fuentes (la etiqueta se calcula DESPUÉS) ---
            ruta_xlsx = Path(self.ruta_activa) / eid / "PROGRESO_QUIMICO.xlsx"
            series_excel, estado_xlsx = [], "ausente"
            if ruta_xlsx.exists():
                ok, msg, enc, filas = leer_progreso_quimico(ruta_xlsx)
                if ok:
                    series_excel = series_progreso(enc, filas)
                    estado_xlsx = "ok"
                else:
                    estado_xlsx = f"ilegible: {msg}"
            ctrls = listar_controles(self.ruta_activa, eid)
            series_ctrl = []
            for campo, titulo in (("peso", "Peso (control)"),
                                  ("ta_s", "TA sistolica (control)"),
                                  ("ta_d", "TA diastolica (control)"),
                                  ("au", "Altura uterina (control)"),
                                  ("fc", "FC (control)"),
                                  ("fcf", "FCF (control)")):
                pts = [(str(c.get("fecha") or ""), c.get(campo)) for c in reversed(ctrls)
                       if isinstance(c.get(campo), (int, float))]
                if pts:
                    series_ctrl.append((titulo, pts))
            # --- U133: cabecera de fuentes honesta (leída, no fijada) ---
            if estado_xlsx == "ok":
                self.lbl_graf_fuentes.config(
                    text=(f"Progreso Químico: {len(series_excel)} columnas  ·  "
                          f"Controles: {len(ctrls)} visita(s)"), fg=VERDE_OK)
            elif estado_xlsx == "ausente":
                self.lbl_graf_fuentes.config(
                    text=(f"Progreso Químico: sin archivo  ·  "
                          f"Controles: {len(ctrls)} visita(s)"), fg="#A67C00")
            else:
                self.lbl_graf_fuentes.config(
                    text=f"Progreso Químico {estado_xlsx}  ·  Controles: {len(ctrls)} visita(s)",
                    fg=ROJO_AVISO)
            # --- U138: merge por clave; superpuestas van en la sección del médico ---
            ctrl_por_clave = {self._clave_serie(t): pts for t, pts in series_ctrl}
            usados = set()
            entradas = []          # {"titulo", "lineas": [(color, pts, fuente_txt)], "badge"}
            for titulo, pts in series_excel:
                k = self._clave_serie(titulo)
                lineas = [(AZUL, pts, "Excel del médico")]
                if k in ctrl_por_clave:
                    lineas.append((VERDE_OK, ctrl_por_clave[k], "Controles"))
                    usados.add(k)
                entradas.append({"titulo": titulo, "lineas": lineas,
                                 "badge": "ambas fuentes" if len(lineas) > 1
                                          else "Excel del médico",
                                 "grupo": "excel"})
            for titulo, pts in series_ctrl:
                if self._clave_serie(titulo) in usados:
                    continue
                entradas.append({"titulo": titulo,
                                 "lineas": [(VERDE_OK, pts, "Controles")],
                                 "badge": "Controles", "grupo": "control"})
            # --- U142: estados vacíos accionables ---
            if not entradas:
                if estado_xlsx == "ausente":
                    self._graf_mensaje(
                        "Este episodio aún no tiene mediciones.\n"
                        "Registra mediciones numéricas en el Progreso Químico "
                        "(peso, glucosa, ...) o visitas con Controles (peso, TA, AU, FCF).",
                        (("Abrir/crear Progreso Quimico", self._graf_progreso),
                         ("Registrar control aqui", self._graf_registrar_control)))
                elif isinstance(estado_xlsx, str) and estado_xlsx.startswith("ilegible"):
                    self._graf_mensaje(
                        f"Progreso Químico ilegible: {estado_xlsx.split(': ', 1)[1]}",
                        (("Abrir/crear Progreso Quimico", self._graf_progreso),))
                elif estado_xlsx == "ok" and not series_excel:
                    self._graf_mensaje(
                        "Progreso Químico creado: registra mediciones en tu hoja "
                        "y pulsa Refrescar.\nTu hoja aún no tiene columnas numéricas.",
                        (("Abrir Progreso Quimico", self._graf_progreso),
                         ("Registrar control aqui", self._graf_registrar_control)))
                else:
                    self._graf_mensaje(
                        "Tu hoja aún no tiene columnas numéricas.",
                        (("Abrir Progreso Quimico", self._graf_progreso),))
                return
            # --- U144: tope visible con pie explícito (nunca corte mudo) ---
            visibles = entradas[:GRAF_MAX_SERIES]
            ocultas = len(entradas) - len(visibles)
            # --- U139: stat-cards en grid con wrap (3 por fila) ---
            cards = []
            for e in visibles:
                todos = [p for _, pts, _ in e["lineas"] for p in pts]
                st = stats_serie(todos)
                if st:
                    cards.append((e, st))
            cards = cards[:GRAF_MAX_CARDS]
            if cards:
                grid = tk.Frame(self._graf_caja, bg=BLANCO)
                grid.pack(fill="x", padx=16, pady=(6, 2))
                for i, (e, st) in enumerate(cards):
                    valor, _et, dias_hace, delta, minimo, maximo, media = st
                    r, c = divmod(i, 3)
                    grid.grid_rowconfigure(r, weight=1)
                    grid.grid_columnconfigure(c, weight=1, minsize=180)
                    badge_col = {"Excel del médico": AZUL, "Controles": VERDE_OK,
                                 "ambas fuentes": ACENTO}[e["badge"]]
                    card = tk.Frame(grid, bg="#F4F7FB", highlightthickness=1,
                                    highlightbackground=GRIS_BORDE)
                    card.grid(row=r, column=c, sticky="nsew", padx=(0, 8), pady=4)
                    fila_t = tk.Frame(card, bg="#F4F7FB")
                    fila_t.pack(fill="x", padx=8, pady=(6, 0))
                    tk.Label(fila_t, text=e["titulo"], font=(FUENTE_SISTEMA, 8, "bold"),
                             bg="#F4F7FB", fg=GRIS_TEXTO, anchor="w"
                             ).pack(side="left", fill="x", expand=True)
                    tk.Label(fila_t, text=e["badge"], font=(FUENTE_SISTEMA, 7, "bold"),
                             bg="#F4F7FB", fg=badge_col, anchor="e"
                             ).pack(side="right")
                    tk.Label(card, text=f"{valor:g}" if isinstance(valor, (int, float))
                             else str(valor), font=(FUENTE_SISTEMA, 15, "bold"),
                             bg="#F4F7FB", fg=AZUL, anchor="w"
                             ).pack(fill="x", padx=8)
                    rango = []
                    if minimo is not None and maximo is not None:
                        rango.append(f"min {minimo:g} · max {maximo:g}")
                    if media is not None:
                        rango.append(f"media {media:g}")
                    if rango:
                        tk.Label(card, text=" · ".join(rango), font=(FUENTE_SISTEMA, 8),
                                 bg="#F4F7FB", fg=GRIS_TERCIARIO, anchor="w"
                                 ).pack(fill="x", padx=8)
                    meta = []
                    if dias_hace is not None:
                        meta.append(f"hace {dias_hace} d" if dias_hace >= 0
                                    else f"en {-dias_hace} d")
                    if delta is not None:
                        meta.append(f"Δ {delta:+g}")
                    if not meta:
                        meta.append(f"{len([1 for _, pts, _ in e['lineas'] for _ in pts])}"
                                    " mediciones")
                    tk.Label(card, text=" · ".join(meta), font=(FUENTE_SISTEMA, 8),
                             bg="#F4F7FB", fg=GRIS_TEXTO, anchor="w"
                             ).pack(fill="x", padx=8, pady=(0, 6))
                # U138: leyenda de fuentes (si hay mezcla)
                fuentes_vis = {e["badge"] for e in visibles}
                if "ambas fuentes" in fuentes_vis or (
                        "Excel del médico" in fuentes_vis and "Controles" in fuentes_vis):
                    ley = tk.Frame(self._graf_caja, bg=BLANCO)
                    ley.pack(fill="x", padx=16, pady=(0, 2))
                    for col, txt in ((AZUL, "— Excel del médico"),
                                     (VERDE_OK, "— Controles")):
                        tk.Label(ley, text=txt, bg=BLANCO, fg=col,
                                 font=(FUENTE_SISTEMA, 8, "bold")).pack(side="left", padx=(0, 12))
            # --- U144: secciones por fuente con encabezado ---
            for grupo, texto in (
                    ("excel", "Series del Progreso Químico (tu hoja)"),
                    ("control", "Series de controles (visitas registradas)")):
                sub = [e for e in visibles if e["grupo"] == grupo]
                if not sub:
                    continue
                tk.Frame(self._graf_caja, bg=GRIS_BORDE, height=1).pack(
                    fill="x", padx=16, pady=(8, 2))
                tk.Label(self._graf_caja, text=texto, bg=BLANCO, fg=AZUL,
                         font=(FUENTE_SISTEMA, 9, "bold"), anchor="w"
                         ).pack(fill="x", padx=16)
                for e in sub:
                    self._graf_mini(e["titulo"], e["lineas"], e["badge"])
            if ocultas > 0:
                tk.Label(self._graf_caja,
                         text=f"+{ocultas} series no mostradas — "
                              f"GRAF_MAX_SERIES ({GRAF_MAX_SERIES}) en app.py",
                         bg=BLANCO, fg=GRIS_TERCIARIO, font=(FUENTE_SISTEMA, 8),
                         anchor="w").pack(fill="x", padx=16, pady=(0, 8))

        def _graf_mensaje(self, texto, ctas=()):
            """Estado vacío accionable (U142): mensaje + CTA inline."""
            tk.Label(self._graf_caja, text=texto, bg=BLANCO, fg=GRIS_TEXTO,
                     font=(FUENTE_SISTEMA, 11), justify="left", wraplength=560,
                     padx=16).pack(anchor="w", pady=(16, 4))
            if ctas:
                fila = tk.Frame(self._graf_caja, bg=BLANCO)
                fila.pack(anchor="w", padx=16, pady=(0, 12))
                for txt, cmd in ctas:
                    tk.Button(fila, text=txt, command=cmd, bg=AZUL, fg=TEXTO_CLARO,
                              activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 10),
                              padx=10, pady=BTN_PADY).pack(side="left", padx=(0, 8))

        def _graf_mini(self, titulo, lineas, badge=""):
            """Mini-gráfica v2.0 (ADR-0045): eje X proporcional al tiempo con
            corte de huecos (B95), rejilla min/media/max, superposición de
            fuentes con leyenda y hover de medición. Nunca altera los datos."""
            frame = tk.Frame(self._graf_caja, bg=BLANCO)
            frame.pack(fill="x", padx=16, pady=6)
            cab = tk.Frame(frame, bg=BLANCO)
            cab.pack(fill="x")
            tk.Label(cab, text=f"{titulo}   {badge and '[' + badge + ']'}",
                     bg=BLANCO, fg=TEXTO_OSCURO,
                     font=(FUENTE_SISTEMA, FUENTE_M, "bold")).pack(anchor="w")
            # --- proyección temporal compartida por TODAS las líneas (B95) ---
            bases = []
            for _, pts, _ in lineas:
                for et, _v in pts:
                    iso = parsear_fecha(str(et)) if et is not None else None
                    if iso:
                        bases.append(datetime.date.fromisoformat(iso))
            temporal = bool(bases) and all(
                all(parsear_fecha(str(et)) is not None for et, _ in pts)
                for _, pts, _ in lineas)
            proys = []  # [(color, [[(x, v, et)...]...], fuente_txt)]
            if temporal:
                glo = min(bases)
                dias_max = 0
                for color, pts, fuente in lineas:
                    # base de la LÍNEA (no del segmento): B95 da días
                    # relativos al mínimo de su propia serie
                    f_lin = min(datetime.date.fromisoformat(f)
                                for f in (parsear_fecha(str(et)) for et, _ in pts)
                                if f)
                    off = (f_lin - glo).days
                    _, segs = segmentos_temporales(pts)
                    reb = [[(x + off, v, et) for x, v, et in seg] for seg in segs]
                    for seg in reb:
                        for x, _v, _et in seg:
                            dias_max = max(dias_max, x)
                    proys.append((color, reb, fuente))
                span = max(dias_max, 1)
            else:
                # modo índice para TODA la gráfica (nunca mezclar unidades):
                # si alguna etiqueta no parsea, todas las líneas van por índice
                for color, pts, fuente in lineas:
                    seg = [(float(i), v, et) for i, (et, v) in enumerate(pts)]
                    proys.append((color, [seg], fuente))
                span = max([x for _, segs, _ in proys for seg in segs
                            for x, _v, _et in seg] or [1])
                span = max(span, 1)
            # --- lienzo al ancho disponible (U137) ---
            ancho_disp = int(self.canvas_graf.winfo_width() or 0)
            W = max(560, ancho_disp - 48) if ancho_disp > 100 else 640
            H = 165
            lienzo = tk.Canvas(frame, width=W, height=H, bg=BLANCO,
                               highlightthickness=1, highlightbackground=GRIS_BORDE)
            lienzo.pack(anchor="w", pady=2, fill="x")
            izq, der, arr, abj = 52, 14, 14, 26
            ancho, alto = W - izq - der, H - arr - abj
            todos = [v for _, segs, _ in proys for seg in segs for _x, v, _et in seg]
            minv, maxv = min(todos), max(todos)
            if maxv <= minv:
                maxv = minv + 1
            medv = sum(todos) / len(todos)

            def y(v):
                return arr + alto - (v - minv) / (maxv - minv) * alto

            def x(xval):
                return izq + xval / span * ancho
            # rejilla min · media · max (U139: el eje dice lo que importa)
            for v in (minv, medv, maxv):
                yy = y(v)
                lienzo.create_line(izq, yy, W - der, yy, fill="#E9EDF3",
                                   dash=(2, 3))
                lienzo.create_text(izq - 6, yy, text=f"{v:g}", anchor="e",
                                   fill=GRIS_TERCIARIO, font=(FUENTE_SISTEMA, 8))
            lienzo.create_line(izq, arr, izq, H - abj, fill=GRIS_BORDE)
            lienzo.create_line(izq, H - abj, W - der, H - abj, fill=GRIS_BORDE)
            # --- series: segmentos (el hueco > umbral queda como línea ausente) ---
            puntos_hover = []  # (px, py, etiqueta, valor)
            for color, segs, _fuente in proys:
                for seg in segs:
                    coords = []
                    for xval, v, et in seg:
                        px, py = x(xval), y(v)
                        coords += [px, py]
                        puntos_hover.append((px, py, str(et), v))
                    if len(coords) >= 4:
                        lienzo.create_line(*coords, fill=color, width=2)
                    for px, py, _et, _v in [
                            (x(xv), y(v), et, v) for xv, v, et in seg]:
                        lienzo.create_oval(px - 2.5, py - 2.5, px + 2.5, py + 2.5,
                                           fill=BLANCO, outline=color, width=2)
            # ejes X: primera · intermedia · última etiqueta (si parsean)
            etiquetas_x = [(xval, et) for _c, segs, _f in proys
                           for seg in segs for xval, _v, et in seg]
            etiquetas_x.sort(key=lambda t: t[0])
            marcas = []
            if etiquetas_x:
                marcas.append(etiquetas_x[0])
                if len(etiquetas_x) >= 3:
                    marcas.append(etiquetas_x[len(etiquetas_x) // 2])
                if len(etiquetas_x) >= 2:
                    marcas.append(etiquetas_x[-1])
            for xval, et in marcas:
                iso = parsear_fecha(str(et))
                txt = fecha_legible(iso) if iso else str(et)[:10]
                lienzo.create_text(x(xval), H - abj + 12, text=txt, anchor="n",
                                   fill=GRIS_TEXTO, font=(FUENTE_SISTEMA, 8))
            lienzo._puntos = puntos_hover
            lienzo.bind("<Motion>", lambda ev, cv=lienzo: self._graf_hover(cv, ev))
            lienzo.bind("<Leave>", lambda ev, cv=lienzo: cv.delete("hov"))

        def _graf_hover(self, cv, ev):
            """U140/v6.0: al acercar el cursor se resalta la medición cercana
            con 'fecha · valor' (lectura; sin clic)."""
            pts = getattr(cv, "_puntos", None) or []
            if not pts:
                return
            mejor, dm2 = None, 16 ** 2
            for px, py, et, v in pts:
                d = (px - ev.x) ** 2 + (py - ev.y) ** 2
                if d <= dm2:
                    dm2, mejor = d, (px, py, et, v)
            cv.delete("hov")
            if not mejor:
                return
            px, py, et, v = mejor
            iso = parsear_fecha(et)
            txt_et = fecha_legible(iso) if iso else et[:10]
            val = f"{v:g}" if isinstance(v, (int, float)) else str(v)
            cv.create_oval(px - 5, py - 5, px + 5, py + 5, outline=TEXTO_OSCURO,
                           width=2, tags="hov")
            cv.create_text(px, py - 12, text=f"{txt_et} · {val}", anchor="s",
                           fill=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 8, "bold"),
                           tags="hov")

        # ---------- PAGINA: REGISTRO ----------

        def _pagina_registro(self):
            # v4.16: scroll vertical (bloque de signos vitales + 2 pasos)
            # v4.17: reg_pagina con weights -> el scroll llena TODO el ancho
            self.reg_pagina = tk.Frame(self.pages_box, bg=CREMA_CLARO)
            self.reg_pagina.grid_rowconfigure(0, weight=1)
            self.reg_pagina.grid_columnconfigure(0, weight=1)
            self.reg_frame, _cont = envoltorio_scroll(self.reg_pagina, CREMA_CLARO)
            _cont.grid(row=0, column=0, sticky="nsew")
            self.reg_vars = {k: tk.StringVar() for k in
                             ("ap_p", "ap_m", "nombres", "fecha", "curp",
                              "telefono", "correo", "medico", "notas",
                              # signos vitales y medidas opcionales (alimentan calculadoras)
                              "peso", "talla", "ta_s", "ta_d", "fc", "fr",
                              "temp", "spo2", "cintura", "cadera",
                              "per_abdominal", "per_cefalico")}
            self.reg_ruta = None
            self._paso1()
            return self.reg_pagina

        def _limpiar_registro(self, claves):
            for k in claves:
                self.reg_vars[k].set("")

        def _marcar_error(self, entry, duracion=900):
            """D9: resalta un campo invalido en rojo y revierte tras 900 ms."""
            try:
                original = entry.cget("bg")
                entry.config(bg="#F5B7B1")
                entry.after(duracion, lambda: entry.config(bg=original))
            except Exception:
                pass

        def toast(self, msg, ms=2200):
            """Confirmacion flotante (ADR-0027): no interrumpe como un dialogo."""
            try:
                tw = tk.Toplevel(self)
                tw.wm_overrideredirect(True)
                try:
                    tw.attributes("-topmost", True)
                except Exception:
                    pass
                x = self.winfo_rootx() + self.winfo_width() - 300
                y = self.winfo_rooty() + self.winfo_height() - 110
                tw.wm_geometry(f"+{max(x, 10)}+{max(y, 10)}")
                tk.Label(tw, text=msg, bg=VERDE_OK, fg=TEXTO_CLARO,
                         font=(FUENTE_SISTEMA, 10, "bold"), padx=14, pady=8
                         ).pack()
                tw.after(ms, tw.destroy)
            except Exception:
                pass

        def aplicar_zoom(self, factor):
            """Zoom de interfaz 100/125/150 % (ADR-0027): tk scaling, persistido."""
            try:
                factor = float(factor)
                self.tk.call("tk", "scaling", self._scaling_base * factor)
                guardar_config({"zoom": factor})
                pct = int(round(factor * 100))
                self.status.set(f"Zoom {pct} % aplicado")
                self.toast(f"Zoom {pct} %")
            except Exception as e:
                self.status.set(f"No se pudo aplicar el zoom: {e}")

        def _paso1(self):
            for w in self.reg_frame.winfo_children():
                w.destroy()
            _badge = tk.Frame(self.reg_frame, bg=AZUL)
            _badge.grid(row=0, column=0, sticky="w", padx=30, pady=(14, 2))
            tk.Label(_badge, text="PASO 1 DE 2 · IDENTIDAD", bg=AZUL, fg=TEXTO_CLARO,
                     font=(FUENTE_SISTEMA, 10, "bold"), padx=10, pady=4).pack()
            tk.Label(self.reg_frame,
                     text="Apellido paterno, materno, nombre(s) y fecha. Enter avanza; en el ultimo campo guarda.",
                     font=(FUENTE_SISTEMA, 10), bg=CREMA_CLARO, fg=GRIS_TEXTO
                     ).grid(row=1, column=0, pady=(0, 8))
            campos = tk.Frame(self.reg_frame, bg=CREMA_CLARO)
            campos.grid(row=2, column=0, sticky="we", padx=30)
            campos.grid_columnconfigure(1, weight=1)
            filas = [("Apellido paterno", "ap_p", True),
                     ("Apellido materno", "ap_m", True),
                     ("Nombre(s)", "nombres", True),
                     ("Fecha nacimiento", "fecha", False),
                     ("CURP (opcional)", "curp", True)]
            self.reg_entries = []
            self._paso1_entries = {}
            for i, (lab, key, upper) in enumerate(filas):
                tk.Label(campos, text=lab, anchor="w", bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                         font=(FUENTE_SISTEMA, 11)).grid(row=i, column=0, sticky="w",
                                                     pady=5, padx=(0, 10))
                e = tk.Entry(campos, textvariable=self.reg_vars[key], font=(FUENTE_SISTEMA, 12),
                             bg=BLANCO, fg=TEXTO_OSCURO, relief="solid", bd=1)
                e.grid(row=i, column=1, sticky="we", pady=5)
                self.reg_entries.append(e)
                self._paso1_entries[key] = e
                if upper:
                    e.bind("<KeyRelease>",
                           lambda ev, k=key: self.reg_vars[k].set(
                               self.reg_vars[k].get().upper()))
                e.bind("<Return>", self._enter_paso1)
                if key == "fecha":
                    boton_calendario(campos, self.reg_vars[key], "legible",
                                     popup_parent=self.reg_frame
                                     ).grid(row=i, column=2, padx=(6, 0))
                if key == "curp":
                    def _validar_curp_ui(vkey=key):
                        ok, msg = validar_curp(self.reg_vars[vkey].get().strip())
                        (messagebox.showinfo if ok else messagebox.showwarning)(
                            "CURP", msg)
                    tk.Button(campos, text="Validar", command=_validar_curp_ui,
                              bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                              font=(FUENTE_SISTEMA, 9)).grid(row=i, column=2, padx=(6, 0))
            # Indicador en vivo: ya puedes registrar / homonima detectada (ADR-0041)
            self.lbl_reg_listo = tk.Label(self.reg_frame, text="",
                                          font=(FUENTE_SISTEMA, 10, "bold"),
                                          bg=CREMA_CLARO, anchor="w")
            self.lbl_reg_listo.grid(row=3, column=0, sticky="w", padx=30, pady=(2, 0))
            for k in ("ap_p", "ap_m", "nombres", "fecha"):
                self.reg_vars[k].trace_add("write", self._reglas_registro_vivo)

            # Bloque opcional: signos vitales + medidas (alimentan calculadoras)
            marco_sv = tk.Frame(self.reg_frame, bg=CREMA_CLARO)
            marco_sv.grid(row=4, column=0, sticky="we", padx=30, pady=(8, 0))
            tk.Label(marco_sv,
                     text="Signos vitales y medidas (opcional) — se guardan como primer control "
                          "y alimentan las calculadoras",
                     font=(FUENTE_SISTEMA, 10, "bold"), bg=CREMA_CLARO, fg=AZUL,
                     anchor="w").grid(row=0, column=0, columnspan=8, sticky="w", pady=(0, 4))
            sv_filas = [("Peso (kg)", "peso", 8), ("Talla (cm)", "talla", 8),
                        ("TA sist.", "ta_s", 6), ("TA diast.", "ta_d", 6),
                        ("FC (lpm)", "fc", 6), ("FR (rpm)", "fr", 6),
                        ("Temp (C)", "temp", 7), ("SpO2 (%)", "spo2", 7),
                        ("Cintura (cm)", "cintura", 9), ("Cadera (cm)", "cadera", 9),
                        ("Per. abdom. (cm)", "per_abdominal", 11),
                        ("Per. cefal. (cm)", "per_cefalico", 11)]
            self._sv_entries = {}
            for idx, (lab, key, w) in enumerate(sv_filas):
                r = 1 + idx // 4
                c = (idx % 4) * 2
                tk.Label(marco_sv, text=lab, anchor="w", bg=CREMA_CLARO,
                         fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 9)).grid(
                    row=r, column=c, sticky="w", pady=2, padx=(0, 3))
                e = tk.Entry(marco_sv, textvariable=self.reg_vars[key],
                             font=(FUENTE_SISTEMA, 10), bg=BLANCO, width=w)
                e.grid(row=r, column=c + 1, sticky="w", pady=2, padx=(0, 10))
                self._sv_entries[key] = e

            btns = tk.Frame(self.reg_frame, bg=CREMA_CLARO)
            btns.grid(row=5, column=0, sticky="we", padx=30, pady=16)
            tk.Button(btns, text="Guardar y crear expediente", command=self._guardar_paso1,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 11, "bold"), padx=12, pady=6
                      ).pack(side="right", padx=6)
            tk.Button(btns, text="Cancelar", command=lambda: self.show_page("home"),
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 10), padx=12, pady=6).pack(side="right")

        def _reglas_registro_vivo(self, *args):
            """ADR-0041: confirmacion en vivo de que ya se puede registrar + homonima."""
            if hasattr(self, "_after_reg_vivo"):
                try:
                    self.after_cancel(self._after_reg_vivo)
                except Exception:
                    pass
            self._after_reg_vivo = self.after(400, self._evaluar_registro_vivo)

        def _evaluar_registro_vivo(self):
            if not hasattr(self, "lbl_reg_listo"):
                return
            ap_p = self.reg_vars["ap_p"].get().strip()
            ap_m = self.reg_vars["ap_m"].get().strip()
            nombres = self.reg_vars["nombres"].get().strip()
            fecha = self.reg_vars["fecha"].get().strip()
            if not all([ap_p, ap_m, nombres, fecha]):
                self.lbl_reg_listo.config(text="", fg=GRIS_TEXTO)
                return
            if not parsear_fecha(fecha):
                self.lbl_reg_listo.config(
                    text="Fecha invalida — Usa 15-MAR-1985, 15/03/1985 o AAAA-MM-DD.",
                    fg=ROJO_AVISO)
                return
            # homonima: misma identidad ya registrada?
            nombre_completo = " ".join([ap_p, ap_m, nombres]).upper()
            clave = clave_paciente(normalizar(nombre_completo), parsear_fecha(fecha))
            ruta = ruta_por_clave(clave)
            if ruta.exists():
                self.lbl_reg_listo.config(
                    text="⚠ Ya existe una paciente con ese nombre y fecha (homónima).",
                    fg=ROJO_AVISO)
            else:
                self.lbl_reg_listo.config(
                    text="✓ Ya puedes registrar a la paciente.",
                    fg=VERDE_OK)

        def _enter_paso1(self, event):
            w = event.widget
            i = self.reg_entries.index(w) if w in self.reg_entries else -1
            if i == len(self.reg_entries) - 1:
                self._guardar_paso1()
            else:
                self.reg_entries[i + 1].focus_set()
            return "break"

        def _guardar_paso1(self):
            ap_p = self.reg_vars["ap_p"].get().strip()
            ap_m = self.reg_vars["ap_m"].get().strip()
            nombres = self.reg_vars["nombres"].get().strip()
            fecha = self.reg_vars["fecha"].get().strip()
            curp = self.reg_vars["curp"].get().strip()
            faltantes = [k for k, v in (("ap_p", ap_p), ("ap_m", ap_m),
                                        ("nombres", nombres), ("fecha", fecha))
                         if not v]
            if faltantes:
                for k in faltantes:
                    e = self._paso1_entries.get(k)
                    if e:
                        self._marcar_error(e)
                messagebox.showwarning("Faltan datos",
                                       "Completa apellido paterno, materno, nombre(s) y fecha.")
                return
            if not parsear_fecha(fecha):
                e = self._paso1_entries.get("fecha")
                if e:
                    self._marcar_error(e)
                messagebox.showwarning("Fecha invalida",
                                       "Usa 15-MAR-1985, 15/03/1985 o AAAA-MM-DD.")
                return
            ok, msg, ruta, datos = crear_paciente(
                ap_p, ap_m, nombres, fecha, curp,
                ventana_callback=lambda m: messagebox.askyesno("Confirmar", m))
            if not ok:
                if "Ya existe" in msg and ruta:
                    self.seleccionar_paciente(ruta)
                    if messagebox.askyesno("Expediente existente",
                                           f"{msg}\n\n¿Deseas abrir su carpeta?"):
                        abrir_carpeta(ruta)
                    self.ir_a_documentos()
                    return
                messagebox.showerror("Error", msg)
                return
            self.reg_ruta = ruta
            # ADR-0041: si capturaron signos vitales, crear el primer control
            # y guardar medidas corporales en H1 de HISTORIA.json
            sv = {k: self.reg_vars[k].get().strip() for k in
                  ("peso", "talla", "ta_s", "ta_d", "fc", "fr", "temp", "spo2")}
            hay_sv = any(sv.values())
            eid_primero = datos.get("episodio_activo") or ""
            if hay_sv and eid_primero:
                campos_ctrl = {"fecha": fecha_hoy()}
                for k, v in sv.items():
                    campos_ctrl[k] = v
                ok_c, msg_c, _ctrl = agregar_control(ruta, eid_primero, campos_ctrl)
                if not ok_c:
                    messagebox.showwarning("Control", f"No se pudo crear el control: {msg_c}")
            # medidas corporales -> H1 (HISTORIA.json)
            medidas = {k: self.reg_vars[k].get().strip() for k in
                       ("cintura", "cadera", "per_abdominal", "per_cefalico")}
            if any(medidas.values()):
                try:
                    h = cargar_historia(ruta)
                    h1 = h.setdefault("paciente", {}).setdefault("H1", {})
                    mapa_h1 = {"cintura": "cintura", "cadera": "cadera",
                               "per_abdominal": "perimetro_abdominal",
                               "per_cefalico": "pc"}
                    for k, v in medidas.items():
                        if v:
                            h1[mapa_h1[k]] = {"e": "si", "v": v}
                    guardar_historia(ruta, h)
                except Exception:
                    pass
            self._limpiar_registro(tuple(self.reg_vars.keys()))
            self.seleccionar_paciente(ruta)
            self.ir_a_documentos()
            # ADR-0041: ofrecer ver calculadoras precargadas
            if hay_sv and eid_primero:
                if messagebox.askyesno(
                        "Expediente creado",
                        f"Creado: {datos.get('nombre_completo', '')}\n"
                        f"Primer control registrado con {len([v for v in sv.values() if v])} "
                        "signos vitales.\n\n¿Ver calculadoras con estos datos?"):
                    self.show_page("calculadoras")  # ADR-0042: pagina embebida
                    return
            if messagebox.askyesno(
                    "Expediente creado",
                    f"Creado: {datos.get('nombre_completo', '')}\n"
                    "Ya esta seleccionada en Documentos.\n\n"
                    "¿Agregar telefono, correo, medico o notas ahora?"):
                self._paso2()
                self.show_page("register")

        def _paso2(self):
            for w in self.reg_frame.winfo_children():
                w.destroy()
            # C-04: prefill desde los datos reales de la paciente activa
            # (evita heredar valores tipeados de un registro anterior)
            try:
                d = cargar_paciente(self.reg_ruta) if self.reg_ruta else {}
            except Exception:
                d = {}
            self.reg_vars["telefono"].set(d.get("telefono", "") or "")
            self.reg_vars["correo"].set(d.get("correo", "") or "")
            self.reg_vars["medico"].set(d.get("medico_remitente", "") or "")
            self.reg_vars["notas"].set(d.get("notas", "") or "")
            _badge2 = tk.Frame(self.reg_frame, bg=ACENTO)
            _badge2.grid(row=0, column=0, sticky="w", padx=30, pady=(14, 2))
            tk.Label(_badge2, text="PASO 2 DE 2 · DATOS ADICIONALES (OPCIONAL)",
                     bg=ACENTO, fg=NAV_BG, font=(FUENTE_SISTEMA, 10, "bold"),
                     padx=10, pady=4).pack()
            tk.Label(self.reg_frame,
                     text="Puedes omitirlos: se completan despues desde la Ficha de la paciente.",
                     font=(FUENTE_SISTEMA, 10), bg=CREMA_CLARO, fg=GRIS_TEXTO
                     ).grid(row=1, column=0, sticky="w", padx=30, pady=(0, 8))
            campos = tk.Frame(self.reg_frame, bg=CREMA_CLARO)
            campos.grid(row=2, column=0, sticky="we", padx=30)
            campos.grid_columnconfigure(1, weight=1)
            filas = [("Telefono", "telefono"), ("Correo", "correo"),
                     ("Medico remitente", "medico"), ("Notas", "notas")]
            self.reg_entries2 = []
            for i, (lab, key) in enumerate(filas):
                tk.Label(campos, text=lab, anchor="w", bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                         font=(FUENTE_SISTEMA, 11)).grid(row=i, column=0, sticky="w",
                                                     pady=5, padx=(0, 10))
                e = tk.Entry(campos, textvariable=self.reg_vars[key], font=(FUENTE_SISTEMA, 12),
                             bg=BLANCO, fg=TEXTO_OSCURO, relief="solid", bd=1)
                e.grid(row=i, column=1, sticky="we", pady=5)
                self.reg_entries2.append(e)
                e.bind("<Return>", self._enter_paso2)
                if key == "medico":
                    e.bind("<KeyRelease>",
                           lambda ev, k=key: self.reg_vars[k].set(
                               self.reg_vars[k].get().upper()))
            btns = tk.Frame(self.reg_frame, bg=CREMA_CLARO)
            btns.grid(row=3, column=0, sticky="we", padx=30, pady=16)
            tk.Button(btns, text="Guardar datos", command=self._guardar_paso2,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 11, "bold"), padx=12, pady=6
                      ).pack(side="right", padx=6)
            tk.Button(btns, text="Omitir", command=self.ir_a_documentos,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 10), padx=12, pady=6).pack(side="right")

        def _enter_paso2(self, event):
            w = event.widget
            i = self.reg_entries2.index(w) if w in self.reg_entries2 else -1
            if i == len(self.reg_entries2) - 1:
                self._guardar_paso2()
            else:
                self.reg_entries2[i + 1].focus_set()
            return "break"

        def _guardar_paso2(self):
            if not self.reg_ruta:
                messagebox.showwarning("Aviso", "No hay expediente activo en este paso.")
                return
            guardar_datos_secundarios(
                self.reg_ruta,
                self.reg_vars["telefono"].get(),
                self.reg_vars["correo"].get(),
                "",
                self.reg_vars["medico"].get(),
                self.reg_vars["notas"].get())
            self._limpiar_registro(("telefono", "correo", "medico", "notas"))
            self.seleccionar_paciente(self.reg_ruta)
            messagebox.showinfo("Listo", "Datos adicionales guardados.")
            self.ir_a_documentos()

        # ---------- PAGINA: DOCUMENTOS ----------

        def _pagina_documentos(self):
            docs = tk.Frame(self.pages_box, bg=CREMA_CLARO)
            docs.grid_rowconfigure(0, weight=1)
            docs.grid_columnconfigure(0, weight=1)
            self.tabs = ttk.Notebook(docs)
            self.tabs.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
            tab_pac = tk.Frame(self.tabs, bg=CREMA_CLARO)
            tab_sin = tk.Frame(self.tabs, bg=CREMA_CLARO)
            tab_tpl = tk.Frame(self.tabs, bg=CREMA_CLARO)
            self.tabs.add(tab_pac, text="  Con paciente  ")
            self.tabs.add(tab_sin, text="  Sin paciente  ")
            self.tabs.add(tab_tpl, text="  Plantillas  ")
            self.tabs.bind("<<NotebookTabChanged>>",
                           lambda e: self._subtitulo_pestana())

            # -- Con paciente (v4.17: 2 columnas + barra de accion anclada) --
            # col0 = selector plegable · col1 = plantillas (altura completa)
            # row3 (cuerpo) expande · row4 (accion) jamas se recorta (R6 abajo)
            tab_pac.grid_columnconfigure(0, weight=1)
            tab_pac.grid_columnconfigure(1, weight=1)
            tab_pac.grid_rowconfigure(3, weight=1)

            def _btn_epi(parent, texto, cmd, padx=4):
                # secundario fantasma: no compite con el primario de generar (U74)
                return tk.Button(parent, text=texto, command=cmd, bg=BLANCO,
                                 fg=AZUL, activebackground=AZUL_CHIP, relief="flat",
                                 highlightthickness=1, highlightbackground=GRIS_BORDE,
                                 font=(FUENTE_SISTEMA, 9), padx=8, pady=3, cursor="hand2")

            card_sel = tk.Frame(tab_pac, bg=BLANCO, highlightthickness=1,
                                highlightbackground=GRIS_BORDE)
            card_sel.grid(row=0, column=0, columnspan=2, sticky="we",
                          padx=10, pady=(8, 2))
            self.lbl_sel = tk.Label(card_sel, text="Ninguna paciente seleccionada",
                                    font=(FUENTE_SISTEMA, 12, "bold"), bg=BLANCO,
                                    fg=ROJO_AVISO, anchor="w", padx=10, pady=7)
            self.lbl_sel.pack(fill="x")
            # fila del episodio en 2 lineas: sin desbordar a ~1050 px
            fila_epis = tk.Frame(tab_pac, bg=CREMA_CLARO)
            fila_epis.grid(row=1, column=0, columnspan=2, sticky="we",
                           padx=10, pady=2)
            sup_epis = tk.Frame(fila_epis, bg=CREMA_CLARO)
            sup_epis.pack(fill="x")
            tk.Label(sup_epis, text="Episodio (embarazo):", bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 10)).pack(side="left")
            self.epi_combo = ttk.Combobox(sup_epis, state="readonly", width=26,
                                          font=(FUENTE_SISTEMA, 10))
            self.epi_combo.pack(side="left", padx=6)
            self.epi_combo.bind("<<ComboboxSelected>>", self._cambiar_episodio)
            _btn_epi(sup_epis, "+ Nuevo embarazo",
                     self._nuevo_episodio).pack(side="left")
            _btn_epi(sup_epis, "[Cal] Datos del embarazo",
                     self._dlg_datos_embarazo).pack(side="left", padx=4)
            # C-16 (v2.1): controles en la fila del episodio -> la fila de
            # generacion se mantiene en 5 (R6) — decision ADR-0025/C-16 §9.2
            _btn_epi(sup_epis, "Controles",
                     self._abrir_controles).pack(side="left", padx=4)
            # 2a linea: expediente/herramientas (C-18, C-22, C-21) — R6 intacto
            sup2_epis = tk.Frame(fila_epis, bg=CREMA_CLARO)
            sup2_epis.pack(fill="x", pady=(3, 0))
            _btn_epi(sup2_epis, "Historia A-M",
                     self._historia_completa).pack(side="left")
            _btn_epi(sup2_epis, "Timeline",
                     self._abrir_timeline).pack(side="left", padx=4)
            _btn_epi(sup2_epis, "Historial anonimo",
                     self._exportar_anonimo).pack(side="left", padx=4)
            # U-PROP-6 (T6b): banda destacada FUM | FPP | EG del episodio activo
            banda_tira = tk.Frame(fila_epis, bg=AZUL_CHIP, highlightthickness=1,
                                  highlightbackground="#A9C6EA")
            banda_tira.pack(fill="x", pady=(5, 0))
            self.lbl_tira = tk.Label(banda_tira, text="FUM: -  |  FPP: -  |  EG: -",
                                     bg=AZUL_CHIP, fg=AZUL,
                                     font=(FUENTE_SISTEMA, FUENTE_M, "bold"),
                                     anchor="w", justify="left", padx=8, pady=5)
            self.lbl_tira.pack(fill="x")
            # U131 (v4.17): selector plegable de pacientes (col0)
            fila_cambio = tk.Frame(tab_pac, bg=CREMA_CLARO)
            fila_cambio.grid(row=2, column=0, columnspan=2, sticky="we",
                             padx=10, pady=(4, 2))
            self.docs_sel_btn = tk.Button(
                fila_cambio, text="▸ Elegir paciente",
                command=self._alternar_selector_docs, bg=BLANCO, fg=AZUL,
                activebackground=AZUL_CHIP, relief="flat", highlightthickness=1,
                highlightbackground=GRIS_BORDE, font=(FUENTE_SISTEMA, 9, "bold"),
                padx=10, pady=3, cursor="hand2")
            self.docs_sel_btn.pack(side="left")
            Tooltip(self.docs_sel_btn,
                    "Muestra u oculta los recientes y la busqueda interna")
            self.docs_sel_pista = tk.Label(
                fila_cambio, text="", bg=CREMA_CLARO, fg=GRIS_TERCIARIO,
                font=(FUENTE_SISTEMA, 9), anchor="e")
            self.docs_sel_pista.pack(side="right", fill="x", padx=(10, 0))
            self.docs_sel_frame = tk.Frame(tab_pac, bg=CREMA_CLARO)
            self.docs_sel_frame.grid(row=3, column=0, sticky="nsew",
                                     padx=(10, 5), pady=3)
            self.docs_sel_frame.grid_columnconfigure(0, weight=1)
            self.docs_sel_frame.grid_rowconfigure(1, weight=1)
            self.docs_sel_frame.grid_rowconfigure(4, weight=1)
            tk.Label(self.docs_sel_frame,
                     text="Pacientes recientes (clic = seleccionar, doble clic = Ficha):",
                     bg=CREMA_CLARO, fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 10),
                     anchor="w").grid(row=0, column=0, sticky="we", pady=(2, 0))
            marco_rec = tk.Frame(self.docs_sel_frame, bg=BLANCO, highlightthickness=1,
                                 highlightbackground=GRIS_BORDE)
            marco_rec.grid(row=1, column=0, sticky="nsew", pady=3)
            self.lb_recientes = tk.Listbox(marco_rec, height=4, font=(FUENTE_SISTEMA, 10),
                                           bg=BLANCO, fg=TEXTO_OSCURO,
                                           selectbackground=AZUL,
                                           selectforeground=TEXTO_CLARO, activestyle="none")
            self.lb_recientes.pack(side="left", fill="both", expand=True)
            scr_r = ttk.Scrollbar(marco_rec, orient="vertical",
                                  command=self.lb_recientes.yview)
            scr_r.pack(side="right", fill="y")
            self.lb_recientes.configure(yscrollcommand=scr_r.set)
            self.lb_recientes.bind("<<ListboxSelect>>", self._sel_reciente)
            self.lb_recientes.bind("<Double-Button-1>", lambda e: self._ficha_reciente())
            tk.Label(self.docs_sel_frame,
                     text="O buscar otra paciente (nombre, CURP, telefono o correo):",
                     bg=CREMA_CLARO, fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 10),
                     anchor="w").grid(row=2, column=0, sticky="we", pady=(6, 0))
            marco_busq = tk.Frame(self.docs_sel_frame, bg=CREMA_CLARO)
            marco_busq.grid(row=3, column=0, sticky="we", pady=3)
            marco_busq.grid_columnconfigure(0, weight=1)
            self.entry_busq_docs = tk.Entry(marco_busq, font=(FUENTE_SISTEMA, 11),
                                            bg=BLANCO, fg=TEXTO_OSCURO,
                                            relief="solid", bd=1)
            self.entry_busq_docs.grid(row=0, column=0, sticky="we", padx=(0, 6))
            self.entry_busq_docs.bind("<Return>",
                                      lambda e: self._buscar_en_docs())
            tk.Button(marco_busq, text="Buscar", command=self._buscar_en_docs,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER
                      ).grid(row=0, column=1, padx=3)
            marco_res = tk.Frame(self.docs_sel_frame, bg=BLANCO, highlightthickness=1,
                                 highlightbackground=GRIS_BORDE)
            marco_res.grid(row=4, column=0, sticky="nsew", pady=3)
            self.lb_busqueda = tk.Listbox(marco_res, height=3, font=(FUENTE_SISTEMA, 10),
                                          bg=BLANCO, fg=TEXTO_OSCURO,
                                          selectbackground=AZUL,
                                          selectforeground=TEXTO_CLARO, activestyle="none")
            self.lb_busqueda.pack(side="left", fill="both", expand=True)
            scr_b = ttk.Scrollbar(marco_res, orient="vertical",
                                  command=self.lb_busqueda.yview)
            scr_b.pack(side="right", fill="y")
            self.lb_busqueda.configure(yscrollcommand=scr_b.set)
            self.lb_busqueda.bind("<<ListboxSelect>>", self._sel_busqueda)
            # plantillas en la columna derecha con TODA la altura disponible
            self.docs_tpl_marco = tk.Frame(tab_pac, bg=BLANCO, highlightthickness=1,
                                           highlightbackground=GRIS_BORDE)
            self.docs_tpl_marco.grid(row=3, column=1, sticky="nsew",
                                     padx=(5, 10), pady=3)
            self.docs_tpl_marco.grid_rowconfigure(1, weight=1)
            self.docs_tpl_marco.grid_columnconfigure(0, weight=1)
            tk.Label(self.docs_tpl_marco, text="Plantilla:", bg=BLANCO, fg=TEXTO_OSCURO,
                     font=(FUENTE_SISTEMA, 10, "bold"), anchor="w"
                     ).grid(row=0, column=0, sticky="we", padx=8, pady=(6, 2))
            marco_lb_tpl = tk.Frame(self.docs_tpl_marco, bg=BLANCO)
            marco_lb_tpl.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
            marco_lb_tpl.grid_columnconfigure(0, weight=1)
            marco_lb_tpl.grid_rowconfigure(0, weight=1)
            self.lb_plantillas = tk.Listbox(marco_lb_tpl, height=4,
                                            font=(FUENTE_SISTEMA, 10),
                                            bg=BLANCO, fg=TEXTO_OSCURO,
                                            selectbackground=AZUL,
                                            selectforeground=TEXTO_CLARO, activestyle="none")
            self.lb_plantillas.grid(row=0, column=0, sticky="nsew")
            scr_p = ttk.Scrollbar(marco_lb_tpl, orient="vertical",
                                  command=self.lb_plantillas.yview)
            scr_p.grid(row=0, column=1, sticky="ns")
            self.lb_plantillas.configure(yscrollcommand=scr_p.set)
            # fila de accion anclada (weight 0): siempre visible sin scroll
            fila_acc = tk.Frame(tab_pac, bg=CREMA_CLARO)
            fila_acc.grid(row=4, column=0, columnspan=2, sticky="we",
                          padx=10, pady=(8, 6))
            tk.Button(fila_acc, text="Generar y abrir documento",
                      command=self._generar_docs, bg=AZUL, fg=TEXTO_CLARO,
                      activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 11, "bold"),
                      padx=10, pady=6).pack(side="left", padx=4)
            tips_docs = {
                "Progreso Quimico": "Archivo vivo del episodio; se grafica en Graficos",
                "Ficha": "Ver/editar todos los datos de la paciente",
                "Consentimiento": "Consentimiento informado con hash de contenido",
                "Historia (NOM-004)": "Historia clinica con vista previa (F4)",
            }
            for txt, cmd in [
                    ("Progreso Quimico", self._progreso_quimico_docs),
                    # C1: Ficha reemplaza a "Abrir carpeta"
                    ("Ficha", self._ficha_de_docs),
                    ("Consentimiento", self._dlg_consentimiento),
                    ("Historia (NOM-004)", self._dlg_nom004)]:
                bb = tk.Button(fila_acc, text=txt, command=cmd, bg=GRIS_SECUNDARIO,
                               fg=TEXTO_CLARO, activebackground=GRIS_APAGADO,
                               font=(FUENTE_SISTEMA, 10), padx=10, pady=6)
                bb.pack(side="left", padx=4)
                if txt in tips_docs:
                    Tooltip(bb, tips_docs[txt])

            # -- Sin paciente --
            tab_sin.grid_columnconfigure(0, weight=1)
            tab_sin.grid_rowconfigure(3, weight=1)
            tk.Label(tab_sin, text="Nombre del documento:", bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 11), anchor="w"
                     ).grid(row=0, column=0, sticky="we", padx=12, pady=(12, 2))
            self.doc_sin_var = tk.StringVar()
            tk.Entry(tab_sin, textvariable=self.doc_sin_var, font=(FUENTE_SISTEMA, 11),
                     bg=BLANCO, relief="solid", bd=1).grid(row=1, column=0,
                                                           sticky="we", padx=12, pady=3)
            tk.Label(tab_sin, text="Plantilla:", bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                     font=(FUENTE_SISTEMA, 11), anchor="w").grid(row=2, column=0,
                                                             sticky="we", padx=12, pady=(8, 2))
            self.lb_plantillas_sin = tk.Listbox(tab_sin, height=6, font=(FUENTE_SISTEMA, 10),
                                                bg=BLANCO, fg=TEXTO_OSCURO,
                                                selectbackground=AZUL,
                                                selectforeground=TEXTO_CLARO)
            self.lb_plantillas_sin.grid(row=3, column=0, sticky="nsew", padx=12, pady=3)
            tk.Button(tab_sin, text="Generar documento general",
                      command=self._generar_sin_paciente, bg=AZUL, fg=TEXTO_CLARO,
                      activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 11),
                      pady=6).grid(row=4, column=0, pady=12)

            # -- Plantillas (edicion solo .txt) --
            tab_tpl.grid_columnconfigure(0, weight=1)
            tab_tpl.grid_rowconfigure(2, weight=1)
            tk.Label(tab_tpl, text="Gestionar plantillas",
                     font=(FUENTE_SISTEMA, 14, "bold"), bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).grid(row=0, column=0, pady=10)
            fila_c = tk.Frame(tab_tpl, bg=CREMA_CLARO)
            fila_c.grid(row=1, column=0, sticky="we", padx=12, pady=4)
            fila_c.grid_columnconfigure(1, weight=1)
            tk.Label(fila_c, text="Plantilla:", bg=CREMA_CLARO, fg=TEXTO_OSCURO
                     ).grid(row=0, column=0, padx=(0, 6))
            self.tpl_combo = ttk.Combobox(fila_c, state="readonly",
                                          font=(FUENTE_SISTEMA, 10))
            self.tpl_combo.grid(row=0, column=1, sticky="we")
            self.tpl_combo.bind("<<ComboboxSelected>>",
                                lambda e: self._refrescar_tipo_tpl())
            # U132 (v4.17): indicador de tipo junto al combo
            self.docs_tpl_tipo = tk.Label(fila_c, text="", bg=CREMA_CLARO,
                                          fg=GRIS_TERCIARIO,
                                          font=(FUENTE_SISTEMA, 9))
            self.docs_tpl_tipo.grid(row=0, column=2, padx=(8, 0))
            self.tpl_text = tk.Text(tab_tpl, font=("Consolas", 10), bg=BLANCO,
                                    fg=TEXTO_OSCURO, relief="solid", bd=1)
            self.tpl_text.grid(row=2, column=0, sticky="nsew", padx=12, pady=6)
            fila_b = tk.Frame(tab_tpl, bg=CREMA_CLARO)
            fila_b.grid(row=3, column=0, sticky="we", padx=12, pady=10)
            for txt, cmd in [("Cargar (.txt)", self._cargar_tpl),
                             ("Guardar (.txt)", self._guardar_tpl),
                             ("Abrir plantilla", self._abrir_tpl),
                             ("Nueva plantilla (.txt)", self._nueva_plantilla),
                             ("Recargar lista", self._recargar_tpls)]:
                tk.Button(fila_b, text=txt, command=cmd, bg=AZUL, fg=TEXTO_CLARO,
                          activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 9),
                          padx=8, pady=4).pack(side="left", padx=4)
            tk.Label(tab_tpl,
                     text="Solo .txt se edita aqui. .docx/.xlsx se abren con Word/Excel (boton Abrir) para no corromperlos.",
                     bg=CREMA_CLARO, fg=GRIS_TERCIARIO, font=(FUENTE_SISTEMA, 9)
                     ).grid(row=4, column=0, sticky="w", padx=12, pady=4)

            self.refrescar_recientes()
            self._recargar_tpls()
            self._poblar_plantillas()
            # estado inicial U131: abierto solo si entramos sin paciente
            self._docs_selector(self.ruta_activa is None)
            return docs

        def _docs_selector(self, expandido):
            """U131 (v4.17): selector de pacientes plegable en C-05.

            Sin paciente arranca desplegado; al elegir una se pliega y la
            lista de plantillas ocupa las dos columnas. La fila de accion
            (row 4, weight 0) nunca se recorta.
            """
            if not hasattr(self, "docs_sel_frame"):
                return
            self.docs_sel_expandido = bool(expandido)
            if expandido:
                self.docs_sel_frame.grid()
                self.docs_tpl_marco.grid(column=1, columnspan=1, sticky="nsew",
                                         padx=(5, 10), pady=3)
            else:
                self.docs_sel_frame.grid_remove()
                self.docs_tpl_marco.grid(column=0, columnspan=2, sticky="nsew",
                                         padx=10, pady=3)
            base = "Cambiar paciente" if self.ruta_activa else "Elegir paciente"
            self.docs_sel_btn.config(text=("▾ " if expandido else "▸ ") + base)
            if expandido:
                pista = "clic = seleccionar · doble clic = Ficha"
            else:
                fin = ("cambiar de paciente" if self.ruta_activa
                       else "elegir paciente")
                pista = f"Recientes y búsqueda ocultos — pulsa para {fin}"
            self.docs_sel_pista.config(text=pista)

        def _alternar_selector_docs(self):
            self._docs_selector(not getattr(self, "docs_sel_expandido", True))

        def _poblar_plantillas(self):
            self.lb_plantillas.delete(0, "end")
            self.lb_plantillas_sin.delete(0, "end")
            self.tpl_lista = listar_plantillas()
            if not self.tpl_lista:
                vacio = "(No hay plantillas en _PLANTILLAS)"
                self.lb_plantillas.insert("end", vacio)
                self.lb_plantillas_sin.insert("end", vacio)
                return
            for t in self.tpl_lista:
                self.lb_plantillas.insert("end", t)
                self.lb_plantillas_sin.insert("end", t)

        def _buscar_en_docs(self):
            q = self.entry_busq_docs.get().strip()
            self.lb_busqueda.delete(0, "end")
            self.busq_rutas = []
            if not q:
                return
            res = buscar(q)
            if not res:
                self.lb_busqueda.insert("end", " (Sin resultados) ")
                return
            for nombre, ruta in res:
                self.lb_busqueda.insert("end", nombre)
                self.busq_rutas.append(ruta)

        def _sel_reciente(self, event):
            sel = self.lb_recientes.curselection()
            if sel:
                self.seleccionar_paciente(self.recientes_rutas[sel[0]], origen="recientes")

        def _sel_busqueda(self, event):
            sel = self.lb_busqueda.curselection()
            if sel and sel[0] < len(self.busq_rutas):
                self.seleccionar_paciente(self.busq_rutas[sel[0]], origen="busqueda")

        def _ficha_reciente(self):
            """Doble clic en recientes: expediente EN EL SOFTWARE (Ficha), no la
            carpeta del sistema (requisito del dueño 2026-10-16)."""
            sel = self.lb_recientes.curselection()
            if not sel or sel[0] >= len(self.recientes_rutas):
                return
            ruta = self.recientes_rutas[sel[0]]
            self.seleccionar_paciente(ruta, origen="recientes")
            FichaPaciente(self, ruta)

        def _abrir_ficha(self, ruta):
            if not ruta:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            FichaPaciente(self, ruta)

        def _ficha_de_docs(self):
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            self._abrir_ficha(self.ruta_activa)

        def _actualizar_episodios(self):
            labels = []
            self._epi_mapa = {}
            activo_label = ""
            if self.ruta_activa and self.datos_activos:
                for e in self.datos_activos.get("episodios", []):
                    eid = e.get("id", "")
                    ref = e.get("fecha_referencia")
                    lbl = f"{eid} ({fecha_legible(ref)})" if ref else eid
                    labels.append(lbl)
                    self._epi_mapa[lbl] = eid
                    if eid == self.datos_activos.get("episodio_activo"):
                        activo_label = lbl
            self.epi_combo.config(values=labels)
            self.epi_combo.set(activo_label)

        def _cambiar_episodio(self, event=None):
            if not self.ruta_activa:
                return
            etiqueta = self.epi_combo.get()
            eid = getattr(self, "_epi_mapa", {}).get(etiqueta, etiqueta)
            if eid and fijar_episodio(self.ruta_activa, eid):
                self.datos_activos = cargar_paciente(self.ruta_activa)
                self._actualizar_etiqueta()
                self.status.set(f"Episodio activo: {eid}")

        def _nuevo_episodio(self):
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            win = tk.Toplevel(self)
            win.title("Nuevo embarazo / episodio")
            win.configure(bg=CREMA)
            win.resizable(False, False)
            win.transient(self)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=16, pady=16)
            marco.pack(fill="both", expand=True, padx=10, pady=10)
            tk.Label(marco, text="Nuevo embarazo / episodio",
                     font=(FUENTE_SISTEMA, 14, "bold"), bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).pack(anchor="w")
            tk.Label(marco, text=f"Paciente: {nombre_mostrar(self.ruta_activa)}",
                     font=(FUENTE_SISTEMA, 10, "bold"), bg=CREMA_CLARO, fg=AZUL
                     ).pack(anchor="w", pady=(2, 0))
            tk.Label(marco, text="Elige la fecha del parto (si ya ocurrio) o la fecha de "
                                 "registro del embarazo actual. La FUM nace con hoy; "
                                 "cambiala si la real es otra.",
                     bg=CREMA_CLARO, fg=GRIS_TEXTO, wraplength=440, justify="left"
                     ).pack(anchor="w", pady=(4, 10))

            def _campo(texto_label):
                fila = tk.Frame(marco, bg=CREMA_CLARO)
                fila.pack(fill="x", pady=4)
                tk.Label(fila, text=texto_label, bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                         font=(FUENTE_SISTEMA, 11)).pack(side="left")
                var = tk.StringVar()
                e = tk.Entry(fila, textvariable=var, font=(FUENTE_SISTEMA, 12), width=14,
                             bg=BLANCO, justify="center")
                e.pack(side="left", padx=8)
                return e, var

            e_ref, var_ref = _campo("Fecha del embarazo:")
            var_ref.set(fecha_hoy())
            e_fum, var_fum = _campo("FUM (opcional):")
            # Requisito del dueño (2026-10-16): la FUM nace con HOY; si no es la
            # real, cambiala — el label FPP/EG en vivo hace visible cualquier
            # guardado (C2: nunca bloqueo duro).
            var_fum.set(fecha_hoy())
            lbl_calc, _actualizar_calc = _label_calculo(marco, var_fum)
            lbl_calc.pack(anchor="w", pady=(8, 4))

            for e, var in ((e_ref, var_ref), (e_fum, var_fum)):
                boton_calendario(e.master, var, "iso", popup_parent=win,
                                 on_change=_actualizar_calc).pack(side="left")
                e.bind("<KeyRelease>", _actualizar_calc)

            _actualizar_calc()

            def _crear():
                fecha_ref = parsear_fecha(var_ref.get())
                if not fecha_ref:
                    messagebox.showwarning("Fecha invalida",
                                           "Fecha del embarazo invalida. Usa AAAA-MM-DD "
                                           "o el boton [Cal].", parent=win)
                    return
                texto_fum = var_fum.get().strip()
                fum = None
                if texto_fum:
                    fum = parsear_fecha(texto_fum)
                    if not fum:
                        messagebox.showwarning("FUM invalida",
                                               "La FUM no es una fecha valida.", parent=win)
                        return
                    if fum > fecha_hoy() and not messagebox.askyesno(
                            "FUM futura",
                            "La FUM es posterior a hoy. ¿Continuar igual?", parent=win):
                        return
                    if fum > fecha_ref and not messagebox.askyesno(
                            "Revisar fechas",
                            "La FUM es posterior a la fecha del embarazo/parto.\n"
                            "¿Continuar igual?", parent=win):
                        return
                eid = nuevo_episodio(self.ruta_activa, fecha_ref=fecha_ref, fum=fum)
                self.datos_activos = cargar_paciente(self.ruta_activa)
                self._actualizar_episodios()
                self._actualizar_etiqueta()
                self.status.set(f"Episodio activo: {eid}")
                win.destroy()
                messagebox.showinfo(
                    "Nuevo embarazo",
                    f"Episodio creado y activado: {eid}\n"
                    "Los documentos y el Progreso Quimico se guardaran aqui.")

            pie = tk.Frame(marco, bg=CREMA_CLARO)
            pie.pack(fill="x", pady=(14, 0))
            tk.Button(pie, text="Crear episodio", command=_crear, bg=AZUL,
                      fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 11, "bold"), padx=10, pady=6
                      ).pack(side="right", padx=4)
            tk.Button(pie, text="Cancelar", command=win.destroy, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO,
                      padx=10, pady=6).pack(side="right")
            e_ref.bind("<Return>", lambda ev: e_fum.focus_set())
            e_fum.bind("<Return>", lambda ev: _crear())
            e_ref.focus_set()
            e_ref.select_range(0, "end")
            auto_tamano(win, self)

        def _abrir_controles(self):
            """C-16: pantalla de enfermera sobre el episodio en contexto."""
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            ControlesPrenatales(self)

        def dialogo_calculadoras(self):
            """C-19/ADR-0042: redirige a la pagina embebida (sin ventana)."""
            self.show_page("calculadoras")

        def _historia_completa(self):
            """C-18: editor A-M (si no hay paciente, aviso)."""
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            HistoriaCompleta(self)

        def _dictar(self):
            """ADR-0040: dictado de voz para campos narrativos de C-12."""
            ok, msg = dictado_disponible()
            win = tk.Toplevel(self)
            win.title("Dictado de voz")
            win.configure(bg=CREMA)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=14, pady=12)
            marco.pack(fill="both", expand=True)
            tk.Label(marco, text="Dictado de voz para la historia clinica",
                     font=(FUENTE_SISTEMA, 12, "bold"), bg=CREMA_CLARO, fg=AZUL
                     ).pack(anchor="w")
            if ok:
                tk.Label(marco, text="Motor Vosk disponible. Usa el dictado del SO "
                                     "(Win+H / fn fn) en el campo que quieras, o "
                                     "instala sounddevice para dictado por micrófono.",
                         bg=CREMA_CLARO, fg=VERDE_OK, wraplength=400,
                         justify="left", font=(FUENTE_SISTEMA, 10)).pack(anchor="w", pady=6)
            else:
                tk.Label(marco, text=f"Dictado disponible:\n{msg}\n\n"
                                     "En Windows: presiona Win+H y dicta.\n"
                                     "En macOS: presiona fn dos veces y dicta.\n"
                                     "En Linux: Super+Ctrl+Espacio.\n\n"
                                     "Para dictado por micrófono sin conexión,\n"
                                     "instala Vosk: pip install vosk",
                         bg=CREMA_CLARO, fg=TEXTO_OSCURO, wraplength=400,
                         justify="left", font=(FUENTE_SISTEMA, 10)).pack(anchor="w", pady=6)
            tk.Button(marco, text="Cerrar", command=win.destroy, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO, padx=10, pady=6
                      ).pack(anchor="e", pady=8)
            auto_tamano(win, self)

        def _historia_rapida(self):
            """C-18 §5: Ctrl+H -> A-M si existe HISTORIA.json, si no C-12."""
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            if ruta_historia(self.ruta_activa).exists():
                HistoriaCompleta(self)
            else:
                self._dlg_nom004()

        def _dlg_datos_embarazo(self):
            if not self.ruta_activa or not self.datos_activos:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            eid = self.datos_activos.get("episodio_activo")
            epi = None
            for e in self.datos_activos.get("episodios", []):
                if e.get("id") == eid:
                    epi = e
                    break
            if epi is None:
                messagebox.showinfo("Aviso", "La paciente no tiene episodios.")
                return
            win = tk.Toplevel(self)
            win.title("Datos del embarazo")
            win.configure(bg=CREMA)
            win.resizable(False, False)
            win.transient(self)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=16, pady=16)
            marco.pack(fill="both", expand=True, padx=10, pady=10)
            tk.Label(marco, text=f"Datos del embarazo - Episodio {eid}",
                     font=(FUENTE_SISTEMA, 14, "bold"), bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).pack(anchor="w")
            tk.Label(marco, text="Edicion posterior: FUM y fecha de referencia "
                                 "(parto/registro). Pulsa Guardar para aplicar cambios.",
                     bg=CREMA_CLARO, fg=GRIS_TEXTO, wraplength=450, justify="left"
                     ).pack(anchor="w", pady=(4, 10))

            def _campo(texto_label, valor):
                fila = tk.Frame(marco, bg=CREMA_CLARO)
                fila.pack(fill="x", pady=4)
                tk.Label(fila, text=texto_label, bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                         font=(FUENTE_SISTEMA, 11)).pack(side="left")
                var = tk.StringVar(value=valor or "")
                e = tk.Entry(fila, textvariable=var, font=(FUENTE_SISTEMA, 12), width=14,
                             bg=BLANCO, justify="center")
                e.pack(side="left", padx=8)
                return e, var

            e_ref, var_ref = _campo("Fecha referencia:", epi.get("fecha_referencia", ""))
            # Requisito del dueño: FUM preseleccionada con hoy cuando no hay FUM
            # guardada; si ya existe, se conserva la del episodio.
            e_fum, var_fum = _campo("FUM:", epi.get("fum") or fecha_hoy())
            lbl_calc, _actualizar_calc = _label_calculo(
                marco, var_fum,
                texto_vacio="(sin FUM: FPP y EG no disponibles)", wrap=450)
            lbl_calc.pack(anchor="w", pady=(8, 4))

            for e, var in ((e_ref, var_ref), (e_fum, var_fum)):
                boton_calendario(e.master, var, "iso", popup_parent=win,
                                 on_change=_actualizar_calc).pack(side="left")
                e.bind("<KeyRelease>", _actualizar_calc)

            _actualizar_calc()

            # --- Tier A (v2.0): campos del episodio (ADR-0024) ---
            e_gpa, var_gpa = _campo("G-P-A (ej. G3P2A1):", epi.get("gpa", ""))
            fila_tp = tk.Frame(marco, bg=CREMA_CLARO)
            fila_tp.pack(fill="x", pady=4)
            tk.Label(fila_tp, text="Tipo parto anterior:", bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 11)).pack(side="left")
            var_tp = tk.StringVar(value=epi.get("tipo_parto_anterior", ""))
            ttk.Combobox(fila_tp, textvariable=var_tp, width=16,
                         values=("Espontaneo", "Cesarea", "Instrumentado",
                                 "Multiple", "No aplica"),
                         font=(FUENTE_SISTEMA, 11)).pack(side="left", padx=8)
            e_comp, var_comp = _campo("Complicaciones previas:",
                                      epi.get("complicaciones_previas", ""))
            e_fpa, var_fpa = _campo("Fecha parto anterior (opcional):",
                                    epi.get("fecha_parto_anterior", ""))
            boton_calendario(e_fpa.master, var_fpa, "iso", popup_parent=win
                             ).pack(side="left")

            def _guardar():
                texto_ref = var_ref.get().strip()
                fecha_ref = None
                if texto_ref:
                    fecha_ref = parsear_fecha(texto_ref)
                    if not fecha_ref:
                        messagebox.showwarning("Fecha invalida",
                                               "Fecha de referencia invalida.", parent=win)
                        return
                texto_fum = var_fum.get().strip()
                fum = None
                if texto_fum:
                    fum = parsear_fecha(texto_fum)
                    if not fum:
                        messagebox.showwarning("FUM invalida", "La FUM no es valida.",
                                               parent=win)
                        return
                    if fum > fecha_hoy() and not messagebox.askyesno(
                            "FUM futura",
                            "La FUM es posterior a hoy. ¿Guardar igual?", parent=win):
                        return
                    if fecha_ref and fum > fecha_ref and not messagebox.askyesno(
                            "Revisar fechas",
                            "La FUM es posterior a la fecha de referencia. "
                            "¿Guardar igual?", parent=win):
                        return
                # --- Tier A (v2.0): validacion de campos del episodio ---
                gpa_txt = var_gpa.get().strip()
                ok_gpa, msg_gpa = validar_gpa(gpa_txt)
                if not ok_gpa:
                    messagebox.showwarning("G-P-A invalido", msg_gpa, parent=win)
                    return
                fpa_txt = var_fpa.get().strip()
                fpa_iso = None
                if fpa_txt:
                    fpa_iso = parsear_fecha(fpa_txt)
                    if not fpa_iso:
                        messagebox.showwarning("Fecha invalida",
                                               "Fecha de parto anterior invalida.",
                                               parent=win)
                        return
                    if fpa_iso > fecha_hoy() and not messagebox.askyesno(
                            "Fecha futura",
                            "La fecha del parto anterior es posterior a hoy. "
                            "¿Guardar igual?", parent=win):
                        return
                campos_tier = {
                    "gpa": gpa_txt or None,
                    "tipo_parto_anterior": var_tp.get().strip() or None,
                    "complicaciones_previas": var_comp.get().strip() or None,
                    "fecha_parto_anterior": fpa_iso,
                }
                if actualizar_episodio(self.ruta_activa, eid,
                                       {"fum": fum, "fecha_referencia": fecha_ref,
                                        **campos_tier}):
                    self.datos_activos = cargar_paciente(self.ruta_activa)
                    self._actualizar_episodios()
                    self._actualizar_etiqueta()
                    self.status.set(f"Episodio {eid} actualizado")
                    win.destroy()
                    messagebox.showinfo("Guardado",
                                        "Datos del embarazo actualizados.")
                else:
                    messagebox.showerror("Error", "No se pudo guardar.",
                                         parent=win)

            def _limpiar_fum():
                var_fum.set("")
                _actualizar_calc()

            e_fum.bind("<Return>", lambda ev: _actualizar_calc())
            pie = tk.Frame(marco, bg=CREMA_CLARO)
            pie.pack(fill="x", pady=(14, 0))
            tk.Button(pie, text="Guardar", command=_guardar, bg=AZUL,
                      fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 11, "bold"), padx=10, pady=6
                      ).pack(side="right", padx=4)
            tk.Button(pie, text="Limpiar FUM", command=_limpiar_fum, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO,
                      padx=8, pady=6).pack(side="right", padx=4)
            tk.Button(pie, text="Cerrar", command=win.destroy, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO,
                      padx=8, pady=6).pack(side="right")
            e_fum.focus_set()
            auto_tamano(win, self)

        def _generar_docs(self):
            if not self.ruta_activa:
                messagebox.showwarning("Aviso", "Selecciona una paciente.")
                return
            if not self.tpl_lista:
                messagebox.showinfo("Aviso", "No hay plantillas en _PLANTILLAS. "
                                             "Coloca tus plantillas ahi.")
                return
            sel = self.lb_plantillas.curselection()
            if not sel:
                messagebox.showwarning("Aviso", "Selecciona una plantilla.")
                return
            tpl = self.tpl_lista[sel[0]]
            self.generar_con(self.ruta_activa, tpl)

        def generar_con(self, ruta, tpl, directorio=None):
            mapa_extra = None
            try:
                datos = cargar_paciente(ruta)
                falta_doctor = (plantilla_requiere_doctor(tpl)
                                and not datos.get("medico_remitente"))
            except Exception:
                falta_doctor = False
            if falta_doctor:
                if messagebox.askyesno(
                        "Remitente",
                        "La plantilla usa {{DOCTOR}} y no hay medico remitente registrado.\n\n"
                        "¿Usar 'A QUIEN CORRESPONDA'?"):
                    mapa_extra = {"{{DOCTOR}}": "A QUIEN CORRESPONDA"}

            def cb_existente(nombre):
                return messagebox.askyesno(
                    "Archivo existente",
                    f"Ya existe: {nombre}\n\n"
                    "Si = abrirlo para editarlo (no se toca el archivo)\n"
                    "No = generar un documento nuevo con sufijo numerado")

            ok, msg, salida = generar_documento(
                ruta, tpl, callback_existente=cb_existente,
                directorio=directorio, mapa_extra=mapa_extra)
            if ok and salida:
                try:
                    ir_a_ruta(salida)
                except Exception as e:
                    log_exc(e)
                    messagebox.showwarning("Aviso", f"Documento listo pero no se pudo abrir:\n{e}")
                messagebox.showinfo("Documento", f"{msg}\n\nEdita y guarda (Ctrl+S); cierra cuando termines.")
            else:
                messagebox.showerror("Error", msg)

        def _progreso_quimico_docs(self):
            if not self.ruta_activa:
                messagebox.showwarning("Aviso", "Selecciona una paciente.")
                return
            ok, msg, salida = progreso_quimico(self.ruta_activa)
            if ok and salida:
                try:
                    ir_a_ruta(salida)
                except Exception as e:
                    log_exc(e)
                    messagebox.showwarning("Aviso", f"No se pudo abrir: {e}")
            if not ok:
                messagebox.showerror("Error", msg)

        def _generar_sin_paciente(self):
            if not self.tpl_lista:
                messagebox.showinfo("Aviso", "No hay plantillas en _PLANTILLAS. "
                                             "Coloca tus plantillas ahi.")
                return
            sel = self.lb_plantillas_sin.curselection()
            if not sel:
                messagebox.showwarning("Aviso", "Selecciona una plantilla.")
                return
            tpl = self.lb_plantillas_sin.get(sel[0])
            nombre = self.doc_sin_var.get().strip() or "Documento"
            ok, msg, salida = generar_sin_paciente(nombre, tpl)
            if ok and salida:
                ir_a_ruta(salida)
                messagebox.showinfo("Documento", msg)
            else:
                messagebox.showerror("Error", msg)

        def _cargar_tpl(self):
            nombre = self.tpl_combo.get()
            if not nombre:
                return
            if not nombre.lower().endswith(".txt"):
                messagebox.showinfo("Solo texto",
                                    "La edicion interna es solo para .txt.\n"
                                    "Usa 'Abrir plantilla' para Word/Excel.")
                return
            p = PLANTILLAS_DIR / nombre
            self.tpl_text.delete("1.0", "end")
            try:
                self.tpl_text.insert("1.0", p.read_text(encoding="utf-8"))
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer: {e}")

        def _guardar_tpl(self):
            nombre = self.tpl_combo.get()
            if not nombre or not nombre.lower().endswith(".txt"):
                messagebox.showwarning("Aviso", "Selecciona una plantilla .txt.")
                return
            p = PLANTILLAS_DIR / nombre
            p.write_text(self.tpl_text.get("1.0", "end"), encoding="utf-8")
            messagebox.showinfo("Listo", f"Plantilla {nombre} guardada.")

        def _abrir_tpl(self):
            nombre = self.tpl_combo.get()
            if not nombre:
                return
            try:
                ir_a_ruta(PLANTILLAS_DIR / nombre)
            except Exception as e:
                messagebox.showerror("Error", str(e))

        def _recargar_tpls(self):
            tpls = listar_plantillas()
            self.tpl_combo.config(values=tpls)
            if tpls:
                self.tpl_combo.set(tpls[0])
            else:
                self.tpl_combo.set("")
            self._poblar_plantillas()
            self._refrescar_tipo_tpl()

        def _refrescar_tipo_tpl(self):
            """U132 (v4.17): [.txt editable] / [DOCX solo abrir] junto al combo."""
            if not hasattr(self, "docs_tpl_tipo"):
                return
            nombre = self.tpl_combo.get()
            if not nombre:
                self.docs_tpl_tipo.config(text="", fg=GRIS_TERCIARIO)
            elif nombre.lower().endswith(".txt"):
                self.docs_tpl_tipo.config(text="[.txt · editable aqui]",
                                          fg=VERDE_OK)
            else:
                ext = nombre.rsplit(".", 1)[-1].upper()
                self.docs_tpl_tipo.config(
                    text=f"[{ext} · solo abrir con Word/Excel]", fg=GRIS_TERCIARIO)

        def _nueva_plantilla(self):
            # C3: crear plantilla .txt con nombre validado
            win = tk.Toplevel(self)
            win.title("Nueva plantilla")
            win.configure(bg=CREMA)
            win.resizable(False, False)
            win.transient(self)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=16, pady=14)
            marco.pack(fill="both", expand=True, padx=10, pady=10)
            tk.Label(marco, text="Nueva plantilla de texto",
                     font=(FUENTE_SISTEMA, 13, "bold"), bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).pack(anchor="w")
            tk.Label(marco, text="Nombre (se guardara como .txt en _PLANTILLAS):",
                     bg=CREMA_CLARO, fg=GRIS_TEXTO).pack(anchor="w", pady=(6, 2))
            var = tk.StringVar(value="PLANTILLA_NUEVA.txt")
            e = tk.Entry(marco, textvariable=var, font=(FUENTE_SISTEMA, 12), bg=BLANCO)
            e.pack(fill="x")
            e.select_range(0, "end")
            e.focus_set()

            def crear(*_args):
                crudo = var.get().strip()
                if not crudo:
                    messagebox.showwarning("Aviso", "Escribe un nombre.", parent=win)
                    return
                if crudo.lower().endswith(".txt"):
                    crudo = crudo[:-4]
                stem = normalizar(crudo)
                if not stem:
                    messagebox.showwarning(
                        "Nombre invalido",
                        "El nombre debe contener letras o numeros.", parent=win)
                    return
                destino = PLANTILLAS_DIR / f"{stem}.txt"
                if destino.exists():
                    messagebox.showwarning(
                        "Ya existe", f"Ya hay una plantilla llamada {destino.name}.",
                        parent=win)
                    return
                destino.write_text(
                    "Escribe aqui tu plantilla.\n"
                    "Marcadores: {{PACIENTE}} {{CURP}} {{FECHA}} {{DOCTOR}} "
                    "{{FUM}} {{FPP}} {{EG}} {{EDAD}} {{GP_A}} {{ALERGIAS}} "
                    "{{GRUPO}} {{PARTO_ANT}}\n", encoding="utf-8")
                win.destroy()
                self._recargar_tpls()
                self.tpl_combo.set(destino.name)
                self._refrescar_tipo_tpl()
                self._cargar_tpl()
                self.status.set(f"Plantilla creada: {destino.name}")

            e.bind("<Return>", crear)
            fila = tk.Frame(marco, bg=CREMA_CLARO)
            fila.pack(fill="x", pady=(12, 0))
            tk.Button(fila, text="Crear", command=crear, bg=AZUL, fg=TEXTO_CLARO,
                      activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 11, "bold"),
                      padx=10, pady=5).pack(side="right", padx=4)
            tk.Button(fila, text="Cancelar", command=win.destroy, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO,
                      padx=8, pady=5).pack(side="right")
            auto_tamano(win, self)

        # ---------- DIALOGOS CLINICOS ----------

        def _dlg_consentimiento(self):
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero "
                                             "(Inicio o Documentos).")
                return
            win = tk.Toplevel(self)
            win.title("Consentimiento informado")
            win.configure(bg=CREMA)
            win.resizable(False, False)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=16, pady=16)
            marco.pack(fill="both", expand=True, padx=10, pady=10)
            tk.Label(marco, text="Consentimiento informado",
                     font=(FUENTE_SISTEMA, 14, "bold"), bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).pack(anchor="w")
            _epi_c11 = (self.datos_activos or {}).get("episodio_activo", "")
            tk.Label(marco,
                     text=f"Paciente: {nombre_mostrar(self.ruta_activa)}"
                          + (f"   ·   Episodio: {_epi_c11}" if _epi_c11 else ""),
                     bg=CREMA_CLARO, fg=AZUL, font=(FUENTE_SISTEMA, 10, "bold"),
                     wraplength=500, justify="left").pack(anchor="w", pady=4)
            tk.Label(marco, text="Procedimiento / acto clinico:",
                     bg=CREMA_CLARO, fg=TEXTO_OSCURO).pack(anchor="w", pady=(8, 2))
            e = tk.Entry(marco, font=(FUENTE_SISTEMA, 12), bg=BLANCO, width=60)
            e.pack(anchor="w", fill="x")
            e.focus_set()
            tk.Label(marco, text="Se generara un .txt en el episodio activo con tu firma "
                                 "y hash de integridad.",
                     bg=CREMA_CLARO, fg=GRIS_TERCIARIO, font=(FUENTE_SISTEMA, 9)
                     ).pack(anchor="w", pady=6)

            def generar():
                ok, msg, salida = generar_consentimiento(self.ruta_activa, e.get())
                if ok:
                    win.destroy()
                    ir_a_ruta(salida)
                    messagebox.showinfo("Consentimiento", f"{msg}\n"
                                        "Revisa, imprime para firma y guardalo.")
                else:
                    messagebox.showerror("Error", msg)

            tk.Button(marco, text="Generar consentimiento", command=generar,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 11, "bold"), pady=6).pack(anchor="e", pady=10)
            e.bind("<Return>", lambda ev: generar())
            auto_tamano(win, self)

        def _dlg_nom004(self):
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            win = tk.Toplevel(self)
            win.title("Historia clinica (estructura NOM-004)")
            win.configure(bg=CREMA)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=14, pady=14)
            marco.pack(fill="both", expand=True)
            tk.Label(marco, text="Historia clinica - NOM-004-SSA3-2012",
                     font=(FUENTE_SISTEMA, 14, "bold"), bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).grid(row=0, column=0, columnspan=2, sticky="w")
            _epi_c12 = (self.datos_activos or {}).get("episodio_activo", "")
            tk.Label(marco,
                     text=f"Paciente: {nombre_mostrar(self.ruta_activa)}"
                          + (f"   ·   Episodio: {_epi_c12}" if _epi_c12 else ""),
                     bg=CREMA_CLARO, fg=AZUL, font=(FUENTE_SISTEMA, 10, "bold")
                     ).grid(row=1, column=0, columnspan=2, sticky="w", pady=4)
            campos = {}
            # Documentos del episodio (ADR-0028/T14)
            _eid_c12 = (self.datos_activos or {}).get("episodio_activo", "")
            docs_ep = documentos_episodio(self.ruta_activa, _eid_c12)
            _hays = (f"Este episodio: {len(docs_ep['historias'])} historia(s) · "
                     f"{len(docs_ep['consentimientos'])} consentimiento(s) · "
                     f"Progreso Quimico {'si' if docs_ep['progreso'] else 'no'}")
            tk.Label(marco, text=_hays, bg=CREMA_CLARO, fg=GRIS_TEXTO,
                     font=(FUENTE_SISTEMA, 9), anchor="w"
                     ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 4))
            etiquetas = [
                ("heredofamiliares", "2. Antecedentes heredofamiliares"),
                ("patologicos", "2. Antecedentes personales patologicos"),
                ("gineco", "2. Antecedentes gineco-obstetricos"),
                ("exploracion", "3. Exploracion fisica"),
                ("estudios", "4. Estudios de gabinete/laboratorio"),
                ("diagnostico", "5. Diagnostico (lista CIE-10 si aplica)"),
                ("plan", "6. Plan y tratamiento"),
                ("evolucion", "7. Evolucion / notas de evolucion"),
            ]
            for i, (key, lab) in enumerate(etiquetas):
                fila = 3 + i
                tk.Label(marco, text=lab, bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                         font=(FUENTE_SISTEMA, 10), anchor="w"
                         ).grid(row=fila, column=0, sticky="nw", pady=3, padx=(0, 8))
                t = tk.Text(marco, height=2, width=64, font=(FUENTE_SISTEMA, 10),
                            bg=BLANCO, relief="solid", bd=1)
                t.grid(row=fila, column=1, sticky="we", pady=3)
                campos[key] = t
            marco.grid_columnconfigure(1, weight=1)

            def _vals():
                return {k: t.get("1.0", "end").strip() for k, t in campos.items()}

            def vista_previa():
                texto, err = historia_nom004_texto(self.ruta_activa, _vals())
                if texto is None:
                    messagebox.showerror("Error", err or "No se pudo armar",
                                         parent=win)
                    return
                vw = tk.Toplevel(win)
                vw.title("Vista previa — historia clinica")
                vw.configure(bg=CREMA)
                caja = tk.Frame(vw, bg=CREMA_CLARO, padx=10, pady=8)
                caja.pack(fill="both", expand=True)
                tk.Label(caja, text="Vista previa (aun NO guardada)",
                         font=(FUENTE_SISTEMA, 11, "bold"), bg=CREMA_CLARO,
                         fg=AZUL).pack(anchor="w")
                txt = tk.Text(caja, width=84, height=24, font=("Courier", 10),
                              bg=BLANCO, relief="solid", bd=1, wrap="none")
                txt.insert("1.0", texto)
                txt.config(state="disabled")
                txt.pack(fill="both", expand=True, pady=6)
                tk.Button(caja, text="Cerrar vista previa", command=vw.destroy,
                          bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                          activebackground=GRIS_APAGADO, padx=10, pady=6
                          ).pack(anchor="e")
                auto_tamano(vw, win)

            def abrir_ultima():
                h = docs_ep["historias"]
                if not h:
                    messagebox.showinfo("Aviso",
                                        "Este episodio aun no tiene historia generada.",
                                        parent=win)
                    return
                abrir_archivo(h[0])

            def generar():
                vals = _vals()
                ok, msg, salida = generar_historia_nom004(self.ruta_activa, vals)
                if ok:
                    win.destroy()
                    ir_a_ruta(salida)
                    messagebox.showinfo("Historia clinica",
                                        f"{msg}\nEl formato requiere firma profesional "
                                        "para validez oficial.")
                else:
                    messagebox.showerror("Error", msg)

            fila_btns = tk.Frame(marco, bg=CREMA_CLARO)
            fila_btns.grid(row=11, column=0, columnspan=2, sticky="we", pady=10)
            # ADR-0040: dictado de voz para campos narrativos
            tk.Button(fila_btns, text="Dictar (voz)", command=lambda: self._dictar(),
                      bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                      activebackground=GRIS_APAGADO, font=(FUENTE_SISTEMA, 9),
                      padx=8, pady=6).pack(side="left", padx=4)
            tk.Button(fila_btns, text="Vista previa", command=vista_previa,
                      bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                      activebackground=GRIS_APAGADO, font=(FUENTE_SISTEMA, 10),
                      padx=10, pady=6).pack(side="left", padx=4)
            tk.Button(fila_btns, text="Abrir ultima historia", command=abrir_ultima,
                      bg=GRIS_SECUNDARIO, fg=TEXTO_CLARO,
                      activebackground=GRIS_APAGADO, font=(FUENTE_SISTEMA, 10),
                      padx=10, pady=6).pack(side="left", padx=4)
            tk.Button(fila_btns, text="Generar historia clinica", command=generar,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 11, "bold"), pady=6
                      ).pack(side="right", padx=4)
            auto_tamano(win, self)

        # ---------- PAGINA: ADMINISTRACION ----------

        def _pagina_admin(self):
            admin = tk.Frame(self.pages_box, bg=CREMA_CLARO)
            admin.grid_rowconfigure(1, weight=1)
            admin.grid_columnconfigure(0, weight=1)
            # U36: contexto visible — nunca acciones sobre paciente a ciegas
            ctx_admin = tk.Frame(admin, bg=BLANCO, highlightthickness=1,
                                 highlightbackground=GRIS_BORDE)
            ctx_admin.grid(row=0, column=0, sticky="we", padx=16, pady=(8, 4))
            self.lbl_admin_sel = tk.Label(
                ctx_admin,
                text="Ninguna paciente seleccionada (selecciona en Inicio o Documentos)",
                font=(FUENTE_SISTEMA, 11, "bold"), bg=BLANCO, fg=ROJO_AVISO,
                anchor="w", padx=10, pady=7)
            self.lbl_admin_sel.pack(fill="x")
            cuerpo = tk.Frame(admin, bg=CREMA_CLARO)
            cuerpo.grid(row=1, column=0, sticky="nsew", padx=16, pady=4)
            col1 = tk.Frame(cuerpo, bg=CREMA_CLARO)
            col1.pack(side="left", fill="both", expand=True)
            col2 = tk.Frame(cuerpo, bg=CREMA_CLARO)
            col2.pack(side="left", fill="both", expand=True)
            botones = [
                (col1, "Abrir carpeta base", lambda: abrir_carpeta(BASE_DIR)),
                (col1, "Abrir expedientes", lambda: abrir_carpeta(EXPEDIENTES_DIR)),
                (col1, "Abrir plantillas", lambda: abrir_carpeta(PLANTILLAS_DIR)),
                (col1, "Abrir bitacora (_BITACORA.log)",
                 lambda: self._abrir_seguro(BITACORA_PATH)),
                (col1, "Abrir log de errores",
                 lambda: self._abrir_seguro(ERROR_LOG_PATH)),
                (col1, "Verificar integridad de bitacora", self._verificar_bitacora),
                (col1, "Cambiar autor registrado", self._cambiar_autor),
                (col1, "Cambiar / quitar PIN de apertura", self._cambiar_pin),
                (col1, "Exportar diagnostico (anonimo)", self._exportar_diagnostico),
                (col1, "Fusionar duplicados", self._fusionar_duplicados),
                (col1, "Aviso de privacidad de paciente", self._aviso_privacidad),
                (col2, "Abrir exportaciones", lambda: abrir_carpeta(EXPORTACIONES_DIR)),
                (col2, "Abrir documentos sin paciente",
                 lambda: abrir_carpeta(GENERAL_DIR)),
                (col2, "Abrir respaldos internos", lambda: abrir_carpeta(RESPALDOS_DIR)),
                (col1, "Importar pacientes (CSV)", self._importar_csv),
                (col2, "Exportar pacientes a Excel", lambda: self._exportar("xlsx")),
                (col2, "Exportar pacientes a CSV", lambda: self._exportar("csv")),
                (col2, "Exportar pacientes a PDF", lambda: self._exportar("pdf")),
                (col2, "Exportar expediente completo (zip)", self._exportar_expediente),
                (col2, "Respaldar / Verificar / Restaurar", self.dialogo_respaldo),
                (col2, "Zoom 100 %", lambda: self.aplicar_zoom(1.0)),
                (col2, "Zoom 125 %", lambda: self.aplicar_zoom(1.25)),
                (col2, "Zoom 150 %", lambda: self.aplicar_zoom(1.50)),
                (col2, "Ayuda", self._ayuda),
            ]
            for col, txt, cmd in botones:
                tk.Button(col, text=txt, command=cmd, bg=AZUL, fg=TEXTO_CLARO,
                          activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 10),
                          anchor="w", padx=12, pady=6).pack(fill="x", pady=3)
            tk.Label(admin,
                     text="Soporte normativo (consentimiento, bitacora, NOM-004, aviso de "
                          "privacidad), no certificacion oficial. Datos 100% locales.",
                     bg=CREMA_CLARO, fg=GRIS_TERCIARIO, font=(FUENTE_SISTEMA, 9),
                     wraplength=700, justify="left").grid(row=2, column=0,
                                                          sticky="w", padx=16, pady=(4, 10))
            admin.grid_remove()
            return admin

        def _actualizar_admin_sel(self):
            """Refresca la etiqueta de contexto de Administracion (U36)."""
            lbl = getattr(self, "lbl_admin_sel", None)
            if lbl is None:
                return
            if self.ruta_activa and self.datos_activos:
                lbl.config(text=f"Paciente seleccionada: "
                               f"{self.datos_activos.get('nombre_completo', '')}",
                           fg=VERDE_OK)
            else:
                lbl.config(text="Ninguna paciente seleccionada "
                                "(selecciona en Inicio o Documentos)",
                           fg=ROJO_AVISO)

        def _abrir_seguro(self, ruta):
            if Path(ruta).exists():
                abrir_archivo(ruta)
            else:
                messagebox.showinfo("Aviso", f"No existe aun: {Path(ruta).name}")

        def _verificar_bitacora(self):
            ok, msg = verificar_bitacora()
            if ok:
                messagebox.showinfo("Bitacora", msg)
            else:
                messagebox.showerror("Bitacora ALTERADA", msg)

        def _cambiar_autor(self):
            win = tk.Toplevel(self)
            win.title("Autor registrado")
            win.configure(bg=CREMA)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=14, pady=14)
            marco.pack(fill="both", expand=True)
            tk.Label(marco, text="Nombre que se registra en bitacora, "
                                 "consentimientos y documentos:",
                     bg=CREMA_CLARO, fg=TEXTO_OSCURO, wraplength=340,
                     justify="left").pack(anchor="w")
            e = tk.Entry(marco, font=(FUENTE_SISTEMA, 12), bg=BLANCO)
            e.insert(0, autor_actual())
            e.pack(fill="x", pady=8)

            def guardar():
                v = e.get().strip().upper()
                if v:
                    guardar_config({"autor": v})
                    for attr in ("lbl_autor", "lbl_nav_autor"):
                        try:
                            getattr(self, attr).config(text=f"Autor: {v}")
                        except Exception:
                            pass
                    messagebox.showinfo("Listo", f"Autor actualizado: {v}")
                    win.destroy()

            e.bind("<Return>", lambda ev: guardar())
            tk.Button(marco, text="Guardar", command=guardar, bg=AZUL, fg=TEXTO_CLARO,
                      activebackground=AZUL_HOVER).pack(anchor="e")
            e.focus_set()
            e.select_range(0, "end")
            auto_tamano(win, self)

        def _chequear_respaldo_programado(self):
            """A8: ejecuta respaldo programado si esta activo y toca."""
            try:
                cfg = leer_config()
                rp = cfg.get("respaldo_programado") or {}
                if not rp.get("activo"):
                    return
                hora_obj = rp.get("hora", "")
                if hora_obj and hora_obj == hora_ahora()[:5]:
                    destino = rp.get("destino", str(RESPALDOS_DIR / "programado"))
                    ok, msg = respaldar(destino)
                    if ok:
                        bitacora(f"Respaldo programado ejecutado: {destino}",
                                 autor=autor_actual())
            except Exception:
                pass
            self.after(3600000, self._chequear_respaldo_programado)

        def _drill_restaurar(self, destino=None):
            if not destino:
                messagebox.showinfo("Aviso", "Elige un destino primero.")
                return
            ok, msg = drill_restauracion(destino)
            if ok:
                self.toast("Drill OK ✓")
                messagebox.showinfo("Drill de restauracion", msg, parent=self)
            else:
                messagebox.showerror("Drill fallo", msg, parent=self)

        def _exportar_anonimo(self):
            """C-21: dialogo de vista previa + export anonimizado."""
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            win = tk.Toplevel(self)
            win.title("Historial clinico anonimizado")
            win.configure(bg=CREMA)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=14, pady=12)
            marco.pack(fill="both", expand=True)
            tk.Label(marco, text="Exportar historial anonimizado (C-21)",
                     font=(FUENTE_SISTEMA, 13, "bold"), bg=CREMA_CLARO, fg=AZUL
                     ).pack(anchor="w")
            # tabla retira/conserva
            tabla = tk.Frame(marco, bg=BLANCO, highlightthickness=1,
                             highlightbackground=GRIS_BORDE)
            tabla.pack(fill="x", pady=6)
            tk.Label(tabla, text="SE RETIRA", bg="#F8C9C4", fg=ROJO_AVISO,
                     font=(FUENTE_SISTEMA, 9, "bold"), padx=6, pady=3).grid(row=0, column=0, sticky="w")
            tk.Label(tabla, text="SE CONSERVA", bg="#DFF3E4", fg=VERDE_OK,
                     font=(FUENTE_SISTEMA, 9, "bold"), padx=6, pady=3).grid(row=0, column=1, sticky="w")
            retira = "nombre, CURP, clave, telefono, correo,\ndomicilio, NSS, fecha nacimiento,\nacompanante, cedula, testigos,\nnombres de medicos"
            conserva = "edad, sexo, ocupacion,\nescolaridad, estado civil, religion,\nsecciones A-M, serie de controles,\nlaboratorio, alergias, grupo,\nG-P-A, FUM/FPP/EG, alertas"
            tk.Label(tabla, text=retira, bg=BLANCO, fg=TEXTO_OSCURO,
                     font=(FUENTE_SISTEMA, 9), justify="left", padx=6, pady=4
                     ).grid(row=1, column=0, sticky="nw", padx=4, pady=4)
            tk.Label(tabla, text=conserva, bg=BLANCO, fg=TEXTO_OSCURO,
                     font=(FUENTE_SISTEMA, 9), justify="left", padx=6, pady=4
                     ).grid(row=1, column=1, sticky="nw", padx=4, pady=4)
            tabla.grid_columnconfigure(0, weight=1)
            tabla.grid_columnconfigure(1, weight=1)
            tk.Label(marco, text="Pregunta de analisis (opcional):", bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 10)).pack(anchor="w", pady=(6, 0))
            e_preg = tk.Text(marco, height=2, width=60, font=(FUENTE_SISTEMA, 10),
                             bg=BLANCO, relief="solid", bd=1)
            e_preg.insert("1.0", "Evaluar: alimentacion, rutina de vida, medicamentos, "
                                 "suplementos, ejercicio, aspectos criticos a vigilar.")
            e_preg.pack(fill="x", pady=3)
            var_csv = tk.BooleanVar(value=True)
            tk.Checkbutton(marco, text="Incluir CSV de controles", variable=var_csv,
                           bg=CREMA_CLARO, fg=TEXTO_OSCURO, selectcolor=BLANCO
                           ).pack(anchor="w", pady=4)
            lbl_estado = tk.Label(marco, text="", bg=CREMA_CLARO, fg=GRIS_TEXTO,
                                  font=(FUENTE_SISTEMA, 9), anchor="w", wraplength=500)
            lbl_estado.pack(fill="x")
            def _generar():
                preg = e_preg.get("1.0", "end").strip()
                ok, msg, salida = exportar_historia_anonima(self.ruta_activa, preg, var_csv.get())
                if ok:
                    self.toast("Exportado ✓")
                    lbl_estado.config(text=msg, fg=VERDE_OK)
                    if messagebox.askyesno("Exportado", f"{msg}\n\n¿Abrir carpeta?",
                                           parent=win):
                        abrir_carpeta(EXPORTACIONES_DIR)
                        win.destroy()
                else:
                    lbl_estado.config(text=msg, fg=ROJO_AVISO)
                    messagebox.showerror("Export anonimizado", msg, parent=win)
            pie = tk.Frame(marco, bg=CREMA_CLARO)
            pie.pack(fill="x", pady=(8, 0))
            tk.Button(pie, text="Cerrar", command=win.destroy, bg=GRIS_SECUNDARIO,
                      fg=TEXTO_CLARO, activebackground=GRIS_APAGADO, padx=10, pady=6
                      ).pack(side="right")
            tk.Button(pie, text="Generar export anonimizado", command=_generar,
                      bg=AZUL, fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 11, "bold"), padx=12, pady=6
                      ).pack(side="right", padx=6)
            auto_tamano(win, self)

        def _importar_csv(self):
            """ADR-0039: importa pacientes desde CSV con mapeo automatico."""
            ruta = filedialog.askopenfilename(
                title="Selecciona el CSV",
                filetypes=[("CSV", "*.csv"), ("Todos", "*.*")])
            if not ruta:
                return
            try:
                with open(ruta, encoding="utf-8-sig") as f:
                    encabezados = next(csv.reader(f))
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el CSV: {e}")
                return
            mapeo = detectar_mapeo_csv(encabezados)
            win = tk.Toplevel(self)
            win.title("Importar pacientes (CSV)")
            win.configure(bg=CREMA)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=14, pady=12)
            marco.pack(fill="both", expand=True)
            tk.Label(marco, text=f"CSV: {Path(ruta).name} ({len(encabezados)} columnas)",
                     font=(FUENTE_SISTEMA, 11, "bold"), bg=CREMA_CLARO, fg=AZUL
                     ).pack(anchor="w")
            tk.Label(marco, text="Ajusta el mapeo si es necesario:",
                     bg=CREMA_CLARO, fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 10)
                     ).pack(anchor="w", pady=(6, 0))
            vars_m = {}
            campos = ["ap_paterno", "ap_materno", "nombre", "fecha_nacimiento",
                      "curp", "telefono", "correo", "medico_remitente"]
            for campo in campos:
                fila = tk.Frame(marco, bg=CREMA_CLARO)
                fila.pack(fill="x", pady=1)
                tk.Label(fila, text=campo.replace("_", " ").capitalize(),
                         bg=CREMA_CLARO, fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 9),
                         width=18, anchor="w").pack(side="left")
                var = tk.StringVar(value=mapeo.get(campo, ""))
                vars_m[campo] = var
                ttk.Combobox(fila, textvariable=var, values=encabezados,
                             width=20, font=(FUENTE_SISTEMA, 9)).pack(side="left", padx=4)
            def _importar():
                m = {k: v.get() for k, v in vars_m.items() if v.get()}
                ok, msg, res = importar_pacientes_csv(ruta, m)
                if ok:
                    self.toast(f"Importación: {res['importados']} pacientes")
                    self.estado.dispatch("paciente:cambiada", {"recientes", "home", "admin"})
                    self.estado.flush()
                    messagebox.showinfo("Importación completa", msg, parent=win)
                    win.destroy()
                else:
                    messagebox.showerror("Error", msg, parent=win)
            tk.Button(marco, text="Importar", command=_importar, bg=AZUL, fg=TEXTO_CLARO,
                      activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 11, "bold"),
                      padx=12, pady=6).pack(anchor="e", pady=10)
            auto_tamano(win, self)

        def _exportar_diagnostico(self):
            ok, msg, salida = exportar_diagnostico()
            if ok:
                self.toast("Diagnostico exportado ✓")
                if messagebox.askyesno("Diagnostico", f"{msg}\n\n¿Abrir?", parent=self):
                    abrir_archivo(salida)
            else:
                messagebox.showerror("Error", msg, parent=self)

        def _fusionar_duplicados(self):
            """A4: fusionar dos expedientes por nombre+fecha de nacimiento."""
            d = tk.Toplevel(self)
            d.title("Fusionar duplicados")
            d.configure(bg=CREMA)
            marco = tk.Frame(d, bg=CREMA_CLARO, padx=14, pady=14)
            marco.pack(fill="both", expand=True)
            tk.Label(marco, text="Selecciona el expediente ORIGEN (se movera a _FUSIONADOS):",
                     bg=CREMA_CLARO, fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 10),
                     wraplength=400, justify="left").pack(anchor="w")
            lb_o = tk.Listbox(marco, height=6, font=(FUENTE_SISTEMA, 9), bg=BLANCO)
            lb_o.pack(fill="x", pady=4)
            rutas_o = []
            for nom, ruta in listar_pacientes():
                lb_o.insert("end", nom)
                rutas_o.append(ruta)
            tk.Label(marco, text="Destino (se conserva):", bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO, font=(FUENTE_SISTEMA, 10)).pack(anchor="w", pady=(8, 0))
            lb_d = tk.Listbox(marco, height=6, font=(FUENTE_SISTEMA, 9), bg=BLANCO)
            lb_d.pack(fill="x", pady=4)
            rutas_d = []
            for nom, ruta in listar_pacientes():
                lb_d.insert("end", nom)
                rutas_d.append(ruta)
            tk.Label(marco, text="Motivo:", bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                     font=(FUENTE_SISTEMA, 10)).pack(anchor="w", pady=(8, 0))
            e_m = tk.Entry(marco, font=(FUENTE_SISTEMA, 10), bg=BLANCO)
            e_m.pack(fill="x")
            def _fusionar():
                io = lb_o.curselection()
                id_ = lb_d.curselection()
                if not io or not id_:
                    messagebox.showwarning("Aviso", "Selecciona origen y destino.", parent=d)
                    return
                if io[0] == id_[0]:
                    messagebox.showwarning("Aviso", "Origen y destino no pueden ser iguales.", parent=d)
                    return
                ok, msg = fusionar_paciente(rutas_o[io[0]], rutas_d[id_[0]],
                                            e_m.get().strip() or "sin motivo")
                if ok:
                    self.toast("Fusionado ✓")
                    self.estado.dispatch("paciente:cambiada", {"recientes", "home", "admin"})
                    self.estado.flush()
                    d.destroy()
                else:
                    messagebox.showerror("Error", msg, parent=d)
            tk.Button(marco, text="Fusionar", command=_fusionar, bg=ROJO_AVISO,
                      fg=TEXTO_CLARO, activebackground="#C43A4B",
                      font=(FUENTE_SISTEMA, 11, "bold"), padx=12, pady=6).pack(anchor="e", pady=10)
            auto_tamano(d, self)

        def _cambiar_pin(self):
            """P2: activar/cambiar el PIN de apertura (hash + sal, nunca claro)."""
            win = tk.Toplevel(self)
            win.title("PIN de apertura")
            win.configure(bg=CREMA)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=14, pady=14)
            marco.pack(fill="both", expand=True)
            activo = pin_configurado()
            tk.Label(marco,
                     text=("PIN activo. Cambia la contrasena (4-6 digitos):"
                           if activo else
                           "Sin PIN. Crea uno de 4-6 digitos (vacío = no crear):"),
                     bg=CREMA_CLARO, fg=TEXTO_OSCURO, wraplength=360,
                     justify="left").pack(anchor="w")
            e_actual = None
            if activo:
                tk.Label(marco, text="PIN actual:", bg=CREMA_CLARO,
                         fg=TEXTO_OSCURO).pack(anchor="w", pady=(8, 0))
                e_actual = tk.Entry(marco, font=(FUENTE_SISTEMA, 12), bg=BLANCO, show="*")
                e_actual.pack(fill="x")
            tk.Label(marco, text="Nuevo PIN (4-6 digitos):", bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).pack(anchor="w", pady=(8, 0))
            e_nuevo = tk.Entry(marco, font=(FUENTE_SISTEMA, 12), bg=BLANCO, show="*")
            e_nuevo.pack(fill="x")
            tk.Label(marco, text="Repite el nuevo PIN:", bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).pack(anchor="w", pady=(8, 0))
            e_rep = tk.Entry(marco, font=(FUENTE_SISTEMA, 12), bg=BLANCO, show="*")
            e_rep.pack(fill="x")

            def guardar(*_a):
                if activo and not verificar_pin(e_actual.get()):
                    messagebox.showwarning("PIN actual incorrecto",
                                           "El PIN actual no coincide.", parent=win)
                    return
                nuevo = e_nuevo.get().strip()
                if not nuevo:
                    messagebox.showinfo("Aviso",
                                        "Escribe un PIN de 4 a 6 digitos.", parent=win)
                    return
                if not (4 <= len(nuevo) <= 6 and nuevo.isdigit()):
                    messagebox.showwarning("PIN invalido",
                                           "Debe tener 4 a 6 digitos.", parent=win)
                    return
                if nuevo != e_rep.get():
                    messagebox.showwarning("No coincide",
                                           "La repeticion no coincide.", parent=win)
                    return
                guardar_pin(nuevo)
                messagebox.showinfo("Listo",
                                    "PIN guardado. Se pedira al abrir el sistema.",
                                    parent=win)
                win.destroy()

            def quitar(*_a):
                if activo and not verificar_pin(e_actual.get()):
                    messagebox.showwarning("PIN actual incorrecto",
                                           "El PIN actual no coincide.", parent=win)
                    return
                if messagebox.askyesno("Quitar PIN",
                                       "El sistema abrira sin PIN. ¿Continuar?",
                                       parent=win):
                    quitar_pin()
                    messagebox.showinfo("Listo", "PIN eliminado.", parent=win)
                    win.destroy()

            fila = tk.Frame(marco, bg=CREMA_CLARO)
            fila.pack(fill="x", pady=(12, 0))
            tk.Button(fila, text="Guardar PIN", command=guardar, bg=AZUL,
                      fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                      font=(FUENTE_SISTEMA, 11, "bold")).pack(side="left")
            if activo:
                tk.Button(fila, text="Quitar PIN", command=quitar, bg=GRIS_SECUNDARIO,
                          fg=TEXTO_CLARO, activebackground=GRIS_APAGADO
                          ).pack(side="left", padx=8)
            for e in (e_actual, e_nuevo, e_rep):
                if e is not None:
                    e.bind("<Return>", guardar)
            (e_actual or e_nuevo).focus_set()
            auto_tamano(win, self)

        def _aviso_privacidad(self):
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            ok, msg, salida = generar_aviso_privacidad(self.ruta_activa)
            if ok:
                abrir_archivo(salida)

        def _exportar(self, formato):
            if formato == "xlsx":
                inicial = EXPORTACIONES_DIR / f"PACIENTES_{timestamp()}.xlsx"
                tipos = [("Excel", "*.xlsx")]
            elif formato == "csv":
                inicial = EXPORTACIONES_DIR / f"PACIENTES_{timestamp()}.csv"
                tipos = [("CSV", "*.csv")]
            else:
                inicial = EXPORTACIONES_DIR / f"PACIENTES_{timestamp()}.pdf"
                tipos = [("PDF", "*.pdf")]
            destino = filedialog.asksaveasfilename(
                title="Guardar exportacion", initialdir=str(EXPORTACIONES_DIR),
                initialfile=str(inicial.name), defaultextension=inicial.suffix,
                filetypes=tipos)
            if not destino:
                return
            if formato == "xlsx":
                ok, msg, salida = exportar_pacientes_xlsx(destino)
            elif formato == "csv":
                ok, msg, salida = exportar_pacientes_csv(destino)
            else:
                ok, msg, salida = exportar_pacientes_pdf(destino)
            if ok:
                if messagebox.askyesno("Exportado", f"{msg}\n¿Abrirlo?"):
                    abrir_archivo(salida)
            else:
                messagebox.showerror("Error", msg)

        def _exportar_expediente(self):
            if not self.ruta_activa:
                messagebox.showinfo("Aviso", "Selecciona una paciente primero.")
                return
            nombre = Path(self.ruta_activa).name
            destino = filedialog.asksaveasfilename(
                title="Exportar expediente", initialdir=str(EXPORTACIONES_DIR),
                initialfile=f"EXPEDIENTE_{nombre}_{timestamp()}.zip",
                defaultextension=".zip", filetypes=[("ZIP", "*.zip")])
            if not destino:
                return
            ok, msg, salida = exportar_expediente(self.ruta_activa, destino)
            if ok:
                messagebox.showinfo("Exportado", msg)
            else:
                messagebox.showerror("Error", msg)

        # ---------- RESPALDO ----------

        def dialogo_respaldo(self):
            win = tk.Toplevel(self)
            win.title("Respaldo del sistema")
            win.configure(bg=CREMA)
            win.resizable(False, False)
            marco = tk.Frame(win, bg=CREMA_CLARO, padx=16, pady=16)
            marco.pack(fill="both", expand=True)
            tk.Label(marco, text="Respaldo / Verificacion / Restauracion",
                     font=(FUENTE_SISTEMA, 14, "bold"), bg=CREMA_CLARO,
                     fg=TEXTO_OSCURO).pack(anchor="w")
            tk.Label(marco,
                     text="5 componentes de datos + manifiesto SHA-256 · "
                          "cifrado opcional con contraseña",
                     font=(FUENTE_SISTEMA, 9), bg=CREMA_CLARO, fg=GRIS_TEXTO
                     ).pack(anchor="w", pady=(2, 6))
            modo = tk.StringVar(value="respaldar")
            for val, txt in [("respaldar", "Respaldar (copiar datos al destino)"),
                             ("verificar", "Verificar un respaldo (SHA-256)"),
                             ("restaurar", "Restaurar desde un respaldo (REEMPLAZA datos actuales)")]:
                tk.Radiobutton(marco, text=txt, variable=modo, value=val,
                               bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                               selectcolor=BLANCO, font=(FUENTE_SISTEMA, 10)
                               ).pack(anchor="w", pady=3)
            fila = tk.Frame(marco, bg=CREMA_CLARO)
            fila.pack(fill="x", pady=10)
            e_dest = tk.Entry(fila, font=(FUENTE_SISTEMA, 11), bg=BLANCO)
            e_dest.pack(side="left", fill="x", expand=True, padx=(0, 6))

            def examinar():
                d = filedialog.askdirectory(title="Carpeta destino")
                if d:
                    e_dest.delete(0, "end")
                    e_dest.insert(0, d)

            tk.Button(fila, text="Examinar...", command=examinar, bg=AZUL,
                      fg=TEXTO_CLARO, activebackground=AZUL_HOVER).pack(side="left")
            # P1 (T9b): contraseña opcional — cifra al respaldar; exige al
            # verificar/restaurar si el respaldo viene cifrado
            fila_clave = tk.Frame(marco, bg=CREMA_CLARO)
            fila_clave.pack(fill="x", pady=(4, 0))
            tk.Label(fila_clave, text="Contraseña (opcional para cifrar):",
                     bg=CREMA_CLARO, fg=TEXTO_OSCURO,
                     font=(FUENTE_SISTEMA, 10)).pack(anchor="w")
            e_clave = tk.Entry(fila_clave, font=(FUENTE_SISTEMA, 11), bg=BLANCO, show="*")
            e_clave.pack(fill="x", pady=2)
            if not CRYPTO_DISPONIBLE:
                tk.Label(fila_clave,
                         text="Aviso: cryptography no instalada → el cifrado no "
                              "está disponible (integridad SHA-256 siempre activa).",
                         bg=CREMA_CLARO, fg=ROJO_AVISO, font=(FUENTE_SISTEMA, 8),
                         wraplength=360, justify="left").pack(anchor="w")
            lbl_est = tk.Label(marco, text="", bg=CREMA_CLARO, fg=AZUL,
                               font=(FUENTE_SISTEMA, 10))
            lbl_est.pack(anchor="w", pady=4)

            def ejecutar():
                destino = e_dest.get().strip()
                if not destino:
                    messagebox.showwarning("Aviso", "Elige un destino.")
                    return
                clave = e_clave.get() or None
                m = modo.get()
                if m == "restaurar":
                    if not messagebox.askyesno(
                            "RESTAURAR",
                            "Esto REEMPLAZA todos tus expedientes actuales\n"
                            "por los del respaldo elegido.\n\n"
                            "Los datos actuales quedaran movidos a _RESPALDOS\n"
                            "como copia de seguridad previa.\n\n"
                            "¿Continuar?"):
                        return
                btn_ejec.config(state="disabled")
                lbl_est.config(text="Procesando... no cierres la ventana.")
                self.progreso.grid()
                self.progreso.start(12)
                self._respaldo_activo = True
                q = queue.Queue()
                self._cola_respaldo = q

                def worker():
                    if m == "respaldar":
                        r = respaldar(destino, clave)
                    elif m == "verificar":
                        r = verificar_respaldo(destino, clave)
                    else:
                        r = self._restaurar_con_previo(destino, clave)
                    q.put(r)

                threading.Thread(target=worker, daemon=True).start()
                self._poll_respaldo(q, win, lbl_est, btn_ejec, m, destino)

            btn_ejec = tk.Button(marco, text="Ejecutar", command=ejecutar, bg=AZUL,
                                 fg=TEXTO_CLARO, activebackground=AZUL_HOVER,
                                 font=(FUENTE_SISTEMA, 11, "bold"), pady=6)
            btn_ejec.pack(anchor="e", pady=8)
            auto_tamano(win, self)

        def _restaurar_con_previo(self, destino, clave=None):
            marca = timestamp()
            previo = RESPALDOS_DIR / f"ANTES_DE_RESTAURAR_{marca}"
            for comp in COMPONENTES_DATOS:
                actual = BASE_DIR / comp
                if actual.exists():
                    previo.mkdir(parents=True, exist_ok=True)
                    if actual.is_dir():
                        shutil.copytree(actual, previo / comp, dirs_exist_ok=True)
                    else:
                        shutil.copy2(actual, previo / comp)
            ok, msg = restaurar_respaldo(destino, clave)
            if ok:
                bitacora(f"Copia previa guardada en {previo}", autor=autor_actual())
                msg += f"\n\nCopia previa de seguridad: {previo.name}"
            return ok, msg

        def _poll_respaldo(self, q, win, lbl, btn, modo, destino):
            try:
                ok, msg = q.get_nowait()
            except queue.Empty:
                self.after(250, lambda: self._poll_respaldo(q, win, lbl, btn, modo, destino))
                return
            self.progreso.stop()
            self.progreso.grid_remove()
            self._respaldo_activo = False
            btn.config(state="normal")
            lbl.config(text="")
            if modo == "verificar":
                (messagebox.showinfo if ok else messagebox.showerror)("Verificacion", msg)
            elif ok:
                messagebox.showinfo("Listo", msg)
                if modo == "respaldar" and messagebox.askyesno(
                        "Respaldado", "¿Verificar el respaldo ahora?"):
                    ok2, msg2 = verificar_respaldo(destino)
                    (messagebox.showinfo if ok2 else messagebox.showerror)(
                        "Verificacion", msg2)
            else:
                messagebox.showerror("Error", msg)


# ===== 11 PUNTO DE ENTRADA =====

def pedir_pin(intentos_max=5):
    """P2: gate de apertura. True = acceso concedido (o sin PIN configurado)."""
    if not pin_configurado():
        return True
    raiz = tk.Tk()
    raiz.withdraw()
    win = tk.Toplevel(raiz)
    win.title("Acceso restringido")
    win.configure(bg=CREMA)
    win.resizable(False, False)
    marco = tk.Frame(win, bg=CREMA_CLARO, padx=16, pady=14)
    marco.pack(fill="both", expand=True)
    tk.Label(marco, text="Ingresa el PIN de apertura",
             font=(FUENTE_SISTEMA, 13, "bold"), bg=CREMA_CLARO,
             fg=TEXTO_OSCURO).pack(anchor="w")
    e = tk.Entry(marco, font=(FUENTE_SISTEMA, 14), show="*", bg=BLANCO, justify="center")
    e.pack(fill="x", pady=10)
    lbl = tk.Label(marco, text=f"Intentos disponibles: {intentos_max}",
                   bg=CREMA_CLARO, fg=GRIS_TEXTO, font=(FUENTE_SISTEMA, 9))
    lbl.pack(anchor="w")
    estado = {"ok": False, "restantes": intentos_max}

    def cerrar_denegado(demasiados=False):
        if demasiados:
            messagebox.showerror("Bloqueado",
                                 "Demasiados intentos. El sistema se cerrara\n"
                                 "vuelve a intentar en 30 segundos.")
            time.sleep(30)
        win.destroy()
        raiz.destroy()

    def intentar(*_a):
        if verificar_pin(e.get()):
            estado["ok"] = True
            win.destroy()
            raiz.destroy()
            return
        estado["restantes"] -= 1
        e.delete(0, "end")
        if estado["restantes"] <= 0:
            cerrar_denegado(demasiados=True)
        else:
            lbl.config(text=f"PIN incorrecto. Intentos: {estado['restantes']}")

    tk.Button(marco, text="Entrar", command=intentar, bg=AZUL, fg=TEXTO_CLARO,
              activebackground=AZUL_HOVER, font=(FUENTE_SISTEMA, 11, "bold")
              ).pack(anchor="e")
    e.bind("<Return>", intentar)
    e.focus_set()
    win.protocol("WM_DELETE_WINDOW", lambda: (win.destroy(), raiz.destroy()))
    win.grab_set()
    raiz.wait_window(win)
    return estado["ok"]


def _avisar_instancia_reciente():
    """P7: aviso best-effort si otra instancia reciente existe. True = seguir."""
    reciente, msg = bloqueo_reciente()
    if not reciente:
        return True
    raiz = tk.Tk()
    raiz.withdraw()
    try:
        seguir = messagebox.askyesno(
            "Otra instancia detectada",
            f"{msg}\n\nEl sistema es monousuario por PC: dos instancias a la vez "
            "pueden pisar datos.\n\n¿Continuar de todos modos?")
    finally:
        raiz.destroy()
    return seguir


def main():
    if not TK_DISPONIBLE:
        print("ERROR: esta instalacion de Python no incluye tkinter.")
        print("En Windows reinstala Python marcando 'tcl/tk'.")
        print("En Linux: sudo apt install python3-tk")
        print("En macOS usa el Python de python.org (ya trae Tk); si usaste")
        print("brew: brew install python-tk")
        sys.exit(1)
    inicializar_estructura()
    _diag = diagnostico_arranque()
    _problemas = [l for l in _diag if "ERROR" in l or "corrupto" in l or "Recuperado" in l or "pendiente" in l.lower()]
    if _problemas:
        raiz0 = tk.Tk()
        raiz0.withdraw()
        try:
            messagebox.showwarning(
                "Diagnostico de arranque",
                "Se encontraron puntos que requieren atencion:\n\n" +
                "\n".join(_problemas[:15]) +
                "\n\nDetalle completo en _DIAGNOSTICO.log")
        finally:
            raiz0.destroy()
    if not _avisar_instancia_reciente():
        return
    if not pedir_pin():
        return
    crear_bloqueo()
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
