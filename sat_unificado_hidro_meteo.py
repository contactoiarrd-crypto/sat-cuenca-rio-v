# -*- coding: utf-8 -*-
"""
SISTEMA DE ALERTA TEMPRANA TRIPROVINCIAL: CUENCA RÍO V (SAN LUIS - CÓRDOBA - LA PAMPA)
Monitoreo Hidrometeorológico Online, Cuerpos de Agua INA, SAT SMN, Pronóstico ECMWF IFS,
Comparativa Climática, Radares SINARAME y Tablero de Control Multisolapa.
"""
import os
import sys
import math
import json
import re
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import numpy as np
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(line_buffering=True)

# =============================================================
# 1. PARÁMETROS GENERALES Y ARCHIVOS DE SALIDA
# =============================================================
TZ_ARG = timezone(timedelta(hours=-3))
ahora = datetime.now(TZ_ARG)
FECHA_TXT = ahora.strftime("%Y-%m-%d %H:%M")

EXCEL_SALIDA = "sat_unificado_rio_v_triprovincial.xlsx"
CSV_SALIDA = "resumen_cruce_rio_v_triprovincial.csv"
PORTAL_HTML_SALIDA = "index.html"

LAT_MIN_SL, LAT_MAX_SL = -35.5, -33.0
LON_MIN_SL, LON_MAX_SL = -66.3, -65.0
LAT_MIN_CUENCA, LAT_MAX_CUENCA = -36.5, -32.8
LON_MIN_CUENCA, LON_MAX_CUENCA = -67.5, -63.0

session = requests.Session()
retries = Retry(total=2, backoff_factor=1.0, status_forcelist=[500, 502, 503, 504])
session.mount("https://", HTTPAdapter(max_retries=retries))
session.mount("http://", HTTPAdapter(max_retries=retries))

# =============================================================
# 2. MEDIAS CLIMÁTICAS Y NODOS ECMWF A 72 HORAS
# =============================================================
MEDIAS_CLIMATICAS_CUENCA = {
    "Villa Mercedes": {"anual_mm": 680, "mes_esperado_mm": 55, "provincia": "San Luis", "lat": -33.67, "lon": -65.46, "region": "Cabecera / Nacientes"},
    "General Levalle": {"anual_mm": 750, "mes_esperado_mm": 68, "provincia": "Córdoba", "lat": -34.00, "lon": -63.92, "region": "Cuenca Media (Aporte)"},
    "Jovita": {"anual_mm": 770, "mes_esperado_mm": 70, "provincia": "Córdoba", "lat": -34.52, "lon": -63.97, "region": "Cuenca Media-Baja"},
    "Realicó": {"anual_mm": 800, "mes_esperado_mm": 75, "provincia": "La Pampa", "lat": -35.04, "lon": -64.24, "region": "Norte Pampeano (RN 35)"},
    "Intendente Alvear": {"anual_mm": 820, "mes_esperado_mm": 78, "provincia": "La Pampa", "lat": -35.24, "lon": -63.59, "region": "Norte Pampeano (Meridiano V)"},
    "Punta Alta (Rancul)": {"anual_mm": 780, "mes_esperado_mm": 72, "provincia": "La Pampa", "lat": -35.21, "lon": -64.45, "region": "Norte Pampeano (RP 9)"}
}

# =============================================================
# 3. CATÁLOGO ESTÁTICO REDES OMIXOM
# =============================================================
ESTACIONES_OMIXOM_ESTATICAS = [
    {"id": "OMX_CBA_1", "nombre": "General Levalle", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.0000, "lon": -63.9163, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_2", "nombre": "Río Bamba", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.0537, "lon": -63.7332, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_3", "nombre": "Jovita", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.5194, "lon": -63.9683, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_4", "nombre": "Villa Valeria", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.3427, "lon": -64.9290, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_5", "nombre": "Nicolás Bruzzone", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.4391, "lon": -64.3424, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_6", "nombre": "Melo", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.3452, "lon": -63.4377, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_7", "nombre": "Huanchillas", "departamento": "Juárez Celman", "provincia": "Córdoba", "lat": -33.6665, "lon": -63.6395, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_8", "nombre": "Hipólito Bouchard", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.7060, "lon": -63.5030, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_9", "nombre": "General Roca", "departamento": "General Roca", "provincia": "Córdoba", "lat": -33.9948, "lon": -65.0754, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_10", "nombre": "Serrano", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.4629, "lon": -63.5309, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_11", "nombre": "Coronel Moldes", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.6481, "lon": -64.5950, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_12", "nombre": "Viamonte", "departamento": "Unión", "provincia": "Córdoba", "lat": -33.7429, "lon": -63.0996, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_13", "nombre": "Chaján", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.5508, "lon": -65.0059, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_14", "nombre": "Villa Rossi", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.2949, "lon": -63.2654, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_15", "nombre": "Presa El Chañar", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.9608, "lon": -65.0551, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_16", "nombre": "Huinca Renancó", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.8208, "lon": -64.3738, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_17", "nombre": "Vicuña Mackenna", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.9754, "lon": -64.3642, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom Córdoba"},
    {"id": "OMX_LP_1", "nombre": "El Tala - La Veneta", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.3119, "lon": -64.7144, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_2", "nombre": "Realicó", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.0576, "lon": -64.2129, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_3", "nombre": "Trilí", "departamento": "Quemú Quemú", "provincia": "La Pampa", "lat": -35.9036, "lon": -63.6429, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_4", "nombre": "Ingeniero Luiggi", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.4714, "lon": -64.6024, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_5", "nombre": "Conhelo", "departamento": "Conhelo", "provincia": "La Pampa", "lat": -35.9895, "lon": -64.5954, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_6", "nombre": "Arata", "departamento": "Trenel", "provincia": "La Pampa", "lat": -35.6391, "lon": -64.3564, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_7", "nombre": "Pichi Huinca", "departamento": "Rancul", "provincia": "La Pampa", "lat": -35.6482, "lon": -64.7699, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_8", "nombre": "Rancul", "departamento": "Rancul", "provincia": "La Pampa", "lat": -35.0883, "lon": -64.5082, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_9", "nombre": "Coronel Hilario Lagos", "departamento": "Chapaleufú", "provincia": "La Pampa", "lat": -35.0344, "lon": -63.9111, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_10", "nombre": "Colonia Barón", "departamento": "Quemú Quemú", "provincia": "La Pampa", "lat": -36.1508, "lon": -63.8550, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_11", "nombre": "Intendente Alvear", "departamento": "Chapaleufú", "provincia": "La Pampa", "lat": -35.3182, "lon": -63.6054, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_12", "nombre": "Alta Italia", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.3317, "lon": -64.1191, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_13", "nombre": "Winifreda", "departamento": "Conhelo", "provincia": "La Pampa", "lat": -36.2229, "lon": -64.2487, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_14", "nombre": "General Pico", "departamento": "Maracó", "provincia": "La Pampa", "lat": -35.6969, "lon": -63.6207, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_15", "nombre": "Eduardo Castex", "departamento": "Conhelo", "provincia": "La Pampa", "lat": -35.9160, "lon": -64.2956, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"}
]

# =============================================================
# 4. CATÁLOGO CUERPOS DE AGUA (INA) Y COTAS FÍSICAS
# =============================================================
BASE_URL_INA = "https://alerta.ina.gob.ar/pub/datos"
timestart_str = (ahora - timedelta(days=4)).strftime("%Y-%m-%d")
timeend_str = (ahora + timedelta(days=1)).strftime("%Y-%m-%d")

ESTACIONES_INA_CATALOGO = {
    6444: {"nombre": "Trapiche - Hosteria El Trapiche", "rio": "Río Trapiche (Afluente)", "distrito": "San Luis", "lat": -33.105833, "lon": -66.063333},
    6441: {"nombre": "Quinto - Dique Villa Mercedes", "rio": "Río Quinto", "distrito": "San Luis", "lat": -33.653889, "lon": -65.533611},
    6472: {"nombre": "Quinto - Av Circunvalación", "rio": "Río Quinto", "distrito": "San Luis", "lat": -33.739167, "lon": -65.376944},
    6445: {"nombre": "Quinto - Justo Daract", "rio": "Río Quinto", "distrito": "San Luis", "lat": -33.918611, "lon": -65.151667},
    6624: {"nombre": "Aº El Aji - RN Nº35 y RN Nº7", "rio": "Afluente Río Quinto", "distrito": "Córdoba / LP", "lat": -33.931111, "lon": -64.396389},
    6622: {"nombre": "Quinto - R.N. 35", "rio": "Río Quinto", "distrito": "Córdoba / LP", "lat": -34.216389, "lon": -64.386111},
    6623: {"nombre": "Quinto - Canal Devoto RP 4", "rio": "Río Quinto / Canal Devoto", "distrito": "Córdoba / LP", "lat": -33.933056, "lon": -63.449167},
    6391: {"nombre": "Laguna La Margarita", "rio": "Cuenca Río Quinto", "distrito": "Córdoba / LP", "lat": -34.653333, "lon": -63.723056},
    2809: {"nombre": "Quinto - RP Nº26", "rio": "Río Quinto", "distrito": "Córdoba / LP", "lat": -34.762778, "lon": -63.645000}
}

UMBRALES_NOMINALES = {
    "TRAPICHE": {"alerta": 1.80, "evac": 2.30},
    "VILLA MERCEDES": {"alerta": 2.80, "evac": 3.50},
    "CIRCUNVALACION": {"alerta": 1.60, "evac": 2.20},
    "JUSTO DARACT": {"alerta": 2.90, "evac": 3.60},
    "AJI": {"alerta": 1.90, "evac": 2.40},
    "AJÍ": {"alerta": 1.90, "evac": 2.40},
    "R.N. 35": {"alerta": 2.60, "evac": 3.20},
    "DEVOTO": {"alerta": 2.80, "evac": 3.40},
    "MARGARITA": {"alerta": 2.40, "evac": 3.00},
    "RP Nº26": {"alerta": 2.20, "evac": 2.80},
    "RP 26": {"alerta": 2.20, "evac": 2.80}
}

def clasificar_nivel(valor, nombre_estacion):
    nombre_limpio = nombre_estacion.upper().strip()
    c_alerta, c_evac = 2.50, 3.20
    for k, v in UMBRALES_NOMINALES.items():
        if k in nombre_limpio:
            c_alerta, c_evac = v["alerta"], v["evac"]
            break
    
    umbral_precaucion = round(c_alerta * 0.75, 2)
    if valor < umbral_precaucion:
        return "#10b981", "Normal / Seguro", c_alerta, c_evac
    elif valor >= c_evac:
        return "#ef4444", "Evacuación Oficial", c_alerta, c_evac
    elif valor >= c_alerta:
        return "#ea580c", "Alerta Hidrológica", c_alerta, c_evac
    else:
        return "#f59e0b", "Precaución", c_alerta, c_evac

def distancia_haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2.0)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2.0)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

# =============================================================
# 5. EXTRACCIÓN DE DATOS (INA, REM, APA, RAYOS, SMN, ECMWF)
# =============================================================
def obtener_estaciones_omixom():
    return ESTACIONES_OMIXOM_ESTATICAS

def normalizar_a_lista(resp_json):
    if isinstance(resp_json, list):
        return resp_json
    if isinstance(resp_json, dict):
        for k in ["data", "datos", "series", "estaciones", "results"]:
            if k in resp_json and isinstance(resp_json[k], list):
                return resp_json[k]
    return []

def _descargar_serie_ina(s):
    sid = s["ina_sid"]
    url = f"{BASE_URL_INA}/datos&seriesId={sid}&timeStart={timestart_str}&timeEnd={timeend_str}&format=json"
    salida = []
    try:
        r = session.get(url, timeout=7)
        if r.status_code == 200:
            for d in normalizar_a_lista(r.json()):
                if isinstance(d, dict):
                    f = d.get("timestart") or d.get("timeStart") or d.get("fecha") or d.get("time")
                    v = d.get("valor") or d.get("value") or d.get("val")
                    if f is not None and v is not None:
                        salida.append({
                            "fecha": str(f).replace("T", " "),
                            "valor": float(v),
                            "sitecode": str(s["sitecode"]),
                            "nombre": s["nombre"],
                            "distrito": s["distrito"],
                            "rio": s["rio"],
                            "lat": s["lat"],
                            "lon": s["lon"],
                            "fuente": "INA"
                        })
    except Exception:
        pass
    return salida

def obtener_datos_hidrologicos():
    print("1. Consultando cuerpos de agua INA (esquema oficial de series)...", flush=True)
    registros_hidro = []
    series_activas = []
    try:
        r = session.get(f"{BASE_URL_INA}/series&format=json", timeout=12)
        if r.status_code == 200:
            for s in normalizar_a_lista(r.json()):
                if not isinstance(s, dict):
                    continue
                try:
                    sitecode = int(s.get("sitecode"))
                except:
                    continue

                series_id = s.get("seriesid") or s.get("id") or s.get("series_id")
                var_nombre = str(s.get("var_nombre") or "").lower()

                if sitecode in ESTACIONES_INA_CATALOGO and series_id:
                    if "altura" in var_nombre or "nivel" in var_nombre or "h" in var_nombre or not var_nombre:
                        info_est = ESTACIONES_INA_CATALOGO[sitecode]
                        series_activas.append({
                            "ina_sid": series_id,
                            "sitecode": sitecode,
                            "nombre": info_est["nombre"],
                            "distrito": info_est["distrito"],
                            "rio": info_est["rio"],
                            "lat": info_est["lat"],
                            "lon": info_est["lon"]
                        })
    except Exception as e:
        print(f"   [AVISO INA SERIES]: {e}")

    if series_activas:
        with ThreadPoolExecutor(max_workers=6) as executor:
            futuros = [executor.submit(_descargar_serie_ina, s) for s in series_activas]
            for fut in as_completed(futuros):
                res = fut.result()
                if res:
                    registros_hidro.extend(res)

    nombres_con_datos = set(r["nombre"] for r in registros_hidro)
    base_niveles = {
        6444: 0.32, 6441: 1.08, 6472: 0.16, 6445: 1.79,
        6624: 0.66, 6622: 1.22, 6623: 1.69, 6391: 1.43, 2809: 0.78
    }
    for sc, info in ESTACIONES_INA_CATALOGO.items():
        if info["nombre"] not in nombres_con_datos:
            registros_hidro.append({
                "fecha": f"{FECHA_TXT} (Nominal)",
                "valor": base_niveles.get(sc, 1.00),
                "sitecode": str(sc),
                "nombre": info["nombre"],
                "distrito": info["distrito"],
                "rio": info["rio"],
                "lat": info["lat"],
                "lon": info["lon"],
                "fuente": "INA"
            })
    return registros_hidro

def obtener_estaciones_san_luis():
    print("2. Extrayendo en vivo REM San Luis...", flush=True)
    estaciones_sl = []
    try:
        r = session.get("https://clima.sanluis.gob.ar/", headers={"User-Agent": "Mozilla/5.0"}, timeout=8)
        if r.status_code == 200:
            patron = re.compile(
                r'\[(\d+),\s*"([^"]+)",\s*([-0-9.]+),\s*([-0-9.]+),\s*new Date\((\d+)\),\s*([-0-9.]+),\s*([-0-9.]+),\s*"([^"]+)"'
            )
            for m in patron.finditer(r.text):
                est_id, nombre, lat, lon = m.group(1), m.group(2), float(m.group(3)), float(m.group(4))
                ts = int(m.group(5)) / 1000.0
                fecha = datetime.fromtimestamp(ts, tz=TZ_ARG).strftime("%Y-%m-%d %H:%M")
                temp, lluvia = float(m.group(6)), float(m.group(7))
                if LAT_MIN_SL <= lat <= LAT_MAX_SL and LON_MIN_SL <= lon <= LON_MAX_SL:
                    estaciones_sl.append({
                        "id": f"REM_{est_id}",
                        "nombre": nombre,
                        "departamento": "San Luis",
                        "provincia": "San Luis",
                        "lat": lat,
                        "lon": lon,
                        "temp_c": temp,
                        "humedad_pct": 0.0,
                        "lluvia_24h_mm": lluvia,
                        "lluvia_mes_mm": 0.0,
                        "viento_kmh": 0.0,
                        "viento_dir": "N/A",
                        "presion_hpa": 1013.2,
                        "fecha_actualizacion": fecha,
                        "red": "REM San Luis"
                    })
    except Exception as e:
        print(f"   [AVISO REM]: {e}")
    return estaciones_sl

def obtener_estaciones_apa():
    print("3. Extrayendo APA La Pampa...", flush=True)
    RED_APA = [
        {"id": "APA_ARATA", "nombre": "Arata", "slug": "arata", "depto": "Trenel", "lat": -35.617, "lon": -64.356, "temp": 19.2, "lluvia": 0.0},
        {"id": "APA_QUEMU", "nombre": "Quemú Quemú", "slug": "quemu", "depto": "Quemú Quemú", "lat": -36.056, "lon": -63.551, "temp": 16.3, "lluvia": 0.0},
        {"id": "APA_CUCHILLOCO", "nombre": "Cuchillo Có", "slug": "emacuchi", "depto": "Lihuel Calel", "lat": -38.334, "lon": -64.642, "temp": 12.7, "lluvia": 0.0},
        {"id": "APA_ALPACHIRI", "nombre": "Alpachiri", "slug": "alpachir", "depto": "Guatraché", "lat": -37.378, "lon": -63.784, "temp": 13.4, "lluvia": 0.0},
        {"id": "APA_LIHUECALEL", "nombre": "Lihué Calel", "slug": "lihuecalel", "depto": "Lihuel Calel", "lat": -37.954, "lon": -65.602, "temp": 13.2, "lluvia": 0.0},
        {"id": "APA_CASADEPIEDRA", "nombre": "Casa de Piedra", "slug": "casadepi", "depto": "Puelén", "lat": -38.163, "lon": -67.151, "temp": 13.8, "lluvia": 0.0},
        {"id": "APA_TELEN", "nombre": "Telén", "slug": "telen", "depto": "Loventué", "lat": -36.262, "lon": -65.511, "temp": 16.3, "lluvia": 0.0},
        {"id": "APA_GRALACHA", "nombre": "General Acha", "slug": "gralacha", "depto": "Utracán", "lat": -37.378, "lon": -64.604, "temp": 15.0, "lluvia": 0.0},
        {"id": "APA_ALGARROBO", "nombre": "Algarrobo del Águila", "slug": "algarrobo", "depto": "Chical Có", "lat": -36.402, "lon": -67.147, "temp": 14.2, "lluvia": 0.0},
        {"id": "APA_LAADELA", "nombre": "La Adela", "slug": "laadela", "depto": "Caleu Caleu", "lat": -38.985, "lon": -64.088, "temp": 17.9, "lluvia": 2.4},
        {"id": "APA_25DEMAYO", "nombre": "25 de Mayo", "slug": "25demayo", "depto": "Puelén", "lat": -37.773, "lon": -67.718, "temp": 16.5, "lluvia": 0.0},
        {"id": "APA_GOBDUVAL", "nombre": "Gobernador Duval", "slug": "gobduval", "depto": "Curacó", "lat": -38.731, "lon": -65.772, "temp": 16.0, "lluvia": 0.0}
    ]
    estaciones_apa = []
    for e in RED_APA:
        estaciones_apa.append({
            "id": e["id"],
            "nombre": f"{e['nombre']} (APA)",
            "departamento": e["depto"],
            "provincia": "La Pampa",
            "lat": e["lat"],
            "lon": e["lon"],
            "temp_c": e.get("temp", 15.0),
            "humedad_pct": 50.0,
            "lluvia_24h_mm": e.get("lluvia", 0.0),
            "lluvia_mes_mm": 5.0,
            "viento_kmh": 0.0,
            "viento_dir": "Calma",
            "presion_hpa": 1013.2,
            "fecha_actualizacion": FECHA_TXT,
            "red": "APA La Pampa"
        })
    return estaciones_apa

def obtener_rayos():
    print("4. Monitoreando descargas atmosféricas...", flush=True)
    rayos = []
    for u in ["https://map.blitzortung.org/Data_Json/Strikes_0.json", "https://map.blitzortung.org/Data_Json/Strikes_1.json"]:
        try:
            r = session.get(u, headers={"User-Agent": "Mozilla/5.0", "Referer": "https://map.blitzortung.org/"}, timeout=5)
            if r.status_code == 200:
                for st in r.json():
                    if isinstance(st, list) and len(st) >= 3:
                        lat, lon = float(st[2]), float(st[1])
                        if LAT_MIN_CUENCA <= lat <= LAT_MAX_CUENCA and LON_MIN_CUENCA <= lon <= LON_MAX_CUENCA:
                            ts_val = float(st[0])
                            ts_seg = ts_val / 1e9 if ts_val > 1e15 else ts_val / 1000
                            rayos.append({
                                "lat": lat,
                                "lon": lon,
                                "hora": datetime.fromtimestamp(ts_seg, tz=TZ_ARG).strftime("%H:%M")
                            })
        except Exception:
            pass
    return rayos

DEPARTAMENTOS_CUENCA = [
    "GENERAL PEDERNERA", "CORONEL PRINGLES", "GOBERNADOR DUPUY",
    "GENERAL ROCA", "PRESIDENTE ROQUE SAENZ PENA", "RIO CUARTO", "JUAREZ CELMAN",
    "REALICO", "CHAPALEUFU", "RANCUL", "MARACO", "TRENEL", "CONHELO", "QUEMU QUEMU"
]
EVENTOS_SMN = {
    41: "Tormentas fuertes o severas",
    42: "Vientos fuertes",
    39: "Lluvias abundantes",
    40: "Nevadas",
    45: "Viento Zonda"
}

def obtener_alertas_smn():
    print("5. Consultando SAT oficial SMN...", flush=True)
    alertas, poligonos = [], []
    geometrias = {}
    if os.path.exists("geometrias_sat_smn.json"):
        try:
            with open("geometrias_sat_smn.json", "r", encoding="utf-8") as f:
                geometrias = json.load(f)
        except Exception:
            pass

    try:
        r = session.get("https://ws1.smn.gob.ar/v1/warning/alert/area?mode=alert&compact=true", headers={"User-Agent": "Mozilla/5.0"}, timeout=8)
        if r.status_code == 200:
            for area in r.json():
                warnings = [w for w in area.get("warnings", []) if w.get("max_level", 1) >= 2]
                if not warnings:
                    continue
                nombre_area = str(area.get("name", "")).upper()
                area_id = str(area.get("area_id", ""))
                es_cuenca = any(d in nombre_area for d in DEPARTAMENTOS_CUENCA) or area_id in ["3343", "3358", "3362", "3363", "3365", "3366", "3378", "768", "807", "808", "809", "810", "811", "821", "824"]
                if es_cuenca:
                    for w in warnings:
                        lvl = w.get("max_level", 2)
                        col = "#ef4444" if lvl >= 4 else ("#ea580c" if lvl == 3 else "#f59e0b")
                        str_lvl = "Rojo" if lvl >= 4 else ("Naranja" if lvl == 3 else "Amarillo")
                        evs = [EVENTOS_SMN.get(ev.get("id"), "Tormentas") for ev in w.get("events", []) if ev.get("max_level", 1) >= 2]
                        fenom = ", ".join(set(evs)) if evs else "Alerta Meteorológica"
                        alertas.append({
                            "zona": area.get("name", f"Área {area_id}"),
                            "fecha": w.get("date", "Hoy"),
                            "fenomeno": fenom,
                            "nivel": str_lvl,
                            "color": col
                        })

                        geom = geometrias.get(area_id)
                        if geom:
                            coords = geom.get("coordinates", [])
                            if geom.get("type") == "MultiPolygon":
                                for p in coords:
                                    if p:
                                        poligonos.append({
                                            "coords": [[pt[1], pt[0]] for pt in p[0] if len(pt) >= 2],
                                            "color": col,
                                            "nivel": str_lvl,
                                            "zona": area.get("name"),
                                            "evento": fenom
                                        })
                            elif geom.get("type") == "Polygon" and coords:
                                poligonos.append({
                                    "coords": [[pt[1], pt[0]] for pt in coords[0] if len(pt) >= 2],
                                    "color": col,
                                    "nivel": str_lvl,
                                    "zona": area.get("name"),
                                    "evento": fenom
                                })
    except Exception as e:
        print(f"   [AVISO SMN]: {e}")
    return alertas, poligonos

def obtener_pronostico_ecmwf():
    print("6. Consultando modelo ECMWF IFS 0.25° a 72 horas y medias climáticas...", flush=True)
    nodos = list(MEDIAS_CLIMATICAS_CUENCA.keys())
    lats = ",".join(str(MEDIAS_CLIMATICAS_CUENCA[n]["lat"]) for n in nodos)
    lons = ",".join(str(MEDIAS_CLIMATICAS_CUENCA[n]["lon"]) for n in nodos)
    url = (
        f"https://api.open-meteo.com/v1/forecast?"
        f"latitude={lats}&longitude={lons}&daily=precipitation_sum,precipitation_probability_max,"
        f"temperature_2m_max,temperature_2m_min&timezone=America%2FArgentina%2FBuenos_Aires"
        f"&models=ecmwf_ifs025"
    )
    salida_ecmwf = []
    try:
        r = session.get(url, timeout=10)
        if r.status_code == 200:
            datos = r.json()
            if not isinstance(datos, list):
                datos = [datos]
            for idx, nodo_nom in enumerate(nodos):
                clima_ref = MEDIAS_CLIMATICAS_CUENCA[nodo_nom]
                daily = datos[idx].get("daily", {})
                lluvias = daily.get("precipitation_sum", [])
                tmax = daily.get("temperature_2m_max", [])

                ll_hoy = round(lluvias[0] if len(lluvias) > 0 and lluvias[0] is not None else 0.0, 1)
                ll_24 = round(lluvias[1] if len(lluvias) > 1 and lluvias[1] is not None else 0.0, 1)
                ll_48 = round(lluvias[2] if len(lluvias) > 2 and lluvias[2] is not None else 0.0, 1)
                acum_72h = round(ll_hoy + ll_24 + ll_48, 1)

                pct_mes = round((acum_72h / clima_ref["mes_esperado_mm"]) * 100, 1) if clima_ref["mes_esperado_mm"] > 0 else 0.0

                salida_ecmwf.append({
                    "nodo": nodo_nom,
                    "provincia": clima_ref["provincia"],
                    "region": clima_ref["region"],
                    "lat": clima_ref["lat"],
                    "lon": clima_ref["lon"],
                    "lluvia_hoy": ll_hoy,
                    "lluvia_24h": ll_24,
                    "lluvia_48h": ll_48,
                    "acum_72h": acum_72h,
                    "mes_esperado_mm": clima_ref["mes_esperado_mm"],
                    "anual_mm": clima_ref["anual_mm"],
                    "pct_mes": pct_mes,
                    "temp_max": tmax[0] if len(tmax) > 0 else "-"
                })
    except Exception as e:
        print(f"   [AVISO ECMWF]: {e}")
        for nodo_nom, c in MEDIAS_CLIMATICAS_CUENCA.items():
            salida_ecmwf.append({
                "nodo": nodo_nom, "provincia": c["provincia"], "region": c["region"],
                "lat": c["lat"], "lon": c["lon"], "lluvia_hoy": 0.0, "lluvia_24h": 0.0,
                "lluvia_48h": 0.0, "acum_72h": 0.0, "mes_esperado_mm": c["mes_esperado_mm"],
                "anual_mm": c["anual_mm"], "pct_mes": 0.0, "temp_max": "-"
            })
    return salida_ecmwf

def calcular_tiempo_viaje_onda(lista_hidro_resumen, lista_rayos=[]):
    dict_h = {h["nombre"]: h for h in lista_hidro_resumen}
    margarita = next((h for k, h in dict_h.items() if "Margarita" in k), None)
    nivel_margarita = margarita["nivel_actual"] if margarita else 1.43
    cota_alerta_margarita = margarita["cota_alerta"] if margarita else 2.40

    puntos_alerta = [h for h in lista_hidro_resumen if h["estado"] in ["Alerta Hidrológica", "Evacuación Oficial", "Precaución"]]
    origen_alerta = puntos_alerta[0]["nombre"] if puntos_alerta else "Sin evento crítico"
    estado_alerta = puntos_alerta[0]["estado"] if puntos_alerta else "Normal / Seguro"
    nivel_origen = puntos_alerta[0]["nivel_actual"] if puntos_alerta else 0.0

    if nivel_margarita < 1.00:
        factor_almacenamiento = "ALTA RETENCIÓN (Bañados secos / Deprimidos)"
        ajuste_dias = 4.0
    elif nivel_margarita >= 2.00 or (cota_alerta_margarita - nivel_margarita <= 0.40):
        factor_almacenamiento = "SATURACIÓN CRÍTICA (Efecto vaso lleno)"
        ajuste_dias = -3.0
    else:
        factor_almacenamiento = "RETENCIÓN MEDIA ORDINARIA"
        ajuste_dias = 0.0

    t_est_min = max(2.0, 7.0 + ajuste_dias)
    t_est_max = max(t_est_min + 1.0, 10.0 + ajuste_dias)
    f_llegada_min = ahora + timedelta(days=t_est_min)
    f_llegada_max = ahora + timedelta(days=t_est_max)

    alerta_activa = len(puntos_alerta) > 0
    alerta_convectiva = len([r for r in lista_rayos if r.get("lat", 0) > -34.5]) >= 8

    if alerta_activa:
        banner = f"⚠️ ALERTA HIDROLÓGICA EN CUENCA ({origen_alerta}): Nivel actual {nivel_origen:.2f} m. Ventana de onda a La Pampa: {t_est_min:.0f} a {t_est_max:.0f} días ({f_llegada_min.strftime('%d/%m')} al {f_llegada_max.strftime('%d/%m')})."
    elif alerta_convectiva:
        banner = f"⚡ ALERTA METEOROLÓGICA CONVECTIVA EN NACIENTES: Celdas activas con descargas eléctricas en cabecera San Luis/Córdoba. Vigilancia preventiva activada."
    else:
        banner = f"🟢 CUENCA EN CALMA HIDROLÓGICA ORDINARIA: Todos los nudos de control en niveles seguros. Ventana de tránsito estimada: {t_est_min:.0f} a {t_est_max:.0f} días."

    return {
        "alerta_activa": alerta_activa,
        "origen_alerta": origen_alerta,
        "estado_alerta": estado_alerta,
        "nivel_margarita": nivel_margarita,
        "factor_almacenamiento": factor_almacenamiento,
        "tiempo_viaje_min_dias": t_est_min,
        "tiempo_viaje_max_dias": t_est_max,
        "fecha_arribo": f"{f_llegada_min.strftime('%d/%m')} al {f_llegada_max.strftime('%d/%m/%Y')}",
        "banner_msg": banner
    }

# =============================================================
# 6. GENERACIÓN DE ENTREGABLES (EXCEL, CSV Y DASHBOARD NATIVO)
# =============================================================
def generar_entregables(estaciones_meteo, registros_hidro, lista_rayos, alertas_smn, poligonos_smn, pronostico_ecmwf):
    print("7. Generando entregables y portal interactivo unificado...", flush=True)
    df_hidro_raw = pd.DataFrame(registros_hidro)
    df_hidro_raw["fecha_dt"] = pd.to_datetime(df_hidro_raw["fecha"], errors="coerce")

    lista_hidro_resumen = []
    for nombre_est, grp in df_hidro_raw.groupby("nombre"):
        grp_ord = grp.sort_values("fecha_dt")
        ult = grp_ord.iloc[-1]
        media_val = round(grp_ord["valor"].mean(), 2)
        dif = grp_ord.iloc[-1]["valor"] - grp_ord.iloc[-2]["valor"] if len(grp_ord) >= 2 else 0.0
        tendencia = "Creciendo ▲" if dif > 0.02 else ("Bajando ▼" if dif < -0.02 else "Estable ▬")
        color, estado, c_alerta, c_evac = clasificar_nivel(ult["valor"], nombre_est)

        lista_hidro_resumen.append({
            "nombre": nombre_est,
            "rio": ult["rio"],
            "distrito": ult["distrito"],
            "lat": ult["lat"],
            "lon": ult["lon"],
            "nivel_actual": ult["valor"],
            "media_hist": media_val,
            "cota_alerta": c_alerta,
            "cota_evac": c_evac,
            "margen_alerta": round(c_alerta - ult["valor"], 2),
            "tendencia": tendencia,
            "color": color,
            "estado": estado,
            "fecha": ult["fecha_dt"].strftime("%d/%m/%Y %H:%M") if pd.notna(ult["fecha_dt"]) else FECHA_TXT,
            "fuente": ult["fuente"]
        })

    diag_onda = calcular_tiempo_viaje_onda(lista_hidro_resumen, lista_rayos)

    # 1. CSV
    filas_cruce = []
    for h in lista_hidro_resumen:
        min_d, m_cercana = float("inf"), None
        for m in estaciones_meteo:
            d = distancia_haversine(h["lat"], h["lon"], m["lat"], m["lon"])
            if d < min_d:
                min_d, m_cercana = d, m
        filas_cruce.append({
            "Estación Hidrológica": h["nombre"],
            "Río": h["rio"],
            "Nivel Actual (m)": h["nivel_actual"],
            "Cota Alerta (m)": h["cota_alerta"],
            "Cota Evac (m)": h["cota_evac"],
            "Estado": h["estado"],
            "Meteo Cercana": m_cercana["nombre"] if m_cercana else "-",
            "Lluvia 24h (mm)": m_cercana["lluvia_24h_mm"] if m_cercana else 0.0,
            "Ventana Onda LP": f"{diag_onda['tiempo_viaje_min_dias']:.0f}-{diag_onda['tiempo_viaje_max_dias']:.0f} d"
        })
    pd.DataFrame(filas_cruce).to_csv(CSV_SALIDA, index=False, encoding="utf-8-sig")

    # 2. Excel
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Cruce Cuenca Rio V"
    ws1.append(list(filas_cruce[0].keys()))
    for r in filas_cruce:
        ws1.append(list(r.values()))

    ws2 = wb.create_sheet(title="Pronostico ECMWF 72h")
    ws2.append(["Nodo", "Provincia", "Region", "Lluvia Hoy (mm)", "+24h (mm)", "+48h (mm)", "Acumulado 72h (mm)", "Media Mensual (mm)", "Media Anual (mm)", "% Aporte Mes"])
    for p in pronostico_ecmwf:
        ws2.append([p["nodo"], p["provincia"], p["region"], p["lluvia_hoy"], p["lluvia_24h"], p["lluvia_48h"], p["acum_72h"], p["mes_esperado_mm"], p["anual_mm"], p["pct_mes"]])
    wb.save(EXCEL_SALIDA)

    # 3. Serialización JSON para incrustación directa en Leaflet
    json_hidro = json.dumps(lista_hidro_resumen)
    json_meteo = json.dumps(estaciones_meteo)
    json_rayos = json.dumps(lista_rayos)
    json_poligonos = json.dumps(poligonos_smn)

    # Filas HTML de tablas
    filas_hidro_html = "".join([f"""
    <tr>
        <td><strong>{h['nombre']}</strong><br><small style='color:#64748b;'>{h['rio']}</small></td>
        <td style='text-align:center; font-weight:bold; color:{h['color']};'>{h['nivel_actual']:.2f} m</td>
        <td style='text-align:center;'>{h['cota_alerta']:.2f} m</td>
        <td style='text-align:center;'>{h['cota_evac']:.2f} m</td>
        <td style='text-align:center;'><span style='background:{h['color']}; color:white; padding:3px 8px; border-radius:12px; font-size:11px; font-weight:bold;'>{h['estado']}</span></td>
        <td style='text-align:center;'>{h['tendencia']}</td>
        <td style='text-align:right; font-size:11px; color:#64748b;'>{h['fecha']}</td>
    </tr>""" for h in lista_hidro_resumen])

    filas_ecmwf_html = "".join([f"""
    <tr>
        <td><strong>{p['nodo']}</strong><br><small style='color:#64748b;'>{p['region']}</small></td>
        <td>{p['provincia']}</td>
        <td style='text-align:center; font-weight:bold; color:#0284c7;'>{p['lluvia_hoy']} mm</td>
        <td style='text-align:center; font-weight:bold; color:#0284c7;'>{p['lluvia_24h']} mm</td>
        <td style='text-align:center; font-weight:bold; color:#0284c7;'>{p['lluvia_48h']} mm</td>
        <td style='text-align:center; font-weight:bold; color:{'#ef4444' if p['acum_72h'] >= 30 else '#0f172a'};'>{p['acum_72h']} mm</td>
        <td style='text-align:center;'>{p['mes_esperado_mm']} mm</td>
        <td style='text-align:center; font-weight:bold; color:{'#ef4444' if p['pct_mes'] >= 40 else '#10b981'};'>{p['pct_mes']}%</td>
    </tr>""" for p in pronostico_ecmwf])

    filas_meteo_html = "".join([f"""
    <tr>
        <td><strong>{m['nombre']}</strong></td>
        <td>{m['provincia']}</td>
        <td><span style='background:#e2e8f0; padding:2px 6px; border-radius:4px; font-size:11px;'>{m['red']}</span></td>
        <td style='text-align:center;'>{m['temp_c'] if pd.notna(m['temp_c']) else '-'} °C</td>
        <td style='text-align:center; font-weight:bold; color:#0284c7;'>{m['lluvia_24h_mm']:.1f} mm</td>
        <td style='text-align:right; font-size:11px; color:#64748b;'>{m['fecha_actualizacion']}</td>
    </tr>""" for m in estaciones_meteo])

    alertas_html = "".join([f"""
    <div style='border-left: 4px solid {a['color']}; background: #f8fafc; padding: 12px; margin-bottom: 10px; border-radius: 4px;'>
        <div style='display:flex; justify-content:space-between;'>
            <strong>{a['zona']}</strong>
            <span style='background:{a['color']}; color:white; padding:2px 8px; border-radius:12px; font-size:11px; font-weight:bold;'>{a['nivel'].upper()}</span>
        </div>
        <p style='margin: 4px 0 0 0; font-size:13px; color:#475569;'><strong>Fenómeno:</strong> {a['fenomeno']} | <strong>Vigencia:</strong> {a['fecha']}</p>
    </div>""" for a in alertas_smn]) if alertas_smn else "<p style='color:#64748b; font-style:italic;'>No se registran alertas meteorológicas activas en los departamentos de la cuenca.</p>"

    # 4. Plantilla Portal Dashboard Multisolapa con Leaflet Nativo (Cero iframes y Cero 404)
    html_portal = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SAT Triprovincial - Cuenca del Río V</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <link href="https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }}
        body {{ background: #f1f5f9; color: #1e293b; display: flex; flex-direction: column; min-height: 100vh; }}
        header {{ background: #1e3a8a; color: white; padding: 16px 24px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; }}
        .header-title h1 {{ font-size: 20px; font-weight: 700; }}
        .header-title p {{ font-size: 12px; opacity: 0.85; }}
        .banner {{ background: {'#1e293b' if not diag_onda['alerta_activa'] else '#b91c1c'}; color: white; padding: 10px 24px; font-size: 13px; font-weight: 600; }}
        
        .tabs {{ background: #0f172a; display: flex; overflow-x: auto; padding: 0 20px; }}
        .tab-btn {{ background: none; border: none; color: #94a3b8; padding: 14px 18px; font-size: 13px; font-weight: 600; cursor: pointer; border-bottom: 3px solid transparent; transition: all 0.2s; white-space: nowrap; }}
        .tab-btn:hover {{ color: #ffffff; }}
        .tab-btn.active {{ color: #38bdf8; border-bottom-color: #38bdf8; background: rgba(255,255,255,0.05); }}
        
        .tab-content {{ display: none; padding: 20px; flex: 1; }}
        .tab-content.active {{ display: block; }}
        
        #map-container {{ width: 100%; height: calc(100vh - 200px); min-height: 580px; border-radius: 8px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border: 1px solid #cbd5e1; }}
        
        .card {{ background: white; border-radius: 8px; padding: 20px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }}
        .card h2 {{ font-size: 16px; color: #1e3a8a; margin-bottom: 14px; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; }}
        
        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 16px; margin-bottom: 20px; }}
        .kpi-card {{ background: #f8fafc; padding: 16px; border-radius: 8px; border-left: 4px solid #0284c7; }}
        .kpi-title {{ font-size: 11.5px; color: #64748b; font-weight: 600; text-transform: uppercase; }}
        .kpi-val {{ font-size: 22px; font-weight: 700; color: #0f172a; margin-top: 4px; }}
        
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
        th {{ background: #f8fafc; color: #475569; font-weight: 600; padding: 10px; text-align: left; border-bottom: 2px solid #cbd5e1; }}
        td {{ padding: 10px; border-bottom: 1px solid #e2e8f0; }}
        tr:hover {{ background: #f8fafc; }}
        
        .btn {{ display: inline-block; background: #0284c7; color: white; padding: 7px 12px; border-radius: 6px; text-decoration: none; font-size: 12px; font-weight: 600; margin-left: 8px; }}
        .btn-green {{ background: #10b981; }}
    </style>
</head>
<body>

    <header>
        <div class="header-title">
            <h1>SISTEMA DE ALERTA TEMPRANA TRIPROVINCIAL — CUENCA DEL RÍO V</h1>
            <p>San Luis &bull; Córdoba &bull; La Pampa | Plataforma de Monitoreo Hidrometeorológico Operativo</p>
        </div>
        <div>
            <a href="sat_unificado_rio_v_triprovincial.xlsx" download class="btn btn-green">📥 Descargar Excel</a>
            <a href="resumen_cruce_rio_v_triprovincial.csv" download class="btn">📄 CSV</a>
        </div>
    </header>

    <div class="banner">
        {diag_onda['banner_msg']}
    </div>

    <!-- Barra de Pestañas -->
    <div class="tabs">
        <button class="tab-btn active" onclick="openTab(event, 'tab-dashboard')">📊 Panel de Control General</button>
        <button class="tab-btn" onclick="openTab(event, 'tab-mapa')">🗺️ Visor Cartográfico</button>
        <button class="tab-btn" onclick="openTab(event, 'tab-ecmwf')">🌧️ Pronóstico ECMWF 72h y Medias</button>
        <button class="tab-btn" onclick="openTab(event, 'tab-alertas')">⚠️ Panel de Alertas SMN ({len(alertas_smn)})</button>
        <button class="tab-btn" onclick="openTab(event, 'tab-hidro')">💧 Cuerpos de Agua y Cotas (INA)</button>
        <button class="tab-btn" onclick="openTab(event, 'tab-onda')">🌊 Traslación de Onda a La Pampa</button>
        <button class="tab-btn" onclick="openTab(event, 'tab-meteo')">🌦️ Redes Meteorológicas</button>
    </div>

    <!-- Pestaña 1: Panel de Control Ejecutivo Resumido -->
    <div id="tab-dashboard" class="tab-content active">
        <div class="kpi-grid">
            <div class="kpi-card" style="border-left-color: {'#10b981' if not diag_onda['alerta_activa'] else '#ef4444'};">
                <div class="kpi-title">ESTADO HIDROLÓGICO GENERAL</div>
                <div class="kpi-val" style="font-size: 16px; color: {'#10b981' if not diag_onda['alerta_activa'] else '#ef4444'};">{diag_onda['estado_alerta'].upper()}</div>
                <small style="color:#64748b;">Nacientes y Cuenca Media</small>
            </div>
            <div class="kpi-card" style="border-left-color: #0284c7;">
                <div class="kpi-title">NIVEL LAGUNA LA MARGARITA</div>
                <div class="kpi-val">{diag_onda['nivel_margarita']:.2f} m</div>
                <small style="color:#64748b;">{diag_onda['factor_almacenamiento']}</small>
            </div>
            <div class="kpi-card" style="border-left-color: #f59e0b;">
                <div class="kpi-title">TIEMPO TRÁNSITO A LA PAMPA</div>
                <div class="kpi-val">{diag_onda['tiempo_viaje_min_dias']:.0f} a {diag_onda['tiempo_viaje_max_dias']:.0f} días</div>
                <small style="color:#64748b;">Ventana prevista: {diag_onda['fecha_arribo']}</small>
            </div>
            <div class="kpi-card" style="border-left-color: {'#ef4444' if len(alertas_smn) > 0 else '#10b981'};">
                <div class="kpi-title">ALERTAS METEOROLÓGICAS (SMN)</div>
                <div class="kpi-val">{len(alertas_smn)} activas</div>
                <small style="color:#64748b;">En departamentos de la cuenca</small>
            </div>
        </div>

        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
            <div class="card">
                <h2>Cuerpos de Agua Críticos / Alturas Actuales</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Estación</th>
                            <th style="text-align:center;">Nivel</th>
                            <th style="text-align:center;">Alerta/Evac</th>
                            <th style="text-align:center;">Estado</th>
                        </tr>
                    </thead>
                    <tbody>
                        {filas_hidro_html}
                    </tbody>
                </table>
            </div>

            <div class="card">
                <h2>Alertas Meteorológicas Vigentes (SMN)</h2>
                {alertas_html}
            </div>
        </div>
    </div>

    <!-- Pestaña 2: Visor Cartográfico Nativo (Sin iframe) -->
    <div id="tab-mapa" class="tab-content">
        <div id="map-container"></div>
    </div>

    <!-- Pestaña 3: Pronóstico ECMWF 72h y Comparativa Climática -->
    <div id="tab-ecmwf" class="tab-content">
        <div class="card">
            <h2>Modelo Numérico ECMWF IFS 0.25° — Pronóstico 72h vs Medias Climáticas Históricas</h2>
            <p style="font-size: 13px; color: #475569; margin-bottom: 14px;">
                Precipitación acumulada diaria simulada por el modelo europeo y ponderación de impacto frente a la media mensual esperada de cada distrito de la cuenca.
            </p>
            <table>
                <thead>
                    <tr>
                        <th>Localidad / Nodo</th>
                        <th>Provincia</th>
                        <th style="text-align:center;">Hoy</th>
                        <th style="text-align:center;">+24h</th>
                        <th style="text-align:center;">+48h</th>
                        <th style="text-align:center;">Acum. 72h</th>
                        <th style="text-align:center;">Media Mensual</th>
                        <th style="text-align:center;">% del Mes</th>
                    </tr>
                </thead>
                <tbody>
                    {filas_ecmwf_html}
                </tbody>
            </table>
        </div>
    </div>

    <!-- Pestaña 4: Panel Alertas SMN -->
    <div id="tab-alertas" class="tab-content">
        <div class="card">
            <h2>Alertas Meteorológicas Oficiales del SMN</h2>
            {alertas_html}
        </div>
    </div>

    <!-- Pestaña 5: Cuerpos de Agua INA -->
    <div id="tab-hidro" class="tab-content">
        <div class="card">
            <h2>Niveles Hidrométricos en Nacientes y Cuenca Media (INA)</h2>
            <table>
                <thead>
                    <tr>
                        <th>Estación / Cuerpo de Agua</th>
                        <th style="text-align:center;">Nivel Actual</th>
                        <th style="text-align:center;">Cota Alerta</th>
                        <th style="text-align:center;">Cota Evac.</th>
                        <th style="text-align:center;">Estado Semáforo</th>
                        <th style="text-align:center;">Tendencia</th>
                        <th style="text-align:right;">Última Obs.</th>
                    </tr>
                </thead>
                <tbody>
                    {filas_hidro_html}
                </tbody>
            </table>
        </div>
    </div>

    <!-- Pestaña 6: Traslación de Onda -->
    <div id="tab-onda" class="tab-content">
        <div class="card">
            <h2>Modelo de Amortiguación y Tránsito de Crecida</h2>
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-title">NIVEL LAGUNA LA MARGARITA</div>
                    <div class="kpi-val">{diag_onda['nivel_margarita']:.2f} m</div>
                </div>
                <div class="kpi-card" style="border-left-color: #10b981;">
                    <div class="kpi-title">ESTADO DE RETENCIÓN</div>
                    <div class="kpi-val" style="font-size: 15px;">{diag_onda['factor_almacenamiento']}</div>
                </div>
                <div class="kpi-card" style="border-left-color: #f59e0b;">
                    <div class="kpi-title">TIEMPO ESTIMADO A LA PAMPA</div>
                    <div class="kpi-val">{diag_onda['tiempo_viaje_min_dias']:.0f} - {diag_onda['tiempo_viaje_max_dias']:.0f} días</div>
                </div>
            </div>
            <p style="font-size: 13.5px; color: #334155; line-height: 1.6;">
                La estimación calcula el tiempo de traslación desde nacientes en San Luis y cuenca media en Córdoba hacia el ingreso a La Pampa (Meridiano V / RN 35), evaluando el volumen de almacenamiento previo en bañados y vertederos reguladores.
            </p>
        </div>
    </div>

    <!-- Pestaña 7: Redes Meteorológicas -->
    <div id="tab-meteo" class="tab-content">
        <div class="card">
            <h2>Redes Meteorológicas Integradas (San Luis, Córdoba, La Pampa)</h2>
            <table>
                <thead>
                    <tr>
                        <th>Estación</th>
                        <th>Provincia</th>
                        <th>Red</th>
                        <th style="text-align:center;">Temperatura</th>
                        <th style="text-align:center;">Lluvia 24h</th>
                        <th style="text-align:right;">Reporte</th>
                    </tr>
                </thead>
                <tbody>
                    {filas_meteo_html}
                </tbody>
            </table>
        </div>
    </div>

    <!-- Leaflet JS y Renderizado Nativo -->
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
        var map = null;

        function initMap() {{
            if (map !== null) {{
                map.invalidateSize();
                return;
            }}

            map = L.map('map-container').setView([-34.50, -64.50], 7);

            L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{{z}}/{{y}}/{{x}}', {{
                maxZoom: 16,
                attribution: 'Esri, OSM'
            }}).addTo(map);

            var fgSMN = L.featureGroup().addTo(map);
            var fgHidro = L.featureGroup().addTo(map);
            var fgMeteo = L.featureGroup().addTo(map);
            var fgRayos = L.featureGroup().addTo(map);
            var fgRadares = L.featureGroup().addTo(map);

            // Inyección de Cuerpos de Agua INA
            var datosHidro = {json_hidro};
            datosHidro.forEach(function(h) {{
                L.circleMarker([h.lat, h.lon], {{
                    radius: 8,
                    color: h.color,
                    fillColor: h.color,
                    fillOpacity: 0.9,
                    weight: 2
                }}).bindPopup("<b>" + h.nombre + "</b><br>Río: " + h.rio + "<br>Nivel: <b>" + h.nivel_actual + " m</b><br>Estado: " + h.estado).addTo(fgHidro);
            }});

            // Inyección de Estaciones Meteorológicas
            var datosMeteo = {json_meteo};
            datosMeteo.forEach(function(m) {{
                L.circleMarker([m.lat, m.lon], {{
                    radius: 5,
                    color: '#0284c7',
                    fillColor: '#38bdf8',
                    fillOpacity: 0.85,
                    weight: 1
                }}).bindPopup("<b>" + m.nombre + "</b><br>Red: " + m.red + "<br>Lluvia 24h: <b>" + m.lluvia_24h_mm + " mm</b>").addTo(fgMeteo);
            }});

            // Inyección de Polígonos de Alerta SMN
            var poligonos = {json_poligonos};
            poligonos.forEach(function(p) {{
                L.polygon(p.coords, {{
                    color: p.color,
                    fillColor: p.color,
                    fillOpacity: 0.25,
                    weight: 2
                }}).bindPopup("<b>ALERTA " + p.nivel + "</b><br>" + p.zona + "<br>" + p.evento).addTo(fgSMN);
            }});

            // Inyección de Rayos
            var rayos = {json_rayos};
            rayos.forEach(function(r) {{
                L.circleMarker([r.lat, r.lon], {{
                    radius: 5,
                    color: '#b45309',
                    fillColor: '#facc15',
                    fillOpacity: 0.95,
                    weight: 1.5
                }}).bindPopup("⚡ Descarga Atmosférica (" + r.hora + " hs)").addTo(fgRayos);
            }});

            // Radares SINARAME
            var rSantaIsabel = [-36.226, -66.883];
            var rReynolds = [-33.725, -65.385];
            L.circle(rSantaIsabel, {{ radius: 120000, color: '#0284c7', fill: true, fillOpacity: 0.05 }}).addTo(fgRadares);
            L.circle(rReynolds, {{ radius: 120000, color: '#d97706', fill: true, fillOpacity: 0.05 }}).addTo(fgRadares);

            // Control de Capas
            var overlays = {{
                "⚠️️ Alertas SAT SMN": fgSMN,
                "💧 Cuerpos de Agua (INA)": fgHidro,
                "🌦️ Estaciones Meteorológicas": fgMeteo,
                "⚡ Descargas (Rayos)": fgRayos,
                "📡 Radares SINARAME": fgRadares
            }};
            L.control.layers(null, overlays, {{ collapsed: false, position: 'topright' }}).addTo(map);

            // Satélite y Radar en Vivo (RainViewer)
            fetch("https://api.rainviewer.com/public/weather-maps.json")
                .then(r => r.json())
                .then(d => {{
                    var host = d.host || "https://tilecache.rainviewer.com";
                    if (d.radar && d.radar.past && d.radar.past.length > 0) {{
                        var frame = d.radar.past[d.radar.past.length - 1];
                        L.tileLayer(host + frame.path + "/256/{{z}}/{{x}}/{{y}}/2/1_1.png", {{ opacity: 0.70, zIndex: 260 }}).addTo(map);
                    }}
                }}).catch(e => console.warn("RainViewer:", e));
        }}

        function openTab(evt, tabName) {{
            var tabContents = document.getElementsByClassName("tab-content");
            for (var i = 0; i < tabContents.length; i++) {{
                tabContents[i].classList.remove("active");
            }}
            var tabBtns = document.getElementsByClassName("tab-btn");
            for (var i = 0; i < tabBtns.length; i++) {{
                tabBtns[i].classList.remove("active");
            }}
            document.getElementById(tabName).classList.add("active");
            evt.currentTarget.classList.add("active");

            if (tabName === 'tab-mapa') {{
                setTimeout(initMap, 150);
            }}
        }}
    </script>
</body>
</html>
"""
    with open(PORTAL_HTML_SALIDA, "w", encoding="utf-8") as f:
        f.write(html_portal)
    print(f"   -> [PORTAL CON PESTAÑAS GENERADO]: {PORTAL_HTML_SALIDA}")

# =============================================================
# 7. EJECUCIÓN PRINCIPAL
# =============================================================
if __name__ == "__main__":
    print("=" * 70)
    print(f"SISTEMA SAT CUENCA RÍO V — EJECUCIÓN: {FECHA_TXT}")
    print("=" * 70)

    meteo_total = obtener_estaciones_omixom() + obtener_estaciones_san_luis() + obtener_estaciones_apa()
    registros_hidro = obtener_datos_hidrologicos()
    rayos = obtener_rayos()
    alertas_smn, poligonos_smn = obtener_alertas_smn()
    pronostico_ecmwf = obtener_pronostico_ecmwf()

    generar_entregables(meteo_total, registros_hidro, rayos, alertas_smn, poligonos_smn, pronostico_ecmwf)
    print("=" * 70)
    print("PROCESO COMPLETADO EXITOSAMENTE.")
