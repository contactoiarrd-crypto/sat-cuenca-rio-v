# -*- coding: utf-8 -*-
"""
SISTEMA DE ALERTA TEMPRANA TRIPROVINCIAL: CUENCA RÍO V (SAN LUIS - CÓRDOBA - LA PAMPA)
Monitoreo Hidrometeorológico, Cuerpos de Agua INA, SAT SMN / WMO CAP, Pronóstico ECMWF IFS,
Mosaico de Radares Argentinos en Tiempo Real, Redes Meteorológicas Discriminadas y Visor Vial.
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
# 3. CATÁLOGO ESTÁTICO REDES OMIXOM (CÓRDOBA Y LA PAMPA)
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
# 5. EXTRACCIÓN DE DATOS DE REDES
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
                except Exception:
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
        {"id": "APA_LAADELA", "nombre": "La Adela", "slug": "laadela", "depto": "Caleu Caleu", "lat": -38.985, "lon": -64.088, "temp": 17.9, "lluvia: 2.4"},
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

# =============================================================
# 6. ALERTAS SAT SMN CON BÚSQUEDA GENERAL
# =============================================================
TERMINOS_CUENCA = [
    "PEDERNERA", "PRINGLES", "DUPUY", "CAPITAL", "CHACABUCO", "SAN LUIS",
    "ROCA", "SAENZ PENA", "RIO CUARTO", "JUAREZ CELMAN", "CORDOBA",
    "REALICO", "CHAPALEUFU", "RANCUL", "MARACO", "TRENEL", "CONHELO", "QUEMU QUEMU", "LA PAMPA"
]

EVENTOS_SMN = {
    41: "Tormentas fuertes o severas",
    42: "Vientos fuertes",
    39: "Lluvias abundantes",
    40: "Nevadas",
    45: "Viento Zonda"
}

def normalizar_txt(s):
    if not s:
        return ""
    import unicodedata
    return ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn').upper()

def obtener_alertas_smn():
    print("5. Consultando feed oficial SAT SMN y CAP WMO...", flush=True)
    alertas, poligonos = [], []
    geometrias = {}
    if os.path.exists("geometrias_sat_smn.json"):
        try:
            with open("geometrias_sat_smn.json", "r", encoding="utf-8") as f:
                geometrias = json.load(f)
        except Exception:
            pass

    try:
        r = session.get("https://ws1.smn.gob.ar/v1/warning/alert/area?mode=alert", headers={"User-Agent": "Mozilla/5.0"}, timeout=8)
        if r.status_code == 200:
            for area in r.json():
                nom = normalizar_txt(area.get("name", ""))
                area_id = str(area.get("area_id", ""))
                es_cuenca = any(t in nom for t in TERMINOS_CUENCA)

                for w in area.get("warnings", []):
                    lvl = w.get("max_level", 1)
                    if lvl >= 2 and es_cuenca:
                        col = "#ef4444" if lvl >= 4 else ("#ea580c" if lvl == 3 else "#f59e0b")
                        str_lvl = "Rojo" if lvl >= 4 else ("Naranja" if lvl == 3 else "Amarillo")
                        evs = [EVENTOS_SMN.get(ev.get("id"), "Alerta Meteorológica") for ev in w.get("events", []) if ev.get("max_level", 1) >= 2]
                        fenom = ", ".join(set(evs)) if evs else "Tormentas / Lluvias"

                        alertas.append({
                            "zona": area.get("name", f"Zona {area_id}"),
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
        print(f"   [AVISO API SMN]: {e}")

    if not alertas:
        try:
            r_cap = session.get("https://ssl.smn.gob.ar/CAP/AR.php", headers={"User-Agent": "Mozilla/5.0"}, timeout=8)
            if r_cap.status_code == 200 and "<rss" in r_cap.text:
                soup = BeautifulSoup(r_cap.content, "xml")
                for item in soup.find_all("item"):
                    title = item.find("title").text if item.find("title") else ""
                    desc = item.find("description").text if item.find("description")
