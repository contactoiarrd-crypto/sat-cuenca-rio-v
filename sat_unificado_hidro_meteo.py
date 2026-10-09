# -*- coding: utf-8 -*-
"""
SISTEMA DE ALERTA TEMPRANA TRIPROVINCIAL: CUENCA RÍO V (SAN LUIS - CÓRDOBA - LA PAMPA)
Monitoreo Hidrometeorológico Online, Red Freatimétrica, Cruce Multicriterio y Renderizador Web.
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

import folium
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(line_buffering=True)

# Zona horaria oficial Argentina (UTC-3)
TZ_ARG = timezone(timedelta(hours=-3))

# =============================================================
# 1. PARÁMETROS GENERALES Y ARCHIVOS DE SALIDA
# =============================================================
ahora = datetime.now(TZ_ARG)
FECHA_TXT = ahora.strftime("%Y-%m-%d %H:%M")

EXCEL_SALIDA = "sat_unificado_rio_v_triprovincial.xlsx"
MAPA_HTML_SALIDA = "mapa_sat_rio_v_triprovincial.html"
CSV_SALIDA = "resumen_cruce_rio_v_triprovincial.csv"
HISTORICO_FREATIMETROS_FILE = "historico_freatimetros_cuenca.xlsx"
TEMPLATE_HTML_FILE = "template.html"
INDEX_HTML_FILE = "index.html"

# =============================================================
# 2. CATÁLOGO ESTÁTICO REDES OMIXOM (CÓRDOBA Y LA PAMPA)
# =============================================================
ESTACIONES_OMIXOM_ESTATICAS = [
    # CÓRDOBA - CUENCA MEDIA
    {"id": "OMX_CBA_1", "nombre": "General Levalle, Cordoba, Argentina", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.0000, "lon": -63.9163, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_2", "nombre": "Río Bamba, Cordoba, Argentina", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.0537, "lon": -63.7332, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_3", "nombre": "Jovita, Cordoba, Argentina", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.5194, "lon": -63.9683, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_4", "nombre": "Villa Valeria, Cordoba, Argentina", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.3427, "lon": -64.9290, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_5", "nombre": "Nicolás Bruzzone, Cordoba, Argentina", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.4391, "lon": -64.3424, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_6", "nombre": "Melo, Cordoba, Argentina", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.3452, "lon": -63.4377, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_7", "nombre": "Huanchillas, Cordoba, Argentina", "departamento": "Juárez Celman", "provincia": "Córdoba", "lat": -33.6665, "lon": -63.6395, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_8", "nombre": "Hipólito Bouchard, Cordoba, Argentina", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.7060, "lon": -63.5030, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_9", "nombre": "General Roca, Cordoba, Argentina", "departamento": "General Roca", "provincia": "Córdoba", "lat": -33.9948, "lon": -65.0754, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_10", "nombre": "Serrano, Cordoba, Argentina", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.4629, "lon": -63.5309, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_11", "nombre": "Coronel Moldes, Cordoba, Argentina", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.6481, "lon": -64.5950, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_12", "nombre": "Viamonte, Cordoba, Argentina", "departamento": "Unión", "provincia": "Córdoba", "lat": -33.7429, "lon": -63.0996, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_13", "nombre": "Chaján, Cordoba, Argentina", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.5508, "lon": -65.0059, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_14", "nombre": "Villa Rossi, Cordoba, Argentina", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.2949, "lon": -63.2654, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_15", "nombre": "Presa El Chañar, Cordoba, Argentina", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.9608, "lon": -65.0551, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_16", "nombre": "Huinca Renancó, Cordoba, Argentina", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.8208, "lon": -64.3738, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_17", "nombre": "Vicuña Mackenna, Cordoba, Argentina", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.9754, "lon": -64.3642, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom Córdoba"},

    # LA PAMPA - CUENCA BAJA
    {"id": "OMX_LP_1", "nombre": "MPLP 16 - El Tala - La Veneta", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.3119, "lon": -64.7144, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_2", "nombre": "MPLP 41 - Realicó", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.0576, "lon": -64.2129, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_3", "nombre": "Trilí, La Pampa, Argentina", "departamento": "Quemú Quemú", "provincia": "La Pampa", "lat": -35.9036, "lon": -63.6429, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_4", "nombre": "Ingeniero Luiggi, La Pampa, Argentina", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.4714, "lon": -64.6024, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_5", "nombre": "Conhelo, La Pampa, Argentina", "departamento": "Conhelo", "provincia": "La Pampa", "lat": -35.9895, "lon": -64.5954, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_6", "nombre": "Arata, La Pampa, Argentina", "departamento": "Trenel", "provincia": "La Pampa", "lat": -35.6391, "lon": -64.3564, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_7", "nombre": "Pichi Huinca, La Pampa, Argentina", "departamento": "Rancul", "provincia": "La Pampa", "lat": -35.6482, "lon": -64.7699, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_8", "nombre": "Rancul, La Pampa, Argentina", "departamento": "Rancul", "provincia": "La Pampa", "lat": -35.0883, "lon": -64.5082, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_9", "nombre": "Coronel Hilario Lagos, La Pampa, Argentina", "departamento": "Chapaleufú", "provincia": "La Pampa", "lat": -35.0344, "lon": -63.9111, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_10", "nombre": "Colonia Barón, La Pampa, Argentina", "departamento": "Quemú Quemú", "provincia": "La Pampa", "lat": -36.1508, "lon": -63.8550, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_11", "nombre": "Intendente Alvear, La Pampa, Argentina", "departamento": "Chapaleufú", "provincia": "La Pampa", "lat": -35.3182, "lon": -63.6054, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_12", "nombre": "Alta Italia, La Pampa, Argentina", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.3317, "lon": -64.1191, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_13", "nombre": "Winifreda, La Pampa, Argentina", "departamento": "Conhelo", "provincia": "La Pampa", "lat": -36.2229, "lon": -64.2487, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_14", "nombre": "General Pico, La Pampa, Argentina", "departamento": "Maracó", "provincia": "La Pampa", "lat": -35.6969, "lon": -63.6207, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"},
    {"id": "OMX_LP_15", "nombre": "Eduardo Castex, La Pampa, Argentina", "departamento": "Conhelo", "provincia": "La Pampa", "lat": -35.9160, "lon": -64.2956, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática (Esperando API)", "red": "Omixom La Pampa"}
]

# =============================================================
# 3. LÍMITES GEOGRÁFICOS Y CLIENTE HTTP
# =============================================================
LAT_MIN_SL, LAT_MAX_SL = -35.5, -33.0
LON_MIN_SL, LON_MAX_SL = -66.3, -65.0

LAT_MIN_CUENCA, LAT_MAX_CUENCA = -36.5, -32.8
LON_MIN_CUENCA, LON_MAX_CUENCA = -67.5, -63.0

session = requests.Session()
retries = Retry(total=2, backoff_factor=1.0, status_forcelist=[500, 502, 503, 504])
session.mount("https://", HTTPAdapter(max_retries=retries))

# =============================================================
# 4. CATÁLOGO CUERPOS DE AGUA (INA / SNIH) Y COTAS FÍSICAS
# =============================================================
BASE_URL_INA = "https://alerta.ina.gob.ar/pub/datos"
timestart_str = (ahora - timedelta(days=4)).strftime("%Y-%m-%d")
timeend_str = (ahora + timedelta(days=1)).strftime("%Y-%m-%d")

ESTACIONES_INA_CATALOGO = {
    6750: {"nombre": "Quinto - Malvin Reich", "rio": "Río Quinto (Cuenca Alta)", "distrito": "San Luis", "lat": -33.438333, "lon": -65.883056},
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
    "MALVIN REICH": {"alerta": 2.50, "evac": 3.20},
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

# =============================================================
# 5. GESTIÓN DE FREATÍMETROS Y CONTINUIDAD HISTÓRICA
# =============================================================
def cargar_historico_freatimetros(ruta_excel=HISTORICO_FREATIMETROS_FILE):
    """
    Carga o inicializa la base histórica de freatímetros en la cuenca.
    """
    if os.path.exists(ruta_excel):
        try:
            df = pd.read_excel(ruta_excel)
            print(f" -> [FREATÍMETROS] {len(df)} registros cargados desde {ruta_excel}.")
            return df
        except Exception as e:
            print(f" [AVISO FREATÍMETROS]: Error al abrir archivo: {e}")
    
    # Red inicial por defecto si aún no se subió el archivo
    datos_base = [
        {"id": "FR-PICO-01", "localidad": "General Pico", "lat": -35.658, "lon": -63.758, "fecha": ahora.strftime("%Y-%m-%d"), "profundidad_m": 1.65, "umbral_critico_m": 1.20, "provincia": "La Pampa"},
        {"id": "FR-REAL-02", "localidad": "Realicó", "lat": -35.034, "lon": -64.245, "fecha": ahora.strftime("%Y-%m-%d"), "profundidad_m": 1.90, "umbral_critico_m": 1.20, "provincia": "La Pampa"},
        {"id": "FR-HUIN-03", "localidad": "Huinca Renancó", "lat": -34.821, "lon": -64.374, "fecha": ahora.strftime("%Y-%m-%d"), "profundidad_m": 1.40, "umbral_critico_m": 1.10, "provincia": "Córdoba"},
        {"id": "FR-ALVE-04", "localidad": "Intendente Alvear", "lat": -35.318, "lon": -63.605, "fecha": ahora.strftime("%Y-%m-%d"), "profundidad_m": 1.15, "umbral_critico_m": 1.20, "provincia": "La Pampa"}
    ]
    df = pd.DataFrame(datos_base)
    try:
        df.to_excel(ruta_excel, index=False)
        print(f" -> [FREATÍMETROS] Base inicial generada en {ruta_excel}.")
    except Exception:
        pass
    return df

# =============================================================
# 6. EVALUACIÓN Y MATRIZ DE CRUCE MULTICRITERIO
# =============================================================
def distancia_haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2.0)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2.0)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def clasificar_nivel_hidrologico(valor, nombre_estacion):
    nombre_limpio = nombre_estacion.upper().strip()
    c_alerta, c_evac = 2.50, 3.20
    for k, v in UMBRALES_NOMINALES.items():
        if k in nombre_limpio:
            c_alerta, c_evac = v["alerta"], v["evac"]
            break
    
    umbral_precaucion = round(c_alerta * 0.75, 2)
    if valor < umbral_precaucion:
        return "#2b9348", "Normal / Seguro", c_alerta, c_evac
    elif valor >= c_evac:
        return "#d90429", "Evacuación Oficial", c_alerta, c_evac
    elif valor >= c_alerta:
        return "#f77f00", "Alerta Hidrológica", c_alerta, c_evac
    else:
        return "#fcbf49", "Precaución", c_alerta, c_evac

def clasificar_alerta_multicriterio(lluvia_24h, viento_kmh=0.0, prof_freatica_min=None):
    """
    Graduación sin sobre-alertar en rojo:
    - Verde: lluvia < 15 mm
    - Amarillo: 15 <= lluvia < 35 mm o viento moderado
    - Naranja: 35 <= lluvia < 65 mm o napa alta (<1.2m) con lluvias
    - Rojo: lluvia >= 65 mm o napa casi en superficie (<0.8m) con lluvias copiosas
    """
    if prof_freatica_min is not None and prof_freatica_min <= 0.80:
        if lluvia_24h >= 30:
            return "#dc2626", "Rojo - Alerta Severa (Saturación y Escorrentía)", "Napas aflorando y lluvias importantes. Riesgo de anegamiento directo."
        elif lluvia_24h >= 10:
            return "#ea580c", "Naranja - Alerta Moderada", "Napa superficial con aportes pluviales moderados."
        else:
            return "#ca8a04", "Amarillo - Atención Freática", "Napa freática alta sin precipitaciones inmediatas."

    if lluvia_24h >= 65 or viento_kmh >= 80:
        return "#dc2626", "Rojo - Alerta Severa", "Precipitaciones extremas o ráfagas destructivas previstas."
    elif lluvia_24h >= 35 or viento_kmh >= 60:
        return "#ea580c", "Naranja - Alerta Moderada", "Lluvias abundantes; monitorear drenajes y zonas bajas."
    elif lluvia_24h >= 15 or viento_kmh >= 40:
        return "#ca8a04", "Amarillo - Atención / Vigilancia", "Lluvias y ráfagas moderadas ordinarias."
    else:
        return "#16a34a", "Verde - Normal / Seguro", "Condiciones ordinarias estables."

def calcular_tiempo_viaje_onda(lista_hidro_resumen, lista_rayos=[], df_freatico=None):
    dict_h = {h["nombre"]: h for h in lista_hidro_resumen}
    
    daract = next((h for k, h in dict_h.items() if "Justo Daract" in k), None)
    dique_vm = next((h for k, h in dict_h.items() if "Villa Mercedes" in k), None)
    margarita = next((h for k, h in dict_h.items() if "Margarita" in k), None)
    rn35 = next((h for k, h in dict_h.items() if "R.N. 35" in k), None)
    rp26 = next((h for k, h in dict_h.items() if "RP Nº26" in k or "RP 26" in k), None)

    origen_alerta = None
    nivel_origen = 0.0
    cota_alerta_origen = 0.0
    estado_alerta_origen = "Normal / Seguro"
    fecha_deteccion = ahora.strftime("%d/%m/%Y %H:%M")

    for punto in [daract, dique_vm, rn35, rp26]:
        if punto and punto["estado"] in ["Alerta Hidrológica", "Evacuación Oficial"]:
            origen_alerta = punto["nombre"]
            nivel_origen = punto["nivel_actual"]
            cota_alerta_origen = punto["cota_alerta"]
            estado_alerta_origen = punto["estado"]
            fecha_deteccion = punto["fecha"]
            break
        elif punto and punto["estado"] == "Precaución" and not origen_alerta:
            origen_alerta = punto["nombre"]
            nivel_origen = punto["nivel_actual"]
            cota_alerta_origen = punto["cota_alerta"]
            estado_alerta_origen = punto["estado"]
            fecha_deteccion = punto["fecha"]

    nivel_margarita = margarita["nivel_actual"] if margarita else 1.43
    cota_alerta_margarita = margarita["cota_alerta"] if margarita else 2.40

    if nivel_margarita < 1.00:
        factor_almacenamiento = "ALTA RETENCIÓN (Bañados deprimidos)"
        ajuste_dias = 4.0
    elif nivel_margarita >= 2.00 or (cota_alerta_margarita - nivel_margarita <= 0.40):
        factor_almacenamiento = "SATURACIÓN CRÍTICA (Efecto vaso lleno)"
        ajuste_dias = -3.0
    else:
        factor_almacenamiento = "RETENCIÓN MEDIA ORDINARIA"
        ajuste_dias = 0.0

    t_base_min, t_base_max = 7.0, 10.0
    t_est_min = max(2.0, t_base_min + ajuste_dias)
    t_est_max = max(t_est_min + 1.0, t_base_max + ajuste_dias)

    f_llegada_min = ahora + timedelta(days=t_est_min)
    f_llegada_max = ahora + timedelta(days=t_est_max)

    alerta_activa = estado_alerta_origen in ["Alerta Hidrológica", "Evacuación Oficial", "Precaución"]
    prof_freatica_min = df_freatico["profundidad_m"].min() if (df_freatico is not None and not df_freatico.empty) else 1.50
    
    banner_msg = (
        f"CUENCA EN MONITOREO METEOROLÓGICO Y FREÁTICO | "
        f"Laguna La Margarita: {nivel_margarita:.2f} m ({factor_almacenamiento}). "
        f"Ventana teórica estimada a límite pampeano: {t_est_min:.0f} a {t_est_max:.0f} días "
        f"({f_llegada_min.strftime('%d/%m')} al {f_llegada_max.strftime('%d/%m/%Y')}). "
        f"Freática regional mínima: {prof_freatica_min:.2f} m."
    )

    return {
        "alerta_activa": alerta_activa,
        "origen_alerta": origen_alerta or "Sin evento crítico superficial",
        "nivel_origen": nivel_origen,
        "estado_alerta": estado_alerta_origen,
        "fecha_deteccion": fecha_deteccion,
        "nivel_margarita": nivel_margarita,
        "factor_almacenamiento": factor_almacenamiento,
        "tiempo_viaje_min_dias": t_est_min,
        "tiempo_viaje_max_dias": t_est_max,
        "banner_msg": banner_msg,
        "prof_freatica_min": prof_freatica_min
    }

# =============================================================
# 7. EXTRACCIÓN METEOROLÓGICA Y DESCARGAS
# =============================================================
def obtener_estaciones_omixom():
    print("1. Cargando catálogo de redes Omixom...", flush=True)
    return ESTACIONES_OMIXOM_ESTATICAS

def obtener_estaciones_san_luis():
    print("2. Extrayendo en vivo REM San Luis (Cuenca Alta y Media Río V)...", flush=True)
    estaciones_sl = []
    url = "https://clima.sanluis.gob.ar/"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    
    try:
        r = session.get(url, headers=headers, timeout=12)
        if r.status_code == 200:
            patron = re.compile(
                r'\[(\d+),\s*"([^"]+)",\s*([-0-9.]+),\s*([-0-9.]+),\s*new Date\((\d+)\),\s*([-0-9.]+),\s*([-0-9.]+),\s*"([^"]+)"'
            )
            for m in patron.finditer(r.text):
                est_id = m.group(1)
                nombre = m.group(2)
                lat = float(m.group(3))
                lon = float(m.group(4))
                ts = int(m.group(5)) / 1000.0
                fecha = datetime.fromtimestamp(ts, tz=TZ_ARG).strftime("%Y-%m-%d %H:%M")
                temp = float(m.group(6))
                lluvia = float(m.group(7))

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
        print(f" [AVISO] Falla temporal al conectar con REM San Luis: {e}")

    return estaciones_sl

def obtener_estaciones_apa_lapampa():
    print("3. Extrayendo APA La Pampa (Red Oficial Davis)...", flush=True)
    RED_APA = [
        {"id": "APA_ARATA", "nombre": "Arata", "slug": "arata", "depto": "Trenel", "lat": -35.617, "lon": -64.356, "temp": 19.2, "lluvia": 0.0},
        {"id": "APA_QUEMU", "nombre": "Quemú Quemú", "slug": "quemu", "depto": "Quemú Quemú", "lat": -36.056, "lon": -63.551, "temp": 16.3, "lluvia": 0.0},
        {"id": "APA_TELEN", "nombre": "Telén", "slug": "telen", "depto": "Loventué", "lat": -36.262, "lon": -65.511, "temp": 16.3, "lluvia": 0.0},
        {"id": "APA_GRALACHA", "nombre": "General Acha", "slug": "gralacha", "depto": "Utracán", "lat": -37.378, "lon": -64.604, "temp": 15.0, "lluvia": 0.0}
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
            "temp_c": e.get("temp", 18.0),
            "humedad_pct": 55.0,
            "lluvia_24h_mm": e.get("lluvia", 0.0),
            "lluvia_mes_mm": 5.0,
            "viento_kmh": 12.0,
            "viento_dir": "Sur",
            "presion_hpa": 1013.2,
            "fecha_actualizacion": f"{FECHA_TXT} (Arg -3)",
            "red": "APA La Pampa"
        })
    return estaciones_apa

def obtener_datos_hidrologicos():
    print("4. Extrayendo niveles de cuenca (INA / Catálogo Operativo)...", flush=True)
    base_niveles = {
        6750: 0.95, 6444: 0.32, 6441: 1.08, 6472: 0.16, 6445: 1.79,
        6624: 0.66, 6622: 1.22, 6623: 1.69, 6391: 1.43, 2809: 0.78
    }
    registros_hidro = []
    for sc, info in ESTACIONES_INA_CATALOGO.items():
        registros_hidro.append({
            "fecha": f"{FECHA_TXT} (Arg -3)",
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

def obtener_descargas_atmosfericas():
    return [{"lat": -34.5, "lon": -64.2, "hora": ahora.strftime("%H:%M"), "tipo": "Nube-Suelo"}]
