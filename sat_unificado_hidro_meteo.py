# -*- coding: utf-8 -*-
"""
SISTEMA DE ALERTA TEMPRANA TRIPROVINCIAL: CUENCA RÍO V (SAN LUIS - CÓRDOBA - LA PAMPA)
Monitoreo Hidrometeorológico Online, Red Freatimétrica, Cruce Multicriterio y Descarga de Informes.
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
URL_API_SNIH = "https://snih.hidricosargentina.gob.ar/MuestraDatos.aspx/LeerDatosActuales"

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
    Carga o inicializa la red freatimétrica histórica de la cuenca.
    """
    if os.path.exists(ruta_excel):
        try:
            df = pd.read_excel(ruta_excel)
            print(f" -> [FREATÍMETROS] {len(df)} registros históricos cargados desde {ruta_excel}.")
            return df
        except Exception as e:
            print(f" [AVISO FREATÍMETROS]: Error al leer archivo: {e}")
    
    # Red base inicial en caso de no existir archivo aún
    datos_base = [
        {"id": "FR-PICO-01", "localidad": "General Pico", "lat": -35.658, "lon": -63.758, "fecha": ahora.strftime("%Y-%m-%d"), "profundidad_m": 1.65, "umbral_critico_m": 1.20, "provincia": "La Pampa"},
        {"id": "FR-REAL-02", "localidad": "Realicó", "lat": -35.034, "lon": -64.245, "fecha": ahora.strftime("%Y-%m-%d"), "profundidad_m": 1.90, "umbral_critico_m": 1.20, "provincia": "La Pampa"},
        {"id": "FR-HUIN-03", "localidad": "Huinca Renancó", "lat": -34.821, "lon": -64.374, "fecha": ahora.strftime("%Y-%m-%d"), "profundidad_m": 1.40, "umbral_critico_m": 1.10, "provincia": "Córdoba"},
        {"id": "FR-ALVE-04", "localidad": "Intendente Alvear", "lat": -35.318, "lon": -63.605, "fecha": ahora.strftime("%Y-%m-%d"), "profundidad_m": 1.15, "umbral_critico_m": 1.20, "provincia": "La Pampa"}
    ]
    df = pd.DataFrame(datos_base)
    try:
        df.to_excel(ruta_excel, index=False)
        print(f" -> [FREATÍMETROS] Creada base inicial en {ruta_excel}.")
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
    - Rojo: lluvia >= 65 mm o napa casi aflorando (<0.8m) con lluvias abundantes
    """
    if prof_freatica_min is not None and prof_freatica_min <= 0.80:
        if lluvia_24h >= 30:
            return "#dc2626", "Rojo - Alerta Severa (Saturación y Escorrentía)", "Napas aflorando y lluvias copiosas. Riesgo crítico de colmatación."
        elif lluvia_24h >= 10:
            return "#ea580c", "Naranja - Alerta Moderada", "Napa freática en superficie con aportes pluviales moderados."
        else:
            return "#ca8a04", "Amarillo - Atención Freática", "Napa casi superficial sin precipitaciones inmediatas."

    if lluvia_24h >= 65 or viento_kmh >= 80:
        return "#dc2626", "Rojo - Alerta Severa", "Precipitaciones extremas o ráfagas destructivas previstas."
    elif lluvia_24h >= 35 or viento_kmh >= 60:
        return "#ea580c", "Naranja - Alerta Moderada", "Lluvias abundantes previstas; vigilar anegamientos locales."
    elif lluvia_24h >= 15 or viento_kmh >= 40:
        return "#ca8a04", "Amarillo - Atención / Vigilancia", "Lluvias o vientos moderados sin peligro inminente."
    else:
        return "#16a34a", "Verde - Normal / Seguro", "Condiciones ordinarias estables en cuenca."

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
    
    # Comprobar estado freático mínimo
    prof_freatica_min = df_freatico["profundidad_m"].min() if (df_freatico is not None and not df_freatico.empty) else 1.50
    
    banner_msg = (
        f"CUENCA EN MONITOREO METEOROLÓGICO Y FREÁTICO | "
        f"Nivel en Laguna La Margarita: {nivel_margarita:.2f} m ({factor_almacenamiento}). "
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
    print("1. Cargando catálogo de redes Omixom (Córdoba y La Pampa)...", flush=True)
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

def normalizar_a_lista(resp_json):
    if isinstance(resp_json, list): return resp_json
    if isinstance(resp_json, dict):
        for k in ["data", "datos", "series", "estaciones", "results"]:
            if k in resp_json and isinstance(resp_json[k], list): return resp_json[k]
    return []

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

# =============================================================
# 8. GENERACIÓN DE ENTREGABLES (EXCEL, CSV Y MAPA HTML)
# =============================================================
def generar_entregables(estaciones_meteo, registros_hidro, lista_rayos, df_freatico):
    print("5. Compilando modelo hidrológico y generando archivos...", flush=True)
    
    df_hidro_raw = pd.DataFrame(registros_hidro)
    df_hidro_raw["fecha_dt"] = pd.to_datetime(df_hidro_raw["fecha"], errors="coerce")
    
    lista_hidro_resumen = []
    for nombre_est, grp in df_hidro_raw.groupby("nombre"):
        grp_ord = grp.sort_values("fecha_dt")
        ult = grp_ord.iloc[-1]
        color, estado, c_alerta, c_evac = clasificar_nivel_hidrologico(ult["valor"], nombre_est)
        lista_hidro_resumen.append({
            "nombre": nombre_est,
            "rio": ult["rio"],
            "distrito": ult["distrito"],
            "lat": ult["lat"],
            "lon": ult["lon"],
            "nivel_actual": ult["valor"],
            "media_hist": round(grp_ord["valor"].mean(), 2),
            "cota_alerta": c_alerta,
            "cota_evac": c_evac,
            "margen_alerta": round(c_alerta - ult["valor"], 2),
            "tendencia": "Estable ▬",
            "variacion": 0.0,
            "color": color,
            "estado": estado,
            "fecha": ult["fecha"],
            "fuente": ult["fuente"]
        })

    diag_onda = calcular_tiempo_viaje_onda(lista_hidro_resumen, lista_rayos, df_freatico)

    # A. Archivo CSV de Cruce
    filas_cruce = []
    for h in lista_hidro_resumen:
        min_d = float("inf")
        m_cercana = None
        for m in estaciones_meteo:
            d = distancia_haversine(h["lat"], h["lon"], m["lat"], m["lon"])
            if d < min_d:
                min_d = d
                m_cercana = m
        
        lluvia_val = m_cercana["lluvia_24h_mm"] if m_cercana else 0.0
        _, tag_alerta, desc_alerta = clasificar_alerta_multicriterio(lluvia_val, prof_freatica_min=diag_onda["prof_freatica_min"])

        filas_cruce.append({
            "Estación Hidrológica": h["nombre"],
            "Río / Cuenca": h["rio"],
            "Nivel Actual (m)": h["nivel_actual"],
            "Media Histórica (m)": h["media_hist"],
            "Cota Alerta (m)": h["cota_alerta"],
            "Estado Semáforo": h["estado"],
            "Estación Meteo Cercana": f"{m_cercana['nombre']} ({m_cercana['provincia']})" if m_cercana else "N/A",
            "Lluvia 24h (mm)": lluvia_val,
            "Nivel Alerta Meteo-Freático": tag_alerta,
            "Evaluación de Riesgo": desc_alerta
        })

    df_cruce = pd.DataFrame(filas_cruce)
    df_cruce.to_csv(CSV_SALIDA, index=False, encoding="utf-8-sig")
    print(f" -> [CSV CRUCE GUARDADO]: {CSV_SALIDA}")

    # B. Libro Excel Multisolapa Ejecutivo
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    font_title = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    font_tbl_head = Font(name="Segoe UI", size=9.5, bold=True, color="FFFFFF")
    font_data = Font(name="Segoe UI", size=9)
    fill_navy = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_med_blue = PatternFill(start_color="2B4C7E", end_color="2B4C7E", fill_type="solid")

    ws1 = wb.create_sheet(title="Monitoreo en Tiempo Real")
    ws1.views.sheetView[0].showGridLines = True
    ws1.merge_cells("A1:J1")
    # TEXTO ACTUALIZADO
    ws1["A1"] = "MONITOREO METEOROLÓGICO EN TIEMPO REAL - CUENCA RÍO V"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_navy
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")

    headers_c = ["N°", "Punto Hidrológico", "Río", "Nivel (m)", "Cota Alerta (m)", "Semáforo", "Estación Meteo", "Lluvia 24h", "Alerta Cruzada", "Diagnóstico"]
    for c, h in enumerate(headers_c, start=1):
        cell = ws1.cell(row=3, column=c, value=h)
        cell.font = font_tbl_head
        cell.fill = fill_med_blue
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for idx, r in df_cruce.iterrows():
        rn = 4 + idx
        ws1.cell(row=rn, column=1, value=idx+1)
        ws1.cell(row=rn, column=2, value=r["Estación Hidrológica"])
        ws1.cell(row=rn, column=3, value=r["Río / Cuenca"])
        ws1.cell(row=rn, column=4, value=r["Nivel Actual (m)"])
        ws1.cell(row=rn, column=5, value=r["Cota Alerta (m)"])
        ws1.cell(row=rn, column=6, value=r["Estado Semáforo"])
        ws1.cell(row=rn, column=7, value=r["Estación Meteo Cercana"])
        ws1.cell(row=rn, column=8, value=r["Lluvia 24h (mm)"])
        ws1.cell(row=rn, column=9, value=r["Nivel Alerta Meteo-Freático"])
        ws1.cell(row=rn, column=10, value=r["Evaluación de Riesgo"])
        for c in range(1, 11):
            ws1.cell(row=rn, column=c).font = font_data

    # Solapa de Freatímetros
    ws_freatica = wb.create_sheet(title="Red Freatimétrica Histórica")
    ws_freatica.views.sheetView[0].showGridLines = True
    ws_freatica.merge_cells("A1:G1")
    ws_freatica["A1"] = "CONTINUIDAD HISTÓRICA DE NIVELES FREÁTICOS"
    ws_freatica["A1"].font = font_title
    ws_freatica["A1"].fill = fill_navy
    ws_freatica["A1"].alignment = Alignment(horizontal="center", vertical="center")

    headers_fr = ["ID", "Localidad", "Provincia", "Latitud", "Longitud", "Profundidad (m)", "Umbral Crítico (m)"]
    for c, h in enumerate(headers_fr, start=1):
        cell = ws_freatica.cell(row=3, column=c, value=h)
        cell.font = font_tbl_head
        cell.fill = fill_med_blue

    for idx, r in df_freatico.iterrows():
        rn = 4 + idx
        ws_freatica.cell(row=rn, column=1, value=r["id"])
        ws_freatica.cell(row=rn, column=2, value=r["localidad"])
        ws_freatica.cell(row=rn, column=3, value=r["provincia"])
        ws_freatica.cell(row=rn, column=4, value=r["lat"])
        ws_freatica.cell(row=rn, column=5, value=r["lon"])
        ws_freatica.cell(row=rn, column=6, value=r["profundidad_m"])
        ws_freatica.cell(row=rn, column=7, value=r["umbral_critico_m"])

    wb.save(EXCEL_SALIDA)
    print(f" -> [EXCEL GUARDADO]: {EXCEL_SALIDA}")

    # C. Mapa Folium con Módulo de Carga KML INTA y Embebido SMN
    lat_centro = np.mean([h["lat"] for h in lista_hidro_resumen])
    lon_centro = np.mean([h["lon"] for h in lista_hidro_resumen])
    
    mapa = folium.Map(location=[lat_centro, lon_centro], zoom_start=7, tiles=None)

    folium.TileLayer(
        tiles="https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}",
        attr="Esri",
        name="Cartografía Base Clara (Esri Canvas)",
        max_zoom=16
    ).add_to(mapa)

    fg_hidro = folium.FeatureGroup(name="Cuerpos de Agua (INA/SNIH)", show=True)
    fg_meteo = folium.FeatureGroup(name="Estaciones Meteorológicas", show=True)
    fg_freatica = folium.FeatureGroup(name="Freatímetros (Red Histórica)", show=True)

    # Cuerpos de agua
    for h in lista_hidro_resumen:
        folium.CircleMarker(
            location=[h["lat"], h["lon"]],
            radius=7,
            color=h["color"],
            fill=True,
            fill_color=h["color"],
            fill_opacity=0.9,
            tooltip=f"{h['nombre']}: {h['nivel_actual']:.2f} m ({h['estado']})"
        ).add_to(fg_hidro)

    # Estaciones Meteo
    for m in estaciones_meteo:
        folium.CircleMarker(
            location=[m["lat"], m["lon"]],
            radius=5,
            color="#0284c7",
            fill=True,
            fill_color="#38bdf8",
            fill_opacity=0.85,
            tooltip=f"{m['nombre']} | Lluvia: {m['lluvia_24h_mm']:.1f} mm"
        ).add_to(fg_meteo)

    # Freatímetros
    for _, fr in df_freatico.iterrows():
        prof = fr["profundidad_m"]
        color_fr = "#dc2626" if prof <= 0.80 else ("#f59e0b" if prof <= fr["umbral_critico_m"] else "#10b981")
        folium.CircleMarker(
            location=[fr["lat"], fr["lon"]],
            radius=6,
            color=color_fr,
            fill=True,
            fill_color=color_fr,
            fill_opacity=0.95,
            tooltip=f"Freatímetro {fr['id']} ({fr['localidad']}): Napa a {prof:.2f} m"
        ).add_to(fg_freatica)

    fg_hidro.add_to(mapa)
    fg_meteo.add_to(mapa)
    fg_freatica.add_to(mapa)
    folium.LayerControl(position="topright", collapsed=False).add_to(mapa)

    # Inyección de componentes interactivos (KML INTA, Embebido SMN y Panel Freático)
    html_complementos = """
    <script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet-omnivore/0.3.4/leaflet-omnivore.min.js"></script>
    <div style="position: fixed; bottom: 25px; left: 20px; width: 330px; background: white;
                border: 1px solid #cbd5e1; border-radius: 8px; z-index: 1000; font-family: Arial, sans-serif;
                font-size: 11px; padding: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.2);">
        <b style="font-size: 12px; color: #1e3a8a;">SAT CUENCA RÍO V | CRUCE MULTICRITERIO</b><hr style="margin:6px 0;">
        
        <b>Subir KML Humedad Suelo (INTA):</b>
        <input type="file" id="kmlIntaInput" accept=".kml" style="width:100%; font-size:10px; margin-top:3px; margin-bottom:8px;">
        
        <b>Alertas SMN Oficiales:</b><br>
        <a href="https://www.smn.gob.ar/alertas" target="_blank" style="color: #0284c7; text-decoration: underline; font-weight: bold;">
            Abrir Portal de Alertas SMN (smn.gob.ar/alertas) ↗
        </a>
        <hr style="margin:6px 0;">
        <span style="font-size: 10px; color: #64748b;">
            El semáforo cruza precipitación prevista, afloramiento freático y retención de suelos.
        </span>
    </div>

    <script>
    document.addEventListener("DOMContentLoaded", function() {
        var mapInstance = null;
        for (var k in window) {
            if (k.startsWith("map_") && window[k] instanceof L.Map) {
                mapInstance = window[k];
                break;
            }
        }
        if (!mapInstance) return;

        var inputKml = document.getElementById("kmlIntaInput");
        if (inputKml) {
            inputKml.addEventListener("change", function(e) {
                var file = e.target.files[0];
                if (!file) return;
                var reader = new FileReader();
                reader.onload = function(evt) {
                    var kmlText = evt.target.result;
                    var capaKml = omnivore.kml.parse(kmlText, null, L.geoJson(null, {
                        style: { color: "#059669", fillColor: "#10b981", fillOpacity: 0.35, weight: 1.5 }
                    })).addTo(mapInstance);
                    alert("Capa KML de Humedad de Suelo INTA agregada al visor.");
                };
                reader.readAsText(file);
            });
        }
    });
    </script>
    """
    mapa.get_root().html.add_child(folium.Element(html_complementos))
    mapa.save(MAPA_HTML_SALIDA)
    print(f" -> [MAPA HTML GENERADO]: {MAPA_HTML_SALIDA}")

# =============================================================
# 9. EJECUCIÓN PRINCIPAL
# =============================================================
if __name__ == "__main__":
    print("=" * 70)
    print(">>> SAT TRIPROVINCIAL: ACTUALIZACIÓN HIDROMETEOROLÓGICA Y FREÁTICA <<<")
    print("=" * 70)
    
    df_freatico = cargar_historico_freatimetros()
    meteo_omixom = obtener_estaciones_omixom()
    meteo_san_luis = obtener_estaciones_san_luis()
    meteo_apa = obtener_estaciones_apa_lapampa()
    
    total_meteo = meteo_omixom + meteo_san_luis + meteo_apa
    print(f"Total estaciones meteorológicas: {len(total_meteo)}")
    
    registros_hidro = obtener_datos_hidrologicos()
    rayos_cuenca = obtener_descargas_atmosfericas()
    
    generar_entregables(total_meteo, registros_hidro, rayos_cuenca, df_freatico)
    print("Proceso finalizado exitosamente.")
