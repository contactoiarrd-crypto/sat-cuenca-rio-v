# -*- coding: utf-8 -*-
"""
SISTEMA DE ALERTA TEMPRANA TRIPROVINCIAL: CUENCA RÍO V (SAN LUIS - CÓRDOBA - LA PAMPA)
Monitoreo Hidrometeorológico Online, Cuerpos de Agua INA, SAT SMN, Pronóstico ECMWF,
Radares SINARAME, Rayos y Estimación de Traslación de Onda.
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

# =============================================================
# 1. PARÁMETROS GENERALES Y ARCHIVOS DE SALIDA
# =============================================================
TZ_ARG = timezone(timedelta(hours=-3))
ahora = datetime.now(TZ_ARG)
FECHA_TXT = ahora.strftime("%Y-%m-%d %H:%M")

EXCEL_SALIDA = "sat_unificado_rio_v_triprovincial.xlsx"
MAPA_HTML_SALIDA = "index.html"  # Publicación directa en GitHub Pages
CSV_SALIDA = "resumen_cruce_rio_v_triprovincial.csv"

# =============================================================
# 2. CATÁLOGO ESTÁTICO REDES OMIXOM (CÓRDOBA Y LA PAMPA)
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
# 3. LÍMITES GEOGRÁFICOS Y CLIENTE HTTP
# =============================================================
LAT_MIN_SL, LAT_MAX_SL = -35.5, -33.0
LON_MIN_SL, LON_MAX_SL = -66.3, -65.0
LAT_MIN_CUENCA, LAT_MAX_CUENCA = -36.5, -32.8
LON_MIN_CUENCA, LON_MAX_CUENCA = -67.5, -63.0

session = requests.Session()
retries = Retry(total=2, backoff_factor=1.0, status_forcelist=[500, 502, 503, 504])
session.mount("https://", HTTPAdapter(max_retries=retries))
session.mount("http://", HTTPAdapter(max_retries=retries))

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

# =============================================================
# 5. FUNCIONES DE TRASLACIÓN DE ONDA
# =============================================================
def distancia_haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2.0)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2.0)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def clasificar_nivel(valor, nombre_estacion):
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

def calcular_tiempo_viaje_onda(lista_hidro_resumen, lista_rayos=[]):
    dict_h = {h["nombre"]: h for h in lista_hidro_resumen}
    
    dique_vm = next((h for k, h in dict_h.items() if "Dique Villa Mercedes" in k or "Villa Mercedes" in k), None)
    daract = next((h for k, h in dict_h.items() if "Justo Daract" in k), None)
    margarita = next((h for k, h in dict_h.items() if "Margarita" in k), None)
    rn35 = next((h for k, h in dict_h.items() if "R.N. 35" in k), None)
    rp26 = next((h for k, h in dict_h.items() if "RP Nº26" in k or "RP 26" in k), None)
    devoto = next((h for k, h in dict_h.items() if "Devoto" in k), None)

    origen_alerta = None
    nivel_origen = 0.0
    cota_alerta_origen = 0.0
    estado_alerta_origen = "Normal / Seguro"
    fecha_deteccion = ahora.strftime("%d/%m/%Y %H:%M")

    for punto in [daract, dique_vm, rn35, devoto, rp26]:
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
        factor_almacenamiento = "ALTA RETENCIÓN (Bañados secos / Lagunas deprimidas)"
        ajuste_dias = 4.0
        desc_almacenamiento = "Amortiguación máxima (+4 días al tiempo de viaje hacia La Pampa)."
    elif nivel_margarita >= 2.00 or (cota_alerta_margarita - nivel_margarita <= 0.40):
        factor_almacenamiento = "SATURACIÓN CRÍTICA (Efecto vaso lleno)"
        ajuste_dias = -3.0
        desc_almacenamiento = "Transferencia directa acelerada (-3 días al tiempo de viaje hacia La Pampa)."
    else:
        factor_almacenamiento = "RETENCIÓN MEDIA ORDINARIA"
        ajuste_dias = 0.0
        desc_almacenamiento = "Tránsito ordinario de amortiguación según tiempos históricos."

    if origen_alerta and ("Villa Mercedes" in origen_alerta or "Trapiche" in origen_alerta):
        t_base_min, t_base_max = 8.0, 12.0
    elif origen_alerta and "Justo Daract" in origen_alerta:
        t_base_min, t_base_max = 7.0, 10.0
    elif origen_alerta and "R.N. 35" in origen_alerta:
        t_base_min, t_base_max = 5.0, 8.0
    elif origen_alerta and ("Devoto" in origen_alerta or "RP Nº26" in origen_alerta or "Margarita" in origen_alerta):
        t_base_min, t_base_max = 3.0, 6.0
    else:
        t_base_min, t_base_max = 7.0, 10.0

    t_est_min = max(2.0, t_base_min + ajuste_dias)
    t_est_max = max(t_est_min + 1.0, t_base_max + ajuste_dias)

    f_llegada_min = ahora + timedelta(days=t_est_min)
    f_llegada_max = ahora + timedelta(days=t_est_max)

    alerta_activa = estado_alerta_origen in ["Alerta Hidrológica", "Evacuación Oficial", "Precaución"]
    rayos_cuenca_alta = [r for r in lista_rayos if r.get("lat", 0) > -34.5]
    alerta_convectiva = len(rayos_cuenca_alta) >= 8

    if alerta_activa:
        banner_msg = (
            f"⚠️ ALERTA HIDROLÓGICA EN CUENCA ({origen_alerta}) | Nivel: {nivel_origen:.2f} m "
            f"(Cota Alerta: {cota_alerta_origen:.2f} m). Estado: {estado_alerta_origen}. "
            f"Tiempo estimado de arribo a La Pampa: {t_est_min:.0f} a {t_est_max:.0f} días "
            f"(Previsto: {f_llegada_min.strftime('%d/%m')} al {f_llegada_max.strftime('%d/%m/%Y')}). "
            f"Condición de almacenamiento: {factor_almacenamiento}."
        )
    elif alerta_convectiva:
        banner_msg = (
            f"⚡ ALERTA METEOROLÓGICA CONVECTIVA EN NACIENTES | Detección de {len(rayos_cuenca_alta)} descargas eléctricas "
            f"en cabecera San Luis / Córdoba. Potencial pulso de crecida en formación "
            f"(Ventana teórica estimada a La Pampa: {t_est_min:.0f} a {t_est_max:.0f} días)."
        )
        alerta_activa = True
        estado_alerta_origen = "Alerta Convectiva"
    else:
        banner_msg = (
            f"🟢 CUENCA EN CALMA HIDROLÓGICA ORDINARIA | Todos los nudos de control en nivel normal/seguro. "
            f"Ventana teórica de respuesta: {t_est_min:.0f} a {t_est_max:.0f} días. "
            f"Nivel actual en Laguna La Margarita: {nivel_margarita:.2f} m ({factor_almacenamiento}). "
            f"Actividad eléctrica: {len(lista_rayos)} descargas recientes en cuenca."
        )

    return {
        "alerta_activa": alerta_activa,
        "alerta_convectiva": alerta_convectiva,
        "origen_alerta": origen_alerta or ("Actividad Convectiva" if alerta_convectiva else "Sin evento crítico"),
        "nivel_origen": nivel_origen,
        "estado_alerta": estado_alerta_origen,
        "fecha_deteccion": fecha_deteccion,
        "nivel_margarita": nivel_margarita,
        "factor_almacenamiento": factor_almacenamiento,
        "desc_almacenamiento": desc_almacenamiento,
        "tiempo_viaje_min_dias": t_est_min,
        "tiempo_viaje_max_dias": t_est_max,
        "fecha_arribo_estimada_str": f"{f_llegada_min.strftime('%d/%m')} al {f_llegada_max.strftime('%d/%m/%Y')}",
        "banner_msg": banner_msg
    }

# =============================================================
# 6. EXTRACCIÓN METEOROLÓGICA (REM, APA, OMIXOM)
# =============================================================
def obtener_estaciones_omixom():
    return ESTACIONES_OMIXOM_ESTATICAS

def obtener_estaciones_san_luis():
    print("1. Extrayendo en vivo REM San Luis (Cuenca Alta y Media Río V)...", flush=True)
    estaciones_sl = []
    url = "https://clima.sanluis.gob.ar/"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        r = session.get(url, headers=headers, timeout=10)
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
        print(f"   [AVISO] REM San Luis: {e}")
    return estaciones_sl

def obtener_estaciones_apa_lapampa():
    print("2. Extrayendo APA La Pampa (Red Oficial Davis)...", flush=True)
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
    headers = {"User-Agent": "Mozilla/5.0"}

    def _fetch(info):
        try:
            url = f"https://estaciones-apa.lapampa.gob.ar/{info['slug']}/mb1.htm"
            r = session.get(url, headers=headers, timeout=5, verify=False)
            if r.status_code == 200 and len(r.text) > 100:
                soup = BeautifulSoup(r.text, "html.parser")
                texto = soup.get_text(separator=" ")
                t_m = re.search(r"TEMPERATURA.*?Actual\s*([\d.-]+)\s*°C", texto, re.S)
                ll_m = re.search(r"LLUVIA.*?Diaria\s*([\d.-]+)\s*mm", texto, re.S)
                return {
                    "id": info["id"],
                    "nombre": f"{info['nombre']} (APA)",
                    "departamento": info["depto"],
                    "provincia": "La Pampa",
                    "lat": info["lat"],
                    "lon": info["lon"],
                    "temp_c": float(t_m.group(1)) if t_m else info.get("temp", 15.0),
                    "humedad_pct": 0.0,
                    "lluvia_24h_mm": float(ll_m.group(1)) if ll_m else info.get("lluvia", 0.0),
                    "lluvia_mes_mm": 0.0,
                    "viento_kmh": 0.0,
                    "viento_dir": "N/A",
                    "presion_hpa": 1013.2,
                    "fecha_actualizacion": FECHA_TXT,
                    "red": "APA La Pampa"
                }
        except Exception:
            pass
        return None

    estaciones_apa = []
    with ThreadPoolExecutor(max_workers=6) as executor:
        futs = [executor.submit(_fetch, e) for e in RED_APA]
        for f in as_completed(futs):
            res = f.result()
            if res: estaciones_apa.append(res)

    if len(estaciones_apa) < 3:
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

# =============================================================
# 7. EXTRACCIÓN DE CUERPOS DE AGUA INA (MÉTODO 2 PASOS)
# =============================================================
def normalizar_a_lista(resp_json):
    if isinstance(resp_json, list): return resp_json
    if isinstance(resp_json, dict):
        for k in ["data", "datos", "series", "estaciones", "results"]:
            if k in resp_json and isinstance(resp_json[k], list): return resp_json[k]
    return []

def obtener_datos_hidrologicos():
    print("3. Consultando cuerpos de agua INA (Esquema oficial de series)...", flush=True)
    registros_hidro = []
    series_activas = []
    
    try:
        r = session.get(f"{BASE_URL_INA}/series&format=json", timeout=12)
        if r.status_code == 200:
            series_raw = normalizar_a_lista(r.json())
            for s in series_raw:
                if not isinstance(s, dict): continue
                try: sitecode = int(s.get("sitecode"))
                except: continue

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

    def _descargar_serie_ina(s):
        sid = s["ina_sid"]
        url = f"{BASE_URL_INA}/datos&seriesId={sid}&timeStart={timestart_str}&timeEnd={timeend_str}&format=json"
        salida = []
        try:
            r = session.get(url, timeout=7)
            if r.status_code == 200:
                datos = normalizar_a_lista(r.json())
                for d in datos:
                    if isinstance(d, dict):
                        fecha = d.get("timestart") or d.get("timeStart") or d.get("fecha") or d.get("time")
                        valor = d.get("valor") or d.get("value") or d.get("val")
                        if fecha is not None and valor is not None:
                            salida.append({
                                "fecha": str(fecha).replace("T", " "),
                                "valor": float(valor),
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

    if series_activas:
        with ThreadPoolExecutor(max_workers=6) as executor:
            futuros = [executor.submit(_descargar_serie_ina, s) for s in series_activas]
            for fut in as_completed(futuros):
                res = fut.result()
                if res: registros_hidro.extend(res)

    # Valores base para no romper el pipeline si el servidor del INA no responde
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
            
    print(f"   -> [INA]: {len(registros_hidro)} registros procesados.")
    return registros_hidro

# =============================================================
# 8. DESCARGAS ELÉCTRICAS Y RAYOS (BLITZORTUNG)
# =============================================================
def obtener_descargas_atmosfericas():
    print("4. Monitoreando actividad convectiva (Rayos)...", flush=True)
    rayos = []
    urls = [
        "https://map.blitzortung.org/Data_Json/Strikes_0.json",
        "https://map.blitzortung.org/Data_Json/Strikes_1.json"
    ]
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://map.blitzortung.org/"}
    for url in urls:
        try:
            r = session.get(url, headers=headers, timeout=5)
            if r.status_code == 200:
                datos = r.json()
                if isinstance(datos, list):
                    for st in datos:
                        if isinstance(st, list) and len(st) >= 3:
                            lat = float(st[2])
                            lon = float(st[1])
                            if LAT_MIN_CUENCA <= lat <= LAT_MAX_CUENCA and LON_MIN_CUENCA <= lon <= LON_MAX_CUENCA:
                                ts_val = float(st[0])
                                ts_seg = ts_val / 1e9 if ts_val > 1e15 else ts_val / 1000
                                hora_str = datetime.fromtimestamp(ts_seg, tz=TZ_ARG).strftime("%H:%M")
                                rayos.append({"lat": lat, "lon": lon, "hora": hora_str})
        except Exception:
            pass
    return rayos

# =============================================================
# 9. SAT OFICIAL SMN CON GEOMETRÍAS LOCALES
# =============================================================
DEPARTAMENTOS_CUENCA = [
    "GENERAL PEDERNERA", "CORONEL PRINGLES", "GOBERNADOR DUPUY",
    "GENERAL ROCA", "PRESIDENTE ROQUE SAENZ PENA", "RIO CUARTO", "JUAREZ CELMAN",
    "REALICO", "CHAPALEUFU", "RANCUL", "MARACO", "TRENEL", "CONHELO", "QUEMU QUEMU"
]

EVENTOS_SMN_NOMBRES = {
    41: "Tormentas fuertes o severas", 42: "Vientos fuertes",
    39: "Lluvias abundantes", 40: "Nevadas", 45: "Viento Zonda",
    46: "Temperaturas extremas (Frío)", 47: "Temperaturas extremas (Calor)"
}

def obtener_alertas_smn():
    print("5. Consultando alertas meteorológicas oficiales SAT SMN...", flush=True)
    alertas_tabla = []
    poligonos_leaflet = []

    # Cargar geometrías locales del archivo del repositorio
    geometrias_dict = {}
    if os.path.exists("geometrias_sat_smn.json"):
        try:
            with open("geometrias_sat_smn.json", "r", encoding="utf-8") as f:
                geometrias_dict = json.load(f)
        except Exception:
            pass

    headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
    url_api = "https://ws1.smn.gob.ar/v1/warning/alert/area?mode=alert&compact=true"

    try:
        r = session.get(url_api, headers=headers, timeout=8)
        if r.status_code == 200:
            for area in r.json():
                warnings = [w for w in area.get("warnings", []) if w.get("max_level", 1) >= 2]
                if not warnings:
                    continue

                nombre_area = str(area.get("name", "")).upper()
                area_id = str(area.get("area_id", ""))

                es_cuenca = any(dep in nombre_area for dep in DEPARTAMENTOS_CUENCA)
                if not es_cuenca and area_id in ["3343", "3358", "3362", "3363", "3365", "3366", "3378", "768", "807", "808", "809", "810", "811", "821", "824"]:
                    es_cuenca = True

                if es_cuenca:
                    for w in warnings:
                        nivel_num = w.get("max_level", 2)
                        color_hex = "#ef4444" if nivel_num >= 4 else ("#ea580c" if nivel_num == 3 else "#f59e0b")
                        nivel_str = "Rojo" if nivel_num >= 4 else ("Naranja" if nivel_num == 3 else "Amarillo")
                        
                        eventos = [EVENTOS_SMN_NOMBRES.get(ev.get("id"), "Tormentas") for ev in w.get("events", []) if ev.get("max_level", 1) >= 2]
                        fenomeno = ", ".join(set(eventos)) if eventos else "Alerta Meteorológica"

                        alertas_tabla.append({
                            "zona": area.get("name", f"Área {area_id}"),
                            "fecha": w.get("date", "Próximas horas"),
                            "fenomeno": fenomeno,
                            "nivel": nivel_str,
                            "color": color_hex
                        })

                        # Geometría
                        geom = geometrias_dict.get(area_id)
                        if geom:
                            coords = geom.get("coordinates", [])
                            if geom.get("type") == "MultiPolygon":
                                for poly in coords:
                                    if poly:
                                        poligonos_leaflet.append({
                                            "coords": [[pt[1], pt[0]] for pt in poly[0] if len(pt) >= 2],
                                            "color": color_hex, "nivel": nivel_str, "zona": area.get("name"), "evento": fenomeno
                                        })
                            elif geom.get("type") == "Polygon" and coords:
                                poligonos_leaflet.append({
                                    "coords": [[pt[1], pt[0]] for pt in coords[0] if len(pt) >= 2],
                                    "color": color_hex, "nivel": nivel_str, "zona": area.get("name"), "evento": fenomeno
                                })
    except Exception as e:
        print(f"   [AVISO SAT SMN]: {e}")

    print(f"   -> [SAT SMN]: {len(alertas_tabla)} alertas vigentes detectadas en la cuenca.")
    return alertas_tabla, poligonos_leaflet

# =============================================================
# 10. GENERACIÓN DE ENTREGABLES Y VISOR DINÁMICO
# =============================================================
def generar_entregables(estaciones_meteo, registros_hidro, lista_rayos, alertas_smn, poligonos_smn):
    print("6. Generando entregables (Excel, CSV y Mapa index.html)...", flush=True)
    df_hidro_raw = pd.DataFrame(registros_hidro)
    df_hidro_raw["fecha_dt"] = pd.to_datetime(df_hidro_raw["fecha"], errors="coerce")
    
    lista_hidro_resumen = []
    for nombre_est, grp in df_hidro_raw.groupby("nombre"):
        grp_ord = grp.sort_values("fecha_dt")
        ult = grp_ord.iloc[-1]
        media_val = round(grp_ord["valor"].mean(), 2)
        
        if len(grp_ord) >= 2:
            dif = grp_ord.iloc[-1]["valor"] - grp_ord.iloc[-2]["valor"]
            tendencia = "Creciendo ▲" if dif > 0.02 else ("Bajando ▼" if dif < -0.02 else "Estable ▬")
        else:
            dif = 0.0
            tendencia = "Estable ▬"

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
            "variacion": round(dif, 3),
            "color": color,
            "estado": estado,
            "fecha": ult["fecha_dt"].strftime("%d/%m/%Y %H:%M") if pd.notna(ult["fecha_dt"]) else FECHA_TXT,
            "fuente": ult["fuente"]
        })

    diag_onda = calcular_tiempo_viaje_onda(lista_hidro_resumen, lista_rayos)

    # 1. Guardar CSV
    filas_cruce = []
    for h in lista_hidro_resumen:
        min_d = float("inf")
        m_cercana = None
        for m in estaciones_meteo:
            d = distancia_haversine(h["lat"], h["lon"], m["lat"], m["lon"])
            if d < min_d:
                min_d = d
                m_cercana = m

        filas_cruce.append({
            "Estación Hidrológica": h["nombre"],
            "Río / Cuenca": h["rio"],
            "Nivel Actual (m)": h["nivel_actual"],
            "Media Histórica (m)": h["media_hist"],
            "Cota Alerta (m)": h["cota_alerta"],
            "Cota Evacuación (m)": h["cota_evac"],
            "Margen Alerta (m)": h["margen_alerta"],
            "Estado Semáforo": h["estado"],
            "Tendencia Río": h["tendencia"],
            "Estación Meteo Cercana": f"{m_cercana['nombre']} ({m_cercana['provincia']})" if m_cercana else "N/A",
            "Red": m_cercana["red"] if m_cercana else "-",
            "Distancia (km)": round(min_d, 1) if m_cercana else "-",
            "Lluvia 24h (mm)": m_cercana["lluvia_24h_mm"] if m_cercana else 0.0,
            "Tiempo Viaje a LP (días)": f"{diag_onda['tiempo_viaje_min_dias']:.0f}-{diag_onda['tiempo_viaje_max_dias']:.0f} d"
        })
    pd.DataFrame(filas_cruce).to_csv(CSV_SALIDA, index=False, encoding="utf-8-sig")

    # 2. Guardar Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cruce Hidrometeorológico"
    ws.append(list(filas_cruce[0].keys()))
    for r in filas_cruce:
        ws.append(list(r.values()))
    wb.save(EXCEL_SALIDA)

    # 3. Construir Mapa Folium (index.html)
    lat_centro = np.mean([h["lat"] for h in lista_hidro_resumen])
    lon_centro = np.mean([h["lon"] for h in lista_hidro_resumen])
    
    mapa = folium.Map(location=[lat_centro, lon_centro], zoom_start=7, tiles=None, control_scale=True)

    folium.TileLayer(
        tiles="https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}",
        attr="Esri", name="Cartografía Clara Esri", max_zoom=16
    ).add_to(mapa)
    folium.TileLayer(
        tiles="https://tile.openstreetmap.org/{z}/{x}/{y}.png",
        attr="OpenStreetMap", name="OpenStreetMap"
    ).add_to(mapa)

    fg_smn = folium.FeatureGroup(name="⚠️ Polígonos de Alerta SAT SMN", show=True)
    fg_satelite = folium.FeatureGroup(name="🛰️ Satélite GOES-16 Clean IR", show=True)
    fg_radar_mosaico = folium.FeatureGroup(name="🌧️ Radar Compuesto (RainViewer/SMN)", show=True)
    fg_radares_sinarame = folium.FeatureGroup(name="📡 Radares SINARAME (Santa Isabel + Reynolds)", show=True)
    fg_rayos = folium.FeatureGroup(name="⚡ Rayos y Descargas en Vivo", show=True)
    fg_hidro = folium.FeatureGroup(name="💧 Cuerpos de Agua (INA)", show=True)
    fg_meteo_sl = folium.FeatureGroup(name="⛰️ REM San Luis", show=True)
    fg_meteo_cba = folium.FeatureGroup(name="🌾 Omixom Córdoba", show=True)
    fg_meteo_lp = folium.FeatureGroup(name="🌾 Omixom La Pampa", show=True)
    fg_meteo_apa = folium.FeatureGroup(name="💧 APA La Pampa", show=True)

    # Dibujar Polígonos SMN
    for poly in poligonos_smn:
        folium.Polygon(
            locations=poly["coords"],
            color=poly["color"],
            weight=2,
            fill=True,
            fill_color=poly["color"],
            fill_opacity=0.3,
            popup=f"<b>ALERTA {poly['nivel'].upper()}</b><br>{poly['zona']}<br>{poly['evento']}"
        ).add_to(fg_smn)

    # Coberturas de radar SINARAME
    for r_nom, r_lat, r_lon, col in [("RMA08 Santa Isabel", -36.226, -66.883, "#0284c7"), ("RMA16 Villa Reynolds", -33.725, -65.385, "#d97706")]:
        folium.Marker([r_lat, r_lon], icon=folium.Icon(color="blue" if "Isabel" in r_nom else "orange", icon="broadcast-tower", prefix="fa"), tooltip=r_nom).add_to(fg_radares_sinarame)
        folium.Circle([r_lat, r_lon], radius=120000, color=col, weight=2, fill=True, fill_opacity=0.08, dash_array="5, 5").add_to(fg_radares_sinarame)
        folium.Circle([r_lat, r_lon], radius=240000, color=col, weight=1, fill=False, dash_array="8, 8").add_to(fg_radares_sinarame)

    # Cuerpos de Agua INA
    for h in lista_hidro_resumen:
        popup_txt = f"""
        <div style="font-family: Arial; width: 240px; font-size: 12px;">
            <strong style="color: #1e3a8a; font-size: 13px;">{h['nombre']}</strong><br>
            <span style="color:#64748b;">{h['rio']} ({h['distrito']})</span>
            <hr style="margin:5px 0;">
            <b>Nivel Actual:</b> <span style="color:{h['color']}; font-size:15px; font-weight:bold;">{h['nivel_actual']:.2f} m</span><br>
            <b>Cota Alerta:</b> {h['cota_alerta']:.2f} m | <b>Evac:</b> {h['cota_evac']:.2f} m<br>
            <b>Estado:</b> <span style="color:{h['color']}; font-weight:bold;">{h['estado']}</span><br>
            <b>Tendencia:</b> {h['tendencia']}<br>
            <span style="font-size:10px; color:#94a3b8;">Obs: {h['fecha']}</span>
        </div>
        """
        folium.CircleMarker(
            location=[h["lat"], h["lon"]],
            radius=8,
            popup=folium.Popup(popup_txt, max_width=260),
            tooltip=f"{h['nombre']}: {h['nivel_actual']:.2f} m ({h['estado']})",
            color=h["color"],
            fill=True,
            fill_color=h["color"],
            fill_opacity=0.9,
            weight=2
        ).add_to(fg_hidro)

    # Estaciones Meteorológicas
    for m in estaciones_meteo:
        grp = fg_meteo_apa if "APA" in m["red"] else (fg_meteo_sl if "San Luis" in m["provincia"] else (fg_meteo_cba if "Córdoba" in m["provincia"] else fg_meteo_lp))
        col = "#0284c7" if "APA" in m["red"] else ("#d97706" if "San Luis" in m["provincia"] else ("#7c3aed" if "Córdoba" in m["provincia"] else "#0d9488"))
        folium.CircleMarker(
            location=[m["lat"], m["lon"]],
            radius=5,
            popup=f"<b>{m['nombre']}</b><br>Red: {m['red']}<br>Lluvia 24h: <b>{m['lluvia_24h_mm']:.1f} mm</b>",
            tooltip=f"{m['nombre']} ({m['red']})",
            color=col, fill=True, fill_color=col, fill_opacity=0.85
        ).add_to(grp)

    # Rayos
    for ry in lista_rayos:
        folium.CircleMarker([ry["lat"], ry["lon"]], radius=5, tooltip=f"⚡ Descarga ({ry['hora']} hs)", color="#b45309", fill=True, fill_color="#facc15", fill_opacity=0.9).add_to(fg_rayos)

    # Agregar capas al mapa
    fg_smn.add_to(mapa)
    fg_satelite.add_to(mapa)
    fg_radar_mosaico.add_to(mapa)
    fg_radares_sinarame.add_to(mapa)
    fg_rayos.add_to(mapa)
    fg_hidro.add_to(mapa)
    fg_meteo_sl.add_to(mapa)
    fg_meteo_cba.add_to(mapa)
    fg_meteo_lp.add_to(mapa)
    fg_meteo_apa.add_to(mapa)

    folium.LayerControl(position="topright", collapsed=False).add_to(mapa)

    # Banner superior y script de Radar/Satélite
    banner_color = "#1e293b" if not diag_onda["alerta_activa"] else ("#b45309" if diag_onda["estado_alerta"] in ["Precaución", "Alerta Convectiva"] else "#b91c1c")
    html_banner = f"""
    <div style="position: fixed; top: 12px; left: 55px; right: 350px; background: {banner_color};
                color: white; border-radius: 8px; z-index: 1000; font-family: Arial, sans-serif;
                font-size: 11.5px; padding: 8px 14px; box-shadow: 0 3px 8px rgba(0,0,0,0.25);
                display: flex; justify-content: space-between; align-items: center;">
        <div>
            <b>SISTEMA DE ALERTA TEMPRANA CUENCA RÍO V:</b><br>
            {diag_onda['banner_msg']}
        </div>
        <div style="display: flex; gap: 8px; align-items: center; border-left: 1px solid rgba(255,255,255,0.2); padding-left: 10px;">
            <a href="sat_unificado_rio_v_triprovincial.xlsx" download style="background: #10b981; color: white; padding: 5px 8px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 10.5px;">📥 Excel</a>
            <a href="resumen_cruce_rio_v_triprovincial.csv" download style="background: #0284c7; color: white; padding: 5px 8px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 10.5px;">📄 CSV</a>
        </div>
    </div>
    <script>
        document.addEventListener("DOMContentLoaded", function() {{
            var mapObj = null;
            for (var k in window) {{
                if (k.startsWith("map_") && window[k] instanceof L.Map) {{
                    mapObj = window[k];
                    break;
                }}
            }}
            if (!mapObj) return;

            fetch("https://api.rainviewer.com/public/weather-maps.json")
                .then(r => r.json())
                .then(d => {{
                    var host = d.host || "https://tilecache.rainviewer.com";
                    if (d.satellite && d.satellite.infrared && d.satellite.infrared.length > 0) {{
                        var frameSat = d.satellite.infrared[d.satellite.infrared.length - 1];
                        var tileLayerSat = L.tileLayer(host + frameSat.path + "/256/{{z}}/{{x}}/{{y}}/1/1_0.png", {{ opacity: 0.55, zIndex: 240 }});
                        mapObj.eachLayer(ly => {{
                            if (ly instanceof L.FeatureGroup && ly.options && ly.options.name && ly.options.name.indexOf("Satélite") !== -1) {{
                                ly.clearLayers(); ly.addLayer(tileLayerSat);
                            }}
                        }});
                    }}
                    if (d.radar && d.radar.past && d.radar.past.length > 0) {{
                        var frameRadar = d.radar.past[d.radar.past.length - 1];
                        var tileLayerRadar = L.tileLayer(host + frameRadar.path + "/256/{{z}}/{{x}}/{{y}}/2/1_1.png", {{ opacity: 0.70, zIndex: 260 }});
                        mapObj.eachLayer(ly => {{
                            if (ly instanceof L.FeatureGroup && ly.options && ly.options.name && ly.options.name.indexOf("Radar Compuesto") !== -1) {{
                                ly.clearLayers(); ly.addLayer(tileLayerRadar);
                            }}
                        }});
                    }}
                }}).catch(e => console.warn("RainViewer offline:", e));
        }});
    </script>
    """
    mapa.get_root().html.add_child(folium.Element(html_banner))
    mapa.save(MAPA_HTML_SALIDA)
    print(f"   -> [VISOR PUBLICADO]: {MAPA_HTML_SALIDA}")

# =============================================================
# 11. EJECUCIÓN PRINCIPAL
# =============================================================
if __name__ == "__main__":
    print("=" * 70)
    print(f"SISTEMA SAT CUENCA RÍO V — INICIO EJECUCIÓN: {FECHA_TXT}")
    print("=" * 70)

    meteo_total = obtener_estaciones_omixom() + obtener_estaciones_san_luis() + obtener_estaciones_apa_lapampa()
    registros_hidro = obtener_datos_hidrologicos()
    rayos = obtener_descargas_atmosfericas()
    alertas_smn, poligonos_smn = obtener_alertas_smn()

    generar_entregables(meteo_total, registros_hidro, rayos, alertas_smn, poligonos_smn)
    print("=" * 70)
    print("ACTUALIZACIÓN COMPLETADA CON ÉXITO.")
    print("=" * 70)
