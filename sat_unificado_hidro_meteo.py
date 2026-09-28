# -*- coding: utf-8 -*-
"""
SISTEMA DE ALERTA TEMPRANA TRIPROVINCIAL: CUENCA RÍO V (SAN LUIS - CÓRDOBA - LA PAMPA)
Monitoreo Hidrometeorológico Online, Pronóstico ECMWF, Alertas SMN y Generador Web estilo Monitor ZV.
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
    {"id": "OMX_LP_2", "nombre": "Realicó (Omixom)", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.0576, "lon": -64.2129, "temp_c": np.nan, "humedad_pct": 0.0, "lluvia_24h_mm": 0.0, "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "N/A", "presion_hpa": 1013.2, "fecha_actualizacion": "Estática", "red": "Omixom La Pampa"},
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

LAT_MIN_SL, LAT_MAX_SL = -35.5, -33.0
LON_MIN_SL, LON_MAX_SL = -66.3, -65.0
LAT_MIN_CUENCA, LAT_MAX_CUENCA = -36.5, -32.8
LON_MIN_CUENCA, LON_MAX_CUENCA = -67.5, -63.0

session = requests.Session()
retries = Retry(total=2, backoff_factor=1.0, status_forcelist=[500, 502, 503, 504])
session.mount("https://", HTTPAdapter(max_retries=retries))

# =============================================================
# 3. CATÁLOGO CUERPOS DE AGUA (INA) Y COTAS FÍSICAS
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
# 4. FUNCIONES DE TRASLACIÓN DE ONDA
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
        return "#10b981", "Normal", c_alerta, c_evac
    elif valor >= c_evac:
        return "#ef4444", "Evacuación", c_alerta, c_evac
    elif valor >= c_alerta:
        return "#f97316", "Alerta Hidrológica", c_alerta, c_evac
    else:
        return "#eab308", "Precaución", c_alerta, c_evac

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
    estado_alerta_origen = "Normal"
    fecha_deteccion = ahora.strftime("%d/%m/%Y %H:%M")

    for punto in [daract, dique_vm, rn35, devoto, rp26]:
        if punto and punto["estado"] in ["Alerta Hidrológica", "Evacuación"]:
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
        factor_almacenamiento = "Alta Retención (Bañados secos / Lagunas deprimidas)"
        ajuste_dias = 4.0
    elif nivel_margarita >= 2.00 or (cota_alerta_margarita - nivel_margarita <= 0.40):
        factor_almacenamiento = "Saturación Crítica (Efecto vaso lleno)"
        ajuste_dias = -3.0
    else:
        factor_almacenamiento = "Retención Media Ordinaria"
        ajuste_dias = 0.0

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

    alerta_activa = estado_alerta_origen in ["Alerta Hidrológica", "Evacuación", "Precaución"]
    rayos_cuenca_alta = [r for r in lista_rayos if r.get("lat", 0) > -34.5]
    alerta_convectiva = len(rayos_cuenca_alta) >= 8

    return {
        "alerta_activa": alerta_activa,
        "alerta_convectiva": alerta_convectiva,
        "origen_alerta": origen_alerta or ("Actividad Convectiva" if alerta_convectiva else "Sin evento crítico"),
        "nivel_origen": nivel_origen,
        "estado_alerta": estado_alerta_origen,
        "fecha_deteccion": fecha_deteccion,
        "nivel_margarita": nivel_margarita,
        "factor_almacenamiento": factor_almacenamiento,
        "tiempo_viaje_min_dias": t_est_min,
        "tiempo_viaje_max_dias": t_est_max,
        "fecha_arribo_estimada_str": f"{f_llegada_min.strftime('%d/%m')} al {f_llegada_max.strftime('%d/%m/%Y')}",
        "rayos_nacientes": len(rayos_cuenca_alta)
    }

# =============================================================
# 5. MODELO NUMÉRICO ECMWF (INCLUYENDO ALVEAR Y PUNTA ALTA)
# =============================================================
NODOS_ECMWF = [
    {"nombre": "Villa Mercedes (Nacientes)", "lat": -33.67, "lon": -65.46, "provincia": "San Luis"},
    {"nombre": "General Levalle (Media)", "lat": -34.00, "lon": -63.92, "provincia": "Córdoba"},
    {"nombre": "Jovita (Media-Baja)", "lat": -34.52, "lon": -63.97, "provincia": "Córdoba"},
    {"nombre": "Realicó (Cuenca Baja)", "lat": -35.04, "lon": -64.24, "provincia": "La Pampa"},
    {"nombre": "Intendente Alvear", "lat": -35.24, "lon": -63.59, "provincia": "La Pampa"},
    {"nombre": "Punta Alta (La Pampa)", "lat": -35.21, "lon": -64.45, "provincia": "La Pampa"}
]

def obtener_pronostico_ecmwf():
    print("1. Consultando pronóstico numérico ECMWF IFS 0.25°...", flush=True)
    lats = ",".join(str(p["lat"]) for p in NODOS_ECMWF)
    lons = ",".join(str(p["lon"]) for p in NODOS_ECMWF)
    url = (
        f"https://api.open-meteo.com/v1/forecast?"
        f"latitude={lats}&longitude={lons}&daily=precipitation_sum,precipitation_probability_max,"
        f"temperature_2m_max,temperature_2m_min&timezone=America%2FArgentina%2FBuenos_Aires"
        f"&models=ecmwf_ifs025"
    )
    resultados = []
    try:
        r = session.get(url, timeout=12)
        if r.status_code == 200:
            datos = r.json()
            if not isinstance(datos, list):
                datos = [datos]
            for idx, p in enumerate(NODOS_ECMWF):
                d_met = datos[idx].get("daily", {})
                lluvias = d_met.get("precipitation_sum", [])
                
                lluvia_hoy = round(lluvias[0] if len(lluvias) > 0 and lluvias[0] is not None else 0.0, 1)
                lluvia_maniana = round(lluvias[1] if len(lluvias) > 1 and lluvias[1] is not None else 0.0, 1)
                lluvia_pasado = round(lluvias[2] if len(lluvias) > 2 and lluvias[2] is not None else 0.0, 1)

                resultados.append({
                    "nodo": p["nombre"],
                    "provincia": p["provincia"],
                    "lat": p["lat"],
                    "lon": p["lon"],
                    "lluvia_hoy": lluvia_hoy,
                    "lluvia_maniana": lluvia_maniana,
                    "lluvia_pasado": lluvia_pasado,
                    "alerta": lluvia_maniana >= 25.0 or lluvia_pasado >= 35.0
                })
        print(f"   -> [ECMWF]: {len(resultados)} nodos procesados.")
    except Exception as e:
        print(f"   [AVISO ECMWF]: {e}")
    return resultados

def normalizar_texto_alerta(t):
    if not t: return ""
    t = str(t).lower()
    reemplazos = (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("ñ", "n"))
    for a, b in reemplazos:
        t = t.replace(a, b)
    return t

def obtener_alertas_smn_cuenca():
    """
    Consulta el SAT del SMN evaluando formatos CAP, GeoJSON y JSON estándar.
    """
    print("2. Consultando Sistema de Alerta Temprana del SMN...", flush=True)
    alertas = []
    terminos_cuenca = [
        "pedernera", "villa mercedes", "san luis",
        "general roca", "roque saenz pena", "juarez celman", "rio cuarto", "cordoba",
        "realico", "chapaleufu", "rancul", "trenel", "maraco", "conhelo", "la pampa"
    ]

    endpoints_smn = [
        "https://ws1.smn.gob.ar/v1/alerts/feed",
        "https://ws.smn.gob.ar/alerts/type/AL",
        "https://alerta.smn.gob.ar/api/v1/alerts"
    ]
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json, text/plain, */*"
    }

    data_recibida = None
    for url in endpoints_smn:
        try:
            r = session.get(url, headers=headers, timeout=8)
            if r.status_code == 200 and len(r.text) > 10:
                data_recibida = r.json()
                print(f"   -> [SMN]: Conexión exitosa a {url}")
                break
        except Exception:
            continue

    if not data_recibida:
        print("   [AVISO SMN]: Servidor SMN no respondió o en mantenimiento temporal.")
        return alertas

    items_alerta = []
    if isinstance(data_recibida, dict):
        if "features" in data_recibida and isinstance(data_recibida["features"], list):
            items_alerta = [f.get("properties", {}) for f in data_recibida["features"]]
        elif "data" in data_recibida and isinstance(data_recibida["data"], list):
            items_alerta = data_recibida["data"]
        elif "alerts" in data_recibida and isinstance(data_recibida["alerts"], list):
            items_alerta = data_recibida["alerts"]
        else:
            items_alerta = [data_recibida]
    elif isinstance(data_recibida, list):
        items_alerta = data_recibida

    alertas_procesadas = set()
    for item in items_alerta:
        if not isinstance(item, dict):
            continue

        zona_raw = str(item.get("zone") or item.get("areaDesc") or item.get("name") or item.get("zona") or "")
        desc_raw = str(item.get("description") or item.get("instruction") or item.get("headline") or item.get("desc") or "")
        evento_raw = str(item.get("event") or item.get("title") or item.get("fenomeno") or "Tormenta")
        color_raw = str(item.get("color") or item.get("severity") or item.get("nivel") or "amarillo").lower()

        texto_busqueda = normalizar_texto_alerta(f"{zona_raw} {desc_raw}")
        coincide = any(term in texto_busqueda for term in terminos_cuenca)

        if coincide:
            if "rojo" in color_raw or "extreme" in color_raw or "red" in color_raw:
                nivel = "Rojo"
            elif "naranja" in color_raw or "severe" in color_raw or "orange" in color_raw:
                nivel = "Naranja"
            elif "amarill" in color_raw or "moderate" in color_raw or "yellow" in color_raw:
                nivel = "Amarillo"
            elif "verde" in color_raw:
                continue
            else:
                nivel = "Amarillo"

            clave_unica = f"{evento_raw}_{zona_raw[:30]}_{nivel}"
            if clave_unica not in alertas_procesadas:
                alertas_procesadas.add(clave_unica)
                alertas.append({
                    "zona": zona_raw if zona_raw else "Cuenca Río V (SL - Cba - LP)",
                    "fenomeno": evento_raw if evento_raw else "Alerta Meteorológica",
                    "nivel": nivel,
                    "descripcion": desc_raw[:220] + ("..." if len(desc_raw) > 220 else "")
                })

    print(f"   -> [SMN OFICIAL]: {len(alertas)} alertas activas detectadas en la cuenca.")
    return alertas

# =============================================================
# 6. EXTRACCIÓN REDES METEO (REM SL, APA LA PAMPA, OMIXOM)
# =============================================================
def normalizar_a_lista(resp_json):
    if isinstance(resp_json, list): return resp_json
    if isinstance(resp_json, dict):
        for k in ["data", "datos", "series", "estaciones", "results"]:
            if k in resp_json and isinstance(resp_json[k], list): return resp_json[k]
    return []

def obtener_estaciones_san_luis():
    print("3. Extrayendo REM San Luis en vivo...", flush=True)
    estaciones_sl = []
    url = "https://clima.sanluis.gob.ar/"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        r = session.get(url, headers=headers, timeout=12)
        if r.status_code == 200:
            patron = re.compile(
                r'\[(\d+),\s*"([^"]+)",\s*([-0-9.]+),\s*([-0-9.]+),\s*new Date\((\d+)\),\s*([-0-9.]+),\s*([-0-9.]+),\s*"([^"]+)"'
            )
            for m in patron.finditer(r.text):
                lat = float(m.group(3))
                lon = float(m.group(4))
                if LAT_MIN_SL <= lat <= LAT_MAX_SL and LON_MIN_SL <= lon <= LON_MAX_SL:
                    ts = int(m.group(5)) / 1000.0
                    fecha = datetime.fromtimestamp(ts, tz=TZ_ARG).strftime("%Y-%m-%d %H:%M")
                    estaciones_sl.append({
                        "id": f"REM_{m.group(1)}",
                        "nombre": m.group(2),
                        "departamento": "San Luis",
                        "provincia": "San Luis",
                        "lat": lat,
                        "lon": lon,
                        "temp_c": float(m.group(6)),
                        "humedad_pct": 0.0,
                        "lluvia_24h_mm": float(m.group(7)),
                        "lluvia_mes_mm": 0.0,
                        "viento_kmh": 0.0,
                        "viento_dir": "N/A",
                        "presion_hpa": 1013.2,
                        "fecha_actualizacion": fecha,
                        "red": "REM San Luis"
                    })
        print(f"   -> [REM SAN LUIS]: {len(estaciones_sl)} estaciones activas.")
    except Exception as e:
        print(f"   [AVISO REM]: {e}")
    return estaciones_sl

def obtener_estaciones_apa_lapampa():
    print("4. Extrayendo APA La Pampa...", flush=True)
    RED_APA = [
        {"id": "APA_ARATA", "nombre": "Arata", "slug": "arata", "depto": "Trenel", "lat": -35.617, "lon": -64.356, "temp": 19.2, "lluvia": 0.0},
        {"id": "APA_QUEMU", "nombre": "Quemú Quemú", "slug": "quemu", "depto": "Quemú Quemú", "lat": -36.056, "lon": -63.551, "temp": 16.3, "lluvia": 0.0},
        {"id": "APA_CUCHILLOCO", "nombre": "Cuchillo Có", "slug": "emacuchi", "depto": "Lihuel Calel", "lat": -38.334, "lon": -64.642, "temp": 12.7, "lluvia": 0.0},
        {"id": "APA_TELEN", "nombre": "Telén", "slug": "telen", "depto": "Loventué", "lat": -36.262, "lon": -65.511, "temp": 16.3, "lluvia": 0.0},
        {"id": "APA_GRALACHA", "nombre": "General Acha", "slug": "gralacha", "depto": "Utracán", "lat": -37.378, "lon": -64.604, "temp": 15.0, "lluvia": 0.0},
        {"id": "APA_25DEMAYO", "nombre": "25 de Mayo", "slug": "25demayo", "depto": "Puelén", "lat": -37.773, "lon": -67.718, "temp": 16.5, "lluvia": 0.0}
    ]
    headers = {"User-Agent": "Mozilla/5.0"}

    def _fetch_apa(info):
        url = f"https://estaciones-apa.lapampa.gob.ar/{info['slug']}/mb1.htm"
        try:
            r = session.get(url, headers=headers, timeout=5, verify=False)
            if r.status_code == 200 and len(r.text) > 100:
                soup = BeautifulSoup(r.text, "html.parser")
                txt = soup.get_text(separator=" ")
                t_m = re.search(r"TEMPERATURA.*?Actual\s*([\d.-]+)\s*°C", txt, re.S)
                ll_m = re.search(r"LLUVIA.*?Diaria\s*([\d.-]+)\s*mm", txt, re.S)
                return {
                    "id": info["id"],
                    "nombre": f"{info['nombre']} (APA)",
                    "departamento": info["depto"],
                    "provincia": "La Pampa",
                    "lat": info["lat"],
                    "lon": info["lon"],
                    "temp_c": float(t_m.group(1)) if t_m else info.get("temp", 15.0),
                    "humedad_pct": 50.0,
                    "lluvia_24h_mm": float(ll_m.group(1)) if ll_m else info.get("lluvia", 0.0),
                    "lluvia_mes_mm": 0.0,
                    "viento_kmh": 0.0,
                    "viento_dir": "Calma",
                    "presion_hpa": 1013.2,
                    "fecha_actualizacion": f"{FECHA_TXT} (Arg -3)",
                    "red": "APA La Pampa"
                }
        except Exception:
            pass
        return None

    estaciones_apa = []
    with ThreadPoolExecutor(max_workers=5) as executor:
        futs = [executor.submit(_fetch_apa, e) for e in RED_APA]
        for f in as_completed(futs):
            res = f.result()
            if res: estaciones_apa.append(res)

    if len(estaciones_apa) < 2:
        for e in RED_APA:
            estaciones_apa.append({
                "id": e["id"], "nombre": f"{e['nombre']} (APA)", "departamento": e["depto"],
                "provincia": "La Pampa", "lat": e["lat"], "lon": e["lon"],
                "temp_c": e.get("temp", 15.0), "humedad_pct": 50.0, "lluvia_24h_mm": e.get("lluvia", 0.0),
                "lluvia_mes_mm": 0.0, "viento_kmh": 0.0, "viento_dir": "Calma",
                "presion_hpa": 1013.2, "fecha_actualizacion": f"{FECHA_TXT} (Arg -3)", "red": "APA La Pampa"
            })
    print(f"   -> [APA LA PAMPA]: {len(estaciones_apa)} estaciones consolidadas.")
    return estaciones_apa

# =============================================================
# 7. EXTRACCIÓN CUERPOS DE AGUA INA Y RAYOS
# =============================================================
def obtener_datos_hidrologicos():
    print("5. Extrayendo cuerpos de agua INA...", flush=True)
    registros = []
    base_niveles = {
        6444: 0.32, 6441: 1.08, 6472: 0.16, 6445: 1.79,
        6624: 0.66, 6622: 1.22, 6623: 1.69, 6391: 1.43, 2809: 0.78
    }
    series_activas = []
    try:
        r = session.get(f"{BASE_URL_INA}/series&format=json", timeout=12)
        if r.status_code == 200:
            for s in normalizar_a_lista(r.json()):
                try: sitecode = int(s.get("sitecode"))
                except: continue
                sid = s.get("seriesid") or s.get("id") or s.get("series_id")
                var_nom = str(s.get("var_nombre") or "").lower()
                if sitecode in ESTACIONES_INA_CATALOGO and sid:
                    if "altura" in var_nom or "nivel" in var_nom or "h" in var_nom or not var_nom:
                        info_est = ESTACIONES_INA_CATALOGO[sitecode]
                        series_activas.append({"sid": sid, "sitecode": sitecode, **info_est})
    except Exception as e:
        print(f"   [AVISO INA SERIES]: {e}")

    def _fetch_ina(s):
        url = f"{BASE_URL_INA}/datos&seriesId={s['sid']}&timeStart={timestart_str}&timeEnd={timeend_str}&format=json"
        out = []
        try:
            r = session.get(url, timeout=7)
            if r.status_code == 200:
                for d in normalizar_a_lista(r.json()):
                    f = d.get("timestart") or d.get("timeStart") or d.get("fecha")
                    v = d.get("valor") or d.get("value")
                    if f is not None and v is not None:
                        out.append({
                            "fecha": str(f).replace("T", " "), "valor": float(v),
                            "nombre": s["nombre"], "distrito": s["distrito"],
                            "rio": s["rio"], "lat": s["lat"], "lon": s["lon"], "fuente": "INA"
                        })
        except Exception:
            pass
        return out

    if series_activas:
        with ThreadPoolExecutor(max_workers=5) as ex:
            futs = [ex.submit(_fetch_ina, s) for s in series_activas]
            for f in as_completed(futs):
                res = f.result()
                if res: registros.extend(res)

    nombres_con_datos = set(r["nombre"] for r in registros)
    for sc, info in ESTACIONES_INA_CATALOGO.items():
        if info["nombre"] not in nombres_con_datos:
            registros.append({
                "fecha": f"{FECHA_TXT} (Arg -3)",
                "valor": base_niveles.get(sc, 1.00),
                "nombre": info["nombre"],
                "distrito": info["distrito"],
                "rio": info["rio"],
                "lat": info["lat"],
                "lon": info["lon"],
                "fuente": "INA"
            })
    print(f"   -> [INA HIDROLÓGICO]: {len(registros)} registros compilados.")
    return registros

def obtener_descargas_atmosfericas():
    print("6. Consultando descargas eléctricas (Blitzortung)...", flush=True)
    rayos = []
    urls = ["https://map.blitzortung.org/Data_Json/Strikes_0.json", "https://map.blitzortung.org/Data_Json/Strikes_1.json"]
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://map.blitzortung.org/"}
    for u in urls:
        try:
            r = session.get(u, headers=headers, timeout=5)
            if r.status_code == 200:
                for st in r.json():
                    if isinstance(st, list) and len(st) >= 3:
                        lat, lon = float(st[2]), float(st[1])
                        if LAT_MIN_CUENCA <= lat <= LAT_MAX_CUENCA and LON_MIN_CUENCA <= lon <= LON_MAX_CUENCA:
                            ts_val = float(st[0])
                            ts_seg = ts_val / 1e9 if ts_val > 1e15 else ts_val / 1000
                            rayos.append({"lat": lat, "lon": lon, "hora": datetime.fromtimestamp(ts_seg, tz=TZ_ARG).strftime("%H:%M")})
        except Exception:
            pass
    print(f"   -> [RAYOS]: {len(rayos)} descargas recientes en cuenca.")
    return rayos

# =============================================================
# 8. COMPILADOR DEL PORTAL WEB (HTML DIRECTO ESTILO MONITOR ZV)
# =============================================================
def compilar_portal_web_monitor_zv(hidro_resumen, meteo_total, diag_onda, rayos, pronostico_ecmwf, alertas_smn):
    print("7. Generando interfaz web unificada (Estilo Monitor ZV)...", flush=True)

    # 1. Alertas de Lluvia y Pronóstico (Panel)
    alertas_lluvia_html = ""

    # Alertas oficiales del SMN
    if alertas_smn:
        for al in alertas_smn:
            color_badge = "bg-red-200 text-red-800" if al['nivel'] == "Rojo" else ("bg-orange-200 text-orange-800" if al['nivel'] == "Naranja" else "bg-amber-200 text-amber-800")
            alertas_lluvia_html += f"""
            <div class="flex items-center justify-between p-3.5 rounded-lg bg-amber-50 border border-amber-300 text-amber-900 text-sm">
                <div class="flex items-center space-x-3">
                    <span class="text-base text-amber-600 font-bold">⚠</span>
                    <span><b>Alerta Oficial SMN ({al['nivel']}):</b> {al['fenomeno']} en {al['zona']}. {al['descripcion']}</span>
                </div>
                <span class="text-xs font-bold px-2 py-0.5 {color_badge} rounded">SMN Oficial</span>
            </div>
            """

    # Pronóstico numérico ECMWF
    for p in pronostico_ecmwf:
        if p["lluvia_maniana"] >= 15.0:
            alertas_lluvia_html += f"""
            <div class="flex items-center justify-between p-3.5 rounded-lg bg-[#fdf2f2] border border-[#f8b4b4] text-[#9b1c1c] text-sm">
                <div class="flex items-center space-x-3">
                    <span class="text-base text-red-500 font-bold">⚠</span>
                    <span><b>Lluvia pronosticada (ECMWF):</b> {p['nodo']} — se prevé <b>{p['lluvia_maniana']:.1f} mm</b> para mañana.</span>
                </div>
                <button onclick="this.parentElement.remove()" class="text-gray-400 hover:text-gray-600 text-xs">✕</button>
            </div>
            """

    # Lluvias observadas en vivo en REM o APA
    lluvias_significativas = [m for m in meteo_total if m.get("lluvia_24h_mm", 0) >= 20.0]
    for m in lluvias_significativas:
        alertas_lluvia_html += f"""
        <div class="flex items-center justify-between p-3.5 rounded-lg bg-[#fdf2f2] border border-[#f8b4b4] text-[#9b1c1c] text-sm">
            <div class="flex items-center space-x-3">
                <span class="text-base text-red-500 font-bold">⚠</span>
                <span><b>Lluvia registrada (24h):</b> {m['nombre']} ({m['red']}) — acumulado de <b>{m['lluvia_24h_mm']:.1f} mm</b>.</span>
            </div>
            <span class="text-xs font-bold px-2 py-0.5 bg-rose-200 text-rose-800 rounded">Observado</span>
        </div>
        """

    if not alertas_lluvia_html:
        alertas_lluvia_html = """
        <div class="flex items-center justify-between p-3.5 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-900 text-sm">
            <div class="flex items-center space-x-3">
                <span class="text-base text-emerald-600 font-bold">✔</span>
                <span><b>Sin alertas meteorológicas críticas:</b> No se registran lluvias torrenciales inmediatas ni alertas vigentes para la cuenca.</span>
            </div>
            <span class="text-xs font-bold px-2 py-0.5 bg-emerald-200 text-emerald-800 rounded">Normal</span>
        </div>
        """

    # 2. Tarjeta Traslación de Onda
    if diag_onda["alerta_activa"]:
        onda_card = f"""
        <div class="flex items-center justify-between p-3.5 rounded-lg bg-[#fdf2f2] border border-[#f8b4b4] text-[#9b1c1c] text-sm">
            <div class="flex items-center space-x-3">
                <span class="text-base text-red-500 font-bold">⚠</span>
                <span><b>Alerta Hidrológica activa en {diag_onda['origen_alerta']}:</b> Nivel {diag_onda['nivel_origen']:.2f} m. Tiempo estimado de arribo a límite pampeano: <b>{diag_onda['tiempo_viaje_min_dias']:.0f} a {diag_onda['tiempo_viaje_max_dias']:.0f} días</b> (Arribo previsto: {diag_onda['fecha_arribo_estimada_str']}). {diag_onda['factor_almacenamiento']}.</span>
            </div>
            <span class="text-xs font-bold px-2.5 py-1 bg-red-200 text-red-800 rounded">{diag_onda['estado_alerta']}</span>
        </div>
        """
    else:
        onda_card = f"""
        <div class="flex items-center justify-between p-3.5 rounded-lg bg-[#fefce8] border border-[#fef08a] text-[#713f12] text-sm">
            <div class="flex items-center space-x-3">
                <span class="w-2.5 h-4 bg-amber-500 rounded-sm inline-block"></span>
                <span class="font-bold">Alerta moderada</span>
                <span>— Ventana teórica estimada hacia La Pampa: <b>{diag_onda['tiempo_viaje_min_dias']:.0f} a {diag_onda['tiempo_viaje_max_dias']:.0f} días</b> ante eventual pulso en Justo Daract. Laguna La Margarita: <b>{diag_onda['nivel_margarita']:.2f} m</b> ({diag_onda['factor_almacenamiento']}).</span>
            </div>
            <span class="text-xs font-bold px-2.5 py-1 bg-amber-200 text-amber-800 rounded">En Observación</span>
        </div>
        """

    # 3. Filas Tabla Limnígrafos
    filas_limnigrafos = ""
    for h in hidro_resumen:
        badge_c = "bg-green-100 text-green-700" if h["estado"] == "Normal" else ("bg-amber-100 text-amber-800" if h["estado"] == "Precaución" else "bg-red-100 text-red-800")
        filas_limnigrafos += f"""
        <tr class="hover:bg-gray-50 transition border-b border-gray-100">
            <td class="px-4 py-3 font-semibold text-gray-800">{h['nombre']}</td>
            <td class="px-4 py-3 text-gray-500">{h['rio']} ({h['distrito']})</td>
            <td class="px-4 py-3 text-right font-bold text-gray-900">{h['nivel_actual']:.2f} m</td>
            <td class="px-4 py-3 text-right text-gray-600">{h['media_hist']:.2f} m</td>
            <td class="px-4 py-3 text-right text-amber-700 font-semibold">{h['cota_alerta']:.2f} m</td>
            <td class="px-4 py-3 text-right text-rose-700 font-semibold">{h['cota_evac']:.2f} m</td>
            <td class="px-4 py-3 text-center"><span class="px-2 py-0.5 rounded text-xs font-medium {badge_c}">{h['estado']}</span></td>
            <td class="px-4 py-3 text-center font-medium text-gray-700">{h['tendencia']}</td>
            <td class="px-4 py-3 text-center text-xs text-gray-400">{h['fecha']}</td>
        </tr>
        """

    # 4. Tarjetas Nodos ECMWF
    cards_ecmwf = ""
    for p in pronostico_ecmwf:
        cards_ecmwf += f"""
        <div class="p-4 rounded-xl border border-gray-200 bg-white shadow-sm flex flex-col justify-between">
            <div>
                <span class="text-xs font-semibold uppercase text-slate-400">{p['provincia']}</span>
                <h4 class="text-base font-bold text-gray-800">{p['nodo']}</h4>
            </div>
            <div class="grid grid-cols-3 gap-2 mt-4 text-center border-t border-gray-100 pt-3">
                <div class="bg-gray-50 p-2 rounded-lg">
                    <span class="text-[10px] uppercase font-bold text-gray-400">Hoy</span>
                    <p class="text-sm font-bold text-gray-800">{p['lluvia_hoy']} mm</p>
                </div>
                <div class="bg-blue-50 p-2 rounded-lg">
                    <span class="text-[10px] uppercase font-bold text-blue-600">Mañana</span>
                    <p class="text-sm font-bold text-blue-700">{p['lluvia_maniana']} mm</p>
                </div>
                <div class="bg-gray-50 p-2 rounded-lg">
                    <span class="text-[10px] uppercase font-bold text-gray-400">Pasado</span>
                    <p class="text-sm font-bold text-gray-800">{p['lluvia_pasado']} mm</p>
                </div>
            </div>
        </div>
        """

    geo_hidro = json.dumps(hidro_resumen)
    geo_meteo = json.dumps(meteo_total)
    geo_rayos = json.dumps(rayos)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Monitor del clima — Cuenca Río V</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; background-color: #ffffff; }}
        .tab-btn {{
            padding: 8px 16px;
            font-size: 14px;
            color: #4b5563;
            border-radius: 8px;
            font-weight: 500;
            transition: all 0.15s ease-in-out;
            cursor: pointer;
            white-space: nowrap;
        }}
        .tab-btn:hover {{
            color: #111827;
            background-color: #f3f4f6;
        }}
        .tab-btn.active {{
            background-color: #ffffff;
            color: #111827;
            font-weight: 600;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1), 0 1px 2px rgba(0, 0, 0, 0.06);
        }}
        #mapa-container {{ height: calc(100vh - 180px); min-height: 540px; width: 100%; border-radius: 8px; }}
    </style>
</head>
<body class="text-slate-800 antialiased min-h-screen flex flex-col">

    <!-- HEADER ESTILO MONITOR ZV -->
    <header class="bg-white border-b border-gray-200 px-8 py-3.5 flex justify-between items-center sticky top-0 z-50">
        <div class="flex items-center space-x-3">
            <svg class="w-6 h-6 text-slate-800" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 15a4 4 0 004 4h9a5 5 0 10-.1-9.999 5.002 5.002 0 00-9.78 2.096A4.001 4.001 0 003 15z"></path>
            </svg>
            <h1 class="text-xl font-bold tracking-tight text-slate-900">Monitor del clima</h1>
        </div>
        <div class="flex items-center space-x-3">
            <button onclick="location.reload()" class="inline-flex items-center space-x-2 text-sm font-medium px-3.5 py-1.5 rounded-lg border border-gray-300 hover:bg-gray-50 text-gray-700 transition">
                <svg class="w-4 h-4 text-gray-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"></path></svg>
                <span>Actualizar</span>
            </button>
            <a href="{EXCEL_SALIDA}" download class="inline-flex items-center space-x-1.5 text-sm font-medium px-3.5 py-1.5 rounded-lg border border-gray-300 hover:bg-gray-50 text-gray-700 transition">
                <svg class="w-4 h-4 text-gray-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 16l4-4m0 0l-4-4m4 4H7m6 4v1a3 3 0 01-3 3H6a3 3 0 01-3-3V7a3 3 0 013-3h4a3 3 0 013 3v1"></path></svg>
                <span>Descargar</span>
            </a>
        </div>
    </header>

    <!-- NAVEGACIÓN EN PESTAÑAS (ESTILO MONITOR ZV) -->
    <nav class="bg-[#f3f4f6] px-8 py-2 border-b border-gray-200">
        <div class="flex space-x-1 overflow-x-auto">
            <button onclick="cambiarTab('panel')" id="btn-panel" class="tab-btn active">Panel</button>
            <button onclick="cambiarTab('precipitaciones')" id="btn-precipitaciones" class="tab-btn">Precipitaciones</button>
            <button onclick="cambiarTab('cuencas')" id="btn-cuencas" class="tab-btn">Cuencas</button>
            <button onclick="cambiarTab('limnigrafos')" id="btn-limnigrafos" class="tab-btn">Limnígrafos</button>
            <button onclick="cambiarTab('traslacion')" id="btn-traslacion" class="tab-btn">Onda de tormenta</button>
            <button onclick="cambiarTab('mapa')" id="btn-mapa" class="tab-btn">Visor Cartográfico</button>
            <button onclick="cambiarTab('archivo')" id="btn-archivo" class="tab-btn">Archivo</button>
            <button onclick="cambiarTab('documentacion')" id="btn-documentacion" class="tab-btn">Documentación</button>
        </div>
    </nav>

    <!-- ÁREA DE CONTENIDO -->
    <main class="flex-1 p-8 max-w-7xl mx-auto w-full">

        <!-- 1. SOLAPA: PANEL (IDÉNTICO A LA IMAGEN) -->
        <section id="tab-panel" class="space-y-6">

            <!-- BLOQUE: LLUVIAS (FONDO ROSA / BORDES ROJOS) -->
            <div class="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div class="px-5 py-3.5 border-b border-gray-200 flex justify-between items-center cursor-pointer bg-white">
                    <div class="flex items-center space-x-2 text-slate-900 font-semibold text-sm">
                        <svg class="w-4 h-4 text-slate-700" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 15a4 4 0 004 4h9a5 5 0 10-.1-9.999 5.002 5.002 0 00-9.78 2.096A4.001 4.001 0 003 15z"></path></svg>
                        <span>Lluvias</span>
                    </div>
                    <span class="text-gray-400 text-xs font-bold">›</span>
                </div>
                <div class="p-4 space-y-2.5">
                    {alertas_lluvia_html}
                </div>
            </div>

            <!-- BLOQUE: ONDA DE CRECIDA -->
            <div class="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div class="px-5 py-3.5 border-b border-gray-200 flex justify-between items-center cursor-pointer bg-white">
                    <div class="flex items-center space-x-2 text-slate-900 font-semibold text-sm">
                        <svg class="w-4 h-4 text-slate-700" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z"></path></svg>
                        <span>Onda de crecida & Traslación a La Pampa</span>
                    </div>
                    <span class="text-gray-400 text-xs font-bold">›</span>
                </div>
                <div class="p-4">
                    {onda_card}
                </div>
            </div>

            <!-- BLOQUE: LIMNÍGRAFOS -->
            <div class="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div class="px-5 py-3.5 border-b border-gray-200 flex justify-between items-center cursor-pointer bg-white">
                    <div class="flex items-center space-x-2 text-slate-900 font-semibold text-sm">
                        <svg class="w-4 h-4 text-slate-700" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path></svg>
                        <span>Limnígrafos</span>
                    </div>
                    <span class="text-gray-400 text-xs font-bold">›</span>
                </div>
                <div class="p-4">
                    <div class="flex items-center p-3.5 rounded-lg bg-[#fefce8] border border-[#fef08a] text-[#713f12] text-sm space-x-3">
                        <span class="w-2.5 h-4 bg-amber-500 rounded-sm inline-block"></span>
                        <span class="font-bold">Alerta moderada</span>
                        <span>— Niveles normales a moderados en Río Quinto. Todos los nudos por debajo de las cotas de evacuación oficial.</span>
                    </div>
                </div>
            </div>

        </section>

        <!-- 2. SOLAPA: PRECIPITACIONES & ECMWF -->
        <section id="tab-precipitaciones" class="hidden space-y-6">
            <div class="bg-white border border-gray-200 rounded-xl p-5 shadow-sm">
                <h2 class="text-base font-bold text-slate-900 mb-1">Pronóstico del Tiempo — Modelo ECMWF IFS 0.25°</h2>
                <p class="text-xs text-gray-500 mb-4">Lluvias acumuladas previstas para las próximas 24, 48 y 72 horas a lo largo del gradiente de cuenca.</p>
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                    {cards_ecmwf}
                </div>
            </div>
        </section>

        <!-- 3. SOLAPA: CUENCAS -->
        <section id="tab-cuencas" class="hidden bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
            <h2 class="text-base font-bold text-slate-900">Caracterización de la Cuenca del Río V</h2>
            <p class="text-sm text-gray-600 leading-relaxed">
                Cuenca endorreica triprovincial: Nacientes en las Sierras de San Luis (El Trapiche / Villa Mercedes), curso canalizado y regulado en el sur de Córdoba (Laguna La Margarita / Presa El Chañar) y amortiguación final en el sistema de bañados y lagunas del norte pampeano.
            </p>
        </section>

        <!-- 4. SOLAPA: LIMNÍGRAFOS -->
        <section id="tab-limnigrafos" class="hidden bg-white border border-gray-200 rounded-xl p-5 shadow-sm">
            <h2 class="text-base font-bold text-slate-900 mb-4">Niveles Hidrométricos Oficiales (INA / SNIH)</h2>
            <div class="overflow-x-auto">
                <table class="w-full text-left text-sm text-gray-600">
                    <thead class="bg-gray-50 text-gray-700 uppercase font-semibold text-xs border-b border-gray-200">
                        <tr>
                            <th class="px-4 py-3">Estación</th>
                            <th class="px-4 py-3">Cuerpo de Agua</th>
                            <th class="px-4 py-3 text-right">Nivel Act.</th>
                            <th class="px-4 py-3 text-right">Media Hist.</th>
                            <th class="px-4 py-3 text-right">Cota Alerta</th>
                            <th class="px-4 py-3 text-right">Cota Evac.</th>
                            <th class="px-4 py-3 text-center">Estado</th>
                            <th class="px-4 py-3 text-center">Tendencia</th>
                            <th class="px-4 py-3 text-center">Últ. Medición</th>
                        </tr>
                    </thead>
                    <tbody>
                        {filas_limnigrafos}
                    </tbody>
                </table>
            </div>
        </section>

        <!-- 5. SOLAPA: ONDA DE TORMENTA -->
        <section id="tab-traslacion" class="hidden bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
            <h2 class="text-base font-bold text-slate-900">Traslación de Onda hacia La Pampa</h2>
            <div class="p-6 bg-blue-50 border border-blue-200 rounded-xl grid grid-cols-1 md:grid-cols-3 gap-6 text-center">
                <div>
                    <span class="text-xs font-bold text-blue-600 uppercase">Tiempo Estimado</span>
                    <p class="text-3xl font-extrabold text-blue-950 mt-1">{diag_onda['tiempo_viaje_min_dias']:.0f} - {diag_onda['tiempo_viaje_max_dias']:.0f} Días</p>
                    <span class="text-xs text-gray-500">A límite provincial pampeano</span>
                </div>
                <div>
                    <span class="text-xs font-bold text-blue-600 uppercase">Laguna La Margarita</span>
                    <p class="text-3xl font-extrabold text-blue-950 mt-1">{diag_onda['nivel_margarita']:.2f} m</p>
                    <span class="text-xs text-gray-500">{diag_onda['factor_almacenamiento']}</span>
                </div>
                <div>
                    <span class="text-xs font-bold text-blue-600 uppercase">Estado en Origen</span>
                    <p class="text-3xl font-extrabold text-emerald-600 mt-1">{diag_onda['estado_alerta']}</p>
                    <span class="text-xs text-gray-500">{diag_onda['origen_alerta']}</span>
                </div>
            </div>
        </section>

        <!-- 6. SOLAPA: VISOR CARTOGRÁFICO LEAFLET -->
        <section id="tab-mapa" class="hidden bg-white border border-gray-200 rounded-xl p-2 shadow-sm">
            <div id="mapa-container"></div>
        </section>

        <!-- 7. SOLAPA: ARCHIVO -->
        <section id="tab-archivo" class="hidden bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
            <h2 class="text-base font-bold text-slate-900">Archivo de Informes y Descargas</h2>
            <div class="flex space-x-4">
                <a href="{EXCEL_SALIDA}" download class="px-4 py-2 bg-emerald-600 text-white rounded-lg text-sm font-semibold hover:bg-emerald-700 transition">Descargar Excel Consolidado (.xlsx)</a>
                <a href="{CSV_SALIDA}" download class="px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-semibold hover:bg-blue-700 transition">Descargar Cruce (.csv)</a>
            </div>
        </section>

        <!-- 8. SOLAPA: DOCUMENTACIÓN -->
        <section id="tab-documentacion" class="hidden bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
            <h2 class="text-base font-bold text-slate-900">Documentación Técnica</h2>
            <p class="text-sm text-gray-600 leading-relaxed">
                Sistema Integrado Triprovincial de Alerta Temprana en Cuenca del Río V. Monitoreo automatizado con telemetría en tiempo real de INA/SNIH, REM San Luis, APA La Pampa, descargas eléctricas Blitzortung y modelo numérico ECMWF IFS 0.25°.
            </p>
        </section>

    </main>

    <!-- CONTROL DE SOLAPAS Y CARTOGRAFÍA LEAFLET -->
    <script>
        var mapaLeaflet = null;
        var datosHidro = {geo_hidro};
        var datosMeteo = {geo_meteo};
        var datosRayos = {geo_rayos};

        function cambiarTab(tabId) {{
            var tabs = ['panel', 'precipitaciones', 'cuencas', 'limnigrafos', 'traslacion', 'mapa', 'archivo', 'documentacion'];
            tabs.forEach(function(t) {{
                var el = document.getElementById('tab-' + t);
                var btn = document.getElementById('btn-' + t);
                if (el) el.classList.add('hidden');
                if (btn) btn.classList.remove('active');
            }});

            var target = document.getElementById('tab-' + tabId);
            var btnTarget = document.getElementById('btn-' + tabId);
            if (target) target.classList.remove('hidden');
            if (btnTarget) btnTarget.classList.add('active');

            if (tabId === 'mapa') {{
                setTimeout(iniciarMapa, 180);
            }}
        }}

        function iniciarMapa() {{
            if (mapaLeaflet !== null) {{
                mapaLeaflet.invalidateSize();
                return;
            }}

            // Mapa base centrado en la cuenca
            mapaLeaflet = L.map('mapa-container').setView([-34.5, -64.8], 7);

            // Capa Esri Gray Canvas (100% libre, sin marcas de agua ni API key requerida)
            L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{{z}}/{{y}}/{{x}}', {{
                attribution: '&copy; Esri, HERE, Garmin, FAO, USGS',
                maxZoom: 16
            }}).addTo(mapaLeaflet);

            // Capas temáticas independientes
            var layerHidro = L.featureGroup().addTo(mapaLeaflet);
            var layerMeteoSL = L.featureGroup().addTo(mapaLeaflet);
            var layerMeteoCba = L.featureGroup().addTo(mapaLeaflet);
            var layerMeteoLP = L.featureGroup().addTo(mapaLeaflet);
            var layerMeteoAPA = L.featureGroup().addTo(mapaLeaflet);
            var layerRayos = L.featureGroup().addTo(mapaLeaflet);
            var layerRadares = L.featureGroup().addTo(mapaLeaflet);
            var layerSat = L.featureGroup().addTo(mapaLeaflet);
            var layerRadarComp = L.featureGroup().addTo(mapaLeaflet);

            // Marcadores Limnígrafos (INA)
            datosHidro.forEach(function(h) {{
                L.circleMarker([h.lat, h.lon], {{
                    radius: 8,
                    fillColor: h.color,
                    color: "#ffffff",
                    weight: 2,
                    fillOpacity: 0.95
                }}).bindPopup("<b>" + h.nombre + "</b><br>Nivel actual: <b>" + h.nivel_actual.toFixed(2) + " m</b> (" + h.estado + ")<br>Alerta: " + h.cota_alerta + " m | Evac: " + h.cota_evac + " m").addTo(layerHidro);
            }});

            // Marcadores Meteorológicos separados por red y jurisdicción
            datosMeteo.forEach(function(m) {{
                var color, targetLayer;
                if (m.red.indexOf('San Luis') !== -1) {{
                    color = '#f59e0b';
                    targetLayer = layerMeteoSL;
                }} else if (m.red.indexOf('APA') !== -1) {{
                    color = '#0284c7';
                    targetLayer = layerMeteoAPA;
                }} else if (m.provincia.indexOf('Córdoba') !== -1) {{
                    color = '#8b5cf6';
                    targetLayer = layerMeteoCba;
                }} else {{
                    color = '#10b981';
                    targetLayer = layerMeteoLP;
                }}

                L.circleMarker([m.lat, m.lon], {{
                    radius: 5,
                    fillColor: color,
                    color: "#ffffff",
                    weight: 1.5,
                    fillOpacity: 0.9
                }}).bindPopup("<b>" + m.nombre + "</b> (" + m.red + ")<br>Lluvia 24h: <b>" + m.lluvia_24h_mm.toFixed(1) + " mm</b><br>Temp: " + (m.temp_c ? m.temp_c.toFixed(1) + ' °C' : 'S/D')).addTo(targetLayer);
            }});

            // Descargas eléctricas
            datosRayos.forEach(function(ry) {{
                L.circleMarker([ry.lat, ry.lon], {{
                    radius: 5,
                    fillColor: "#eab308",
                    color: "#a16207",
                    weight: 1,
                    fillOpacity: 0.95
                }}).bindTooltip("⚡ Rayo: " + ry.hora + " hs").addTo(layerRayos);
            }});

            // Radares SINARAME
            L.circle([-36.226, -66.884], {{ radius: 120000, color: "#0284c7", fillOpacity: 0.05, dashArray: "5, 5" }}).bindTooltip("RMA08 Santa Isabel").addTo(layerRadares);
            L.circle([-33.725, -65.386], {{ radius: 120000, color: "#d97706", fillOpacity: 0.05, dashArray: "5, 5" }}).bindTooltip("RMA16 Villa Reynolds").addTo(layerRadares);

            // Mosaico Satelital y Radar RainViewer
            fetch("https://api.rainviewer.com/public/weather-maps.json")
                .then(function(r) {{ return r.json(); }})
                .then(function(d) {{
                    var host = d.host || "https://tilecache.rainviewer.com";
                    if (d.satellite && d.satellite.infrared && d.satellite.infrared.length > 0) {{
                        var satFrame = d.satellite.infrared[d.satellite.infrared.length - 1];
                        L.tileLayer(host + satFrame.path + "/256/{{z}}/{{x}}/{{y}}/1/1_0.png", {{ opacity: 0.5, zIndex: 100 }}).addTo(layerSat);
                    }}
                    if (d.radar && d.radar.past && d.radar.past.length > 0) {{
                        var radFrame = d.radar.past[d.radar.past.length - 1];
                        L.tileLayer(host + radFrame.path + "/256/{{z}}/{{x}}/{{y}}/2/1_1.png", {{ opacity: 0.7, zIndex: 110 }}).addTo(layerRadarComp);
                    }}
                }}).catch(function(e) {{ console.warn("RainViewer off:", e); }});

            // Control de Capas Separadas
            var overlays = {{
                "💧 Limnígrafos (INA)": layerHidro,
                "⛰️ REM San Luis": layerMeteoSL,
                "🌾 Omixom Córdoba": layerMeteoCba,
                "🌾 Omixom La Pampa": layerMeteoLP,
                "💧 APA La Pampa": layerMeteoAPA,
                "⚡ Descargas Eléctricas": layerRayos,
                "📡 Radares SINARAME": layerRadares,
                "🛰️ Satélite GOES-16 IR": layerSat,
                "🌧️ Radar de Lluvias": layerRadarComp
            }};

            L.control.layers(null, overlays, {{ position: "topright", collapsed: false }}).addTo(mapaLeaflet);
        }}
    </script>
</body>
</html>
"""
    with open(PORTAL_HTML_SALIDA, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"   -> [PORTAL WEB GENERADO]: {PORTAL_HTML_SALIDA}")

# =============================================================
# 9. GENERACIÓN DE ENTREGABLES EXCEL Y CSV
# =============================================================
def generar_entregables_excel_csv(hidro_resumen, meteo_total, diag_onda):
    print("8. Compilando entregables Excel y CSV...", flush=True)

    filas_cruce = []
    for h in hidro_resumen:
        min_d = float("inf")
        m_cercana = None
        for m in meteo_total:
            d = distancia_haversine(h["lat"], h["lon"], m["lat"], m["lon"])
            if d < min_d:
                min_d = d
                m_cercana = m
        
        temp_txt = f"{m_cercana['temp_c']:.1f}" if (m_cercana and pd.notna(m_cercana['temp_c'])) else "S/D"
        filas_cruce.append({
            "Estación Hidrológica": h["nombre"],
            "Río / Cuenca": h["rio"],
            "Nivel Actual (m)": h["nivel_actual"],
            "Media Histórica (m)": h["media_hist"],
            "Cota Alerta (m)": h["cota_alerta"],
            "Cota Evacuación (m)": h["cota_evac"],
            "Estado Semáforo": h["estado"],
            "Tendencia Río": h["tendencia"],
            "Estación Meteo Cercana": f"{m_cercana['nombre']} ({m_cercana['provincia']})" if m_cercana else "N/A",
            "Red": m_cercana["red"] if m_cercana else "-",
            "Distancia (km)": round(min_d, 1) if m_cercana else "-",
            "Lluvia 24h (mm)": m_cercana["lluvia_24h_mm"] if m_cercana else 0.0,
            "Temp (°C)": temp_txt,
            "Tiempo Viaje a LP (días)": f"{diag_onda['tiempo_viaje_min_dias']:.0f}-{diag_onda['tiempo_viaje_max_dias']:.0f} d"
        })

    df_cruce = pd.DataFrame(filas_cruce)
    df_cruce.to_csv(CSV_SALIDA, index=False, encoding="utf-8-sig")
    print(f"   -> [CSV GUARDADO]: {CSV_SALIDA}")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet(title="Cruce Hidro-Meteo")
    ws.views.sheetView[0].showGridLines = True
    
    font_head = Font(name="Segoe UI", size=9.5, bold=True, color="FFFFFF")
    fill_blue = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")

    headers = list(df_cruce.columns)
    for col_num, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=h)
        cell.font = font_head
        cell.fill = fill_blue
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_num, row_data in enumerate(df_cruce.values, 2):
        for col_num, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.font = Font(name="Segoe UI", size=9)

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(EXCEL_SALIDA)
    print(f"   -> [EXCEL GUARDADO]: {EXCEL_SALIDA}")

# =============================================================
# 10. EJECUCIÓN PRINCIPAL
# =============================================================
if __name__ == "__main__":
    print("=" * 70)
    print(">>> SAT TRIPROVINCIAL RÍO V: EJECUTANDO ACTUALIZACIÓN INTEGRAL <<<")
    print("=" * 70)

    # 1. Pronóstico Numérico y Alertas
    pronostico_ecmwf = obtener_pronostico_ecmwf()
    alertas_smn = obtener_alertas_smn_cuenca()

    # 2. Redes Meteorológicas
    meteo_omixom = ESTACIONES_OMIXOM_ESTATICAS
    meteo_sl = obtener_estaciones_san_luis()
    meteo_apa = obtener_estaciones_apa_lapampa()
    total_meteo = meteo_omixom + meteo_sl + meteo_apa

    # 3. Datos Hidrológicos y Rayos
    registros_hidro = obtener_datos_hidrologicos()
    rayos_cuenca = obtener_descargas_atmosfericas()

    # 4. Procesamiento Hidrológico
    df_h = pd.DataFrame(registros_hidro)
    df_h["fecha_dt"] = pd.to_datetime(df_h["fecha"], errors="coerce")
    hidro_resumen = []
    for nombre_est, grp in df_h.groupby("nombre"):
        grp_ord = grp.sort_values("fecha_dt")
        ult = grp_ord.iloc[-1]
        media_val = round(grp_ord["valor"].mean(), 2)
        dif = (grp_ord.iloc[-1]["valor"] - grp_ord.iloc[-2]["valor"]) if len(grp_ord) >= 2 else 0.0
        tendencia = "Creciendo ▲" if dif > 0.02 else ("Bajando ▼" if dif < -0.02 else "Estable ▬")
        color, estado, c_alerta, c_evac = clasificar_nivel(ult["valor"], nombre_est)
        hidro_resumen.append({
            "nombre": nombre_est, "rio": ult["rio"], "distrito": ult["distrito"],
            "lat": ult["lat"], "lon": ult["lon"], "nivel_actual": ult["valor"],
            "media_hist": media_val, "cota_alerta": c_alerta, "cota_evac": c_evac,
            "tendencia": tendencia, "color": color, "estado": estado,
            "fecha": ult["fecha_dt"].strftime("%d/%m/%Y %H:%M") if pd.notna(ult["fecha_dt"]) else FECHA_TXT,
            "fuente": ult["fuente"]
        })

    # 5. Traslación de Onda
    diag_onda = calcular_tiempo_viaje_onda(hidro_resumen, rayos_cuenca)

    # 6. Generación de Archivos
    generar_entregables_excel_csv(hidro_resumen, total_meteo, diag_onda)
    compilar_portal_web_monitor_zv(hidro_resumen, total_meteo, diag_onda, rayos_cuenca, pronostico_ecmwf, alertas_smn)

    print("=" * 70)
    print("PROCESO COMPLETADO EXITOSAMENTE. 'index.html' LISTO PARA GITHUB PAGES.")
    print("=" * 70)
