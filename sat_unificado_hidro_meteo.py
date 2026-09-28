# -*- coding: utf-8 -*-
"""
SISTEMA DE ALERTA TEMPRANA TRIPROVINCIAL: CUENCA RÍO V (SAN LUIS - CÓRDOBA - LA PAMPA)
Monitoreo Hidrometeorológico, Pronóstico ECMWF, Alertas SMN y Portal Web estilo Monitor ZV.
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
# 4. FUNCIONES DE CÁLCULO Y TRASLACIÓN DE ONDA
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
# 5. MODELO NUMÉRICO ECMWF Y ALERTAS SMN
# =============================================================
NODOS_ECMWF = [
    {"nombre": "Villa Mercedes (Nacientes)", "lat": -33.67, "lon": -65.46, "provincia": "San Luis"},
    {"nombre": "General Levalle (Media)", "lat": -34.00, "lon": -63.92, "provincia": "Córdoba"},
    {"nombre": "Jovita (Media-Baja)", "lat": -34.52, "lon": -63.97, "provincia": "Córdoba"},
    {"nombre": "Realicó (Cuenca Baja)", "lat": -35.04, "lon": -64.24, "provincia": "La Pampa"}
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
                fechas = d_met.get("time", [])
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

def obtener_alertas_smn_cuenca():
    print("2. Consultando Sistema de Alerta Temprana del SMN...", flush=True)
    alertas = []
    deptos_cuenca = ["pedernera", "general roca", "roque sáenz peña", "realicó", "chapaleufú", "trenel", "conhelo"]
    try:
        r = session.get("https://ws1.smn.gob.ar/v1/alerts/feed", timeout=8)
        if r.status_code == 200:
            data = r.json()
            for al in data:
                z = str(al.get("zone", "")).lower()
                desc = str(al.get("description", "")).lower()
                if any(d in z or d in desc for d in deptos_cuenca):
                    color = al.get("color", "amarillo").lower()
                    alertas.append({
                        "zona": al.get("zone", "Cuenca Río V"),
                        "fenomeno": al.get("event", "Tormenta"),
                        "nivel": color.capitalize(),
                        "descripcion": al.get("description", "")
                    })
        print(f"   -> [SMN OFICIAL]: {len(alertas)} alertas vigentes en cuenca.")
    except Exception as e:
        print(f"   [AVISO SMN]: Falla feed SMN: {e}")
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
# 8. GENERADOR DEL PORTAL WEB COMPLETO ESTILO MONITOR ZV
# =============================================================
def compilar_portal_web_monitor_zv(hidro_resumen, meteo_total, diag_onda, rayos, pronostico_ecmwf, alertas_smn):
    print("7. Generando interfaz web unificada (Estilo Monitor ZV)...", flush=True)

    # 1. Alertas de Lluvia y Pronóstico (Estilo Monitor ZV)
    alertas_lluvia_html = ""

    # Alertas SMN oficiales
    if alertas_smn:
        for al in alertas_smn:
            alertas_lluvia_html += f"""
            <div class="flex items-center justify-between p-3.5 rounded-lg bg-amber-50 border border-amber-300 text-amber-900 text-sm">
                <div class="flex items-center space-x-3">
                    <i class="fa-solid fa-triangle-exclamation text-amber-500 text-base"></i>
                    <span><b>Alerta Oficial SMN ({al['nivel']}):</b> {al['fenomeno']} previsto en {al['zona']}. {al['descripcion']}</span>
                </div>
                <span class="text-xs font-bold px-2 py-0.5 bg-amber-200 text-amber-800 rounded">SMN</span>
            </div>
            """

    # Pronóstico ECMWF
    for p in pronostico_ecmwf:
        if p["lluvia_maniana"] >= 15.0:
            bg_c = "bg-rose-50 border-rose-200 text-rose-900" if p["lluvia_maniana"] >= 35.0 else "bg-amber-50 border-amber-200 text-amber-900"
            ico_c = "text-rose-500" if p["lluvia_maniana"] >= 35.0 else "text-amber-500"
            alertas_lluvia_html += f"""
            <div class="flex items-center justify-between p-3.5 rounded-lg {bg_c} border text-sm">
                <div class="flex items-center space-x-3">
                    <i class="fa-solid fa-cloud-showers-heavy {ico_c} text-base"></i>
                    <span><b>Lluvia pronosticada (ECMWF):</b> {p['nodo']} — se prevé <b>{p['lluvia_maniana']:.1f} mm</b> para mañana.</span>
                </div>
                <span class="text-xs font-medium text-slate-500">ECMWF IFS</span>
            </div>
            """

    # Lluvias observadas en vivo en REM / APA
    lluvias_significativas = [m for m in meteo_total if m.get("lluvia_24h_mm", 0) >= 20.0]
    for m in lluvias_significativas:
        alertas_lluvia_html += f"""
        <div class="flex items-center justify-between p-3.5 rounded-lg bg-rose-50 border border-rose-200 text-rose-900 text-sm">
            <div class="flex items-center space-x-3">
                <i class="fa-solid fa-droplet text-rose-500 text-base"></i>
                <span><b>Lluvia registrada (24h):</b> {m['nombre']} ({m['red']}) — acumulado de <b>{m['lluvia_24h_mm']:.1f} mm</b>.</span>
            </div>
            <span class="text-xs font-bold px-2 py-0.5 bg-rose-200 text-rose-800 rounded">Observado</span>
        </div>
        """

    if not alertas_lluvia_html:
        alertas_lluvia_html = """
        <div class="flex items-center justify-between p-3.5 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-900 text-sm">
            <div class="flex items-center space-x-3">
                <i class="fa-solid fa-circle-check text-emerald-600 text-base"></i>
                <span><b>Sin alertas de precipitaciones:</b> El modelo ECMWF no prevé eventos extremos para las próximas 48h y no rigen alertas del SMN en la cuenca.</span>
            </div>
            <span class="text-xs font-bold px-2 py-0.5 bg-emerald-200 text-emerald-800 rounded">Normal</span>
        </div>
        """

    # 2. Alerta de Traslación de Onda
    if diag_onda["alerta_activa"]:
        onda_badge = f'<span class="text-xs font-bold px-2.5 py-1 bg-rose-200 text-rose-800 rounded">{diag_onda["estado_alerta"]}</span>'
        onda_card = f"""
        <div class="flex items-center justify-between p-3.5 rounded-lg bg-rose-50 border border-rose-200 text-rose-900 text-sm">
            <div class="flex items-center space-x-3">
                <i class="fa-solid fa-triangle-exclamation text-rose-500 text-base"></i>
                <span><b>Alerta Hidrológica activa en {diag_onda['origen_alerta']}:</b> Nivel {diag_onda['nivel_origen']:.2f} m. Tiempo estimado de arribo a límite pampeano: <b>{diag_onda['tiempo_viaje_min_dias']:.0f} a {diag_onda['tiempo_viaje_max_dias']:.0f} días</b> (Ventana: {diag_onda['fecha_arribo_estimada_str']}). {diag_onda['factor_almacenamiento']}.</span>
            </div>
            {onda_badge}
        </div>
        """
    else:
        onda_card = f"""
        <div class="flex items-center justify-between p-3.5 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-900 text-sm">
            <div class="flex items-center space-x-3">
                <i class="fa-solid fa-circle-check text-emerald-600 text-base"></i>
                <span><b>Cuenca en calma ordinaria:</b> Ventana teórica estimada a La Pampa: <b>{diag_onda['tiempo_viaje_min_dias']:.0f} a {diag_onda['tiempo_viaje_max_dias']:.0f} días</b> ante eventual pulso en nacientes. Laguna La Margarita: <b>{diag_onda['nivel_margarita']:.2f} m</b> ({diag_onda['factor_almacenamiento']}).</span>
            </div>
            <span class="text-xs font-bold px-2.5 py-1 bg-emerald-200 text-emerald-800 rounded">Calma</span>
        </div>
        """

    # 3. Filas Tabla Limnígrafos
    filas_limnigrafos = ""
    for h in hidro_resumen:
        badge_color = "bg-emerald-100 text-emerald-800" if h["estado"] == "Normal" else ("bg-amber-100 text-amber-800" if h["estado"] == "Precaución" else "bg-rose-100 text-rose-800")
        filas_limnigrafos += f"""
        <tr class="hover:bg-slate-50 transition border-b border-slate-100">
            <td class="px-4 py-3 font-semibold text-slate-800">{h['nombre']}</td>
            <td class="px-4 py-3 text-slate-500">{h['rio']} ({h['distrito']})</td>
            <td class="px-4 py-3 text-right font-bold text-slate-900">{h['nivel_actual']:.2f} m</td>
            <td class="px-4 py-3 text-right text-slate-600">{h['media_hist']:.2f} m</td>
            <td class="px-4 py-3 text-right text-amber-700 font-semibold">{h['cota_alerta']:.2f} m</td>
            <td class="px-4 py-3 text-right text-rose-700 font-semibold">{h['cota_evac']:.2f} m</td>
            <td class="px-4 py-3 text-center"><span class="px-2 py-0.5 rounded text-xs font-medium {badge_color}">{h['estado']}</span></td>
            <td class="px-4 py-3 text-center font-medium text-slate-700">{h['tendencia']}</td>
            <td class="px-4 py-3 text-center text-xs text-slate-400">{h['fecha']}</td>
        </tr>
        """

    # 4. Tarjetas Nodos ECMWF Pronóstico
    cards_ecmwf = ""
    for p in pronostico_ecmwf:
        cards_ecmwf += f"""
        <div class="p-4 rounded-xl border border-slate-200 bg-white shadow-sm flex flex-col justify-between">
            <div>
                <span class="text-xs font-semibold uppercase text-slate-400">{p['provincia']}</span>
                <h4 class="text-base font-bold text-slate-800">{p['nodo']}</h4>
            </div>
            <div class="grid grid-cols-3 gap-2 mt-4 text-center border-t border-slate-100 pt-3">
                <div class="bg-slate-50 p-2 rounded-lg">
                    <span class="text-[10px] uppercase font-bold text-slate-400">Hoy</span>
                    <p class="text-sm font-bold text-slate-800">{p['lluvia_hoy']} mm</p>
                </div>
                <div class="bg-blue-50 p-2 rounded-lg">
                    <span class="text-[10px] uppercase font-bold text-blue-500">Mañana</span>
                    <p class="text-sm font-bold text-blue-700">{p['lluvia_maniana']} mm</p>
                </div>
                <div class="bg-slate-50 p-2 rounded-lg">
                    <span class="text-[10px] uppercase font-bold text-slate-400">Pasado</span>
                    <p class="text-sm font-bold text-slate-800">{p['lluvia_pasado']} mm</p>
                </div>
            </div>
        </div>
        """

    # Datos JSON para Leaflet
    geo_hidro = json.dumps(hidro_resumen)
    geo_meteo = json.dumps(meteo_total)
    geo_rayos = json.dumps(rayos)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Monitor del Clima — Cuenca Triprovincial Río V</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" />
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background-color: #f8fafc; }}
        .tab-btn.active {{ background-color: #ffffff; color: #0f172a; font-weight: 600; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
        .tab-btn {{ color: #64748b; transition: all 0.15s ease-in-out; }}
        #mapa-container {{ height: calc(100vh - 165px); min-height: 540px; }}
    </style>
</head>
<body class="text-slate-800 antialiased min-h-screen flex flex-col">

    <!-- HEADER ESTILO MONITOR ZV -->
    <header class="bg-white border-b border-slate-200 px-6 py-3 flex justify-between items-center sticky top-0 z-50">
        <div class="flex items-center space-x-3">
            <i class="fa-solid fa-cloud-sun-rain text-blue-600 text-xl"></i>
            <div>
                <h1 class="text-lg font-bold tracking-tight text-slate-900 leading-none">Monitor del Clima — Cuenca Río V</h1>
                <span class="text-xs text-slate-500 font-medium">SAT Triprovincial: San Luis · Córdoba · La Pampa | Actualizado: {FECHA_TXT}</span>
            </div>
        </div>
        <div class="flex items-center space-x-3">
            <button onclick="location.reload()" class="inline-flex items-center space-x-2 text-xs font-semibold px-3 py-1.5 rounded-lg border border-slate-300 hover:bg-slate-50 text-slate-700 transition">
                <i class="fa-solid fa-rotate text-slate-500"></i>
                <span>Actualizar</span>
            </button>
            <a href="{EXCEL_SALIDA}" download class="inline-flex items-center space-x-1.5 text-xs font-semibold px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white transition shadow-sm">
                <i class="fa-solid fa-file-excel"></i>
                <span>Descargar Excel</span>
            </a>
        </div>
    </header>

    <!-- NAVEGACIÓN EN PESTAÑAS -->
    <nav class="bg-slate-100/90 border-b border-slate-200 px-6 py-2 sticky top-[57px] z-40">
        <div class="flex space-x-1 overflow-x-auto text-sm">
            <button onclick="cambiarTab('panel')" id="btn-panel" class="tab-btn active px-4 py-2 rounded-lg">Panel</button>
            <button onclick="cambiarTab('limnigrafos')" id="btn-limnigrafos" class="tab-btn px-4 py-2 rounded-lg">Limnígrafos</button>
            <button onclick="cambiarTab('precipitaciones')" id="btn-precipitaciones" class="tab-btn px-4 py-2 rounded-lg">Precipitaciones & ECMWF</button>
            <button onclick="cambiarTab('traslacion')" id="btn-traslacion" class="tab-btn px-4 py-2 rounded-lg">Onda de Crecida</button>
            <button onclick="cambiarTab('mapa')" id="btn-mapa" class="tab-btn px-4 py-2 rounded-lg">Visor Cartográfico</button>
            <button onclick="cambiarTab('archivo')" id="btn-archivo" class="tab-btn px-4 py-2 rounded-lg">Descargas & Documentación</button>
        </div>
    </nav>

    <!-- CONTENIDO PRINCIPAL -->
    <main class="flex-1 p-6 max-w-7xl mx-auto w-full">

        <!-- TAB 1: PANEL (ESTILO MONITOR ZV) -->
        <section id="tab-panel" class="space-y-6">
            <!-- BLOQUE: LLUVIAS Y ALERTAS SMN/ECMWF -->
            <div class="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                <div class="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex justify-between items-center cursor-pointer">
                    <div class="flex items-center space-x-2 text-slate-800 font-semibold text-sm">
                        <i class="fa-solid fa-cloud-showers-heavy text-slate-600"></i>
                        <span>Lluvias Pronosticadas (ECMWF) y Alertas Meteorológicas Oficiales (SMN)</span>
                    </div>
                    <i class="fa-solid fa-chevron-right text-xs text-slate-400"></i>
                </div>
                <div class="p-4 space-y-2.5">
                    {alertas_lluvia_html}
                </div>
            </div>

            <!-- BLOQUE: TRASLACIÓN DE ONDA -->
            <div class="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                <div class="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex justify-between items-center cursor-pointer">
                    <div class="flex items-center space-x-2 text-slate-800 font-semibold text-sm">
                        <i class="fa-solid fa-water text-slate-600"></i>
                        <span>Onda de Tormenta & Estimación de Traslación a La Pampa</span>
                    </div>
                    <i class="fa-solid fa-chevron-right text-xs text-slate-400"></i>
                </div>
                <div class="p-4">
                    {onda_card}
                </div>
            </div>

            <!-- BLOQUE: LIMNÍGRAFOS SÍNTESIS -->
            <div class="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                <div class="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex justify-between items-center cursor-pointer">
                    <div class="flex items-center space-x-2 text-slate-800 font-semibold text-sm">
                        <i class="fa-solid fa-gauge-high text-slate-600"></i>
                        <span>Limnígrafos — Niveles en Cuerpos de Agua</span>
                    </div>
                    <i class="fa-solid fa-chevron-right text-xs text-slate-400"></i>
                </div>
                <div class="p-4">
                    <div class="flex items-center justify-between p-3.5 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-900 text-sm">
                        <div class="flex items-center space-x-3">
                            <i class="fa-solid fa-circle-check text-emerald-600 text-base"></i>
                            <span><b>Estado hidrométrico normal:</b> {len(hidro_resumen)} estaciones monitoreadas en San Luis y Córdoba con niveles por debajo de cotas de riesgo.</span>
                        </div>
                        <button onclick="cambiarTab('limnigrafos')" class="text-xs font-semibold text-blue-600 hover:underline">Ver tabla completa →</button>
                    </div>
                </div>
            </div>
        </section>

        <!-- TAB 2: TABLA LIMNÍGRAFOS -->
        <section id="tab-limnigrafos" class="hidden bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden p-5">
            <h2 class="text-base font-bold text-slate-900 mb-4">Red Limnimétrica Oficial — Cuenca Río V (INA / SNIH)</h2>
            <div class="overflow-x-auto">
                <table class="w-full text-left text-sm text-slate-600">
                    <thead class="bg-slate-50 text-slate-700 uppercase font-semibold text-xs border-b border-slate-200">
                        <tr>
                            <th class="px-4 py-3">Estación</th>
                            <th class="px-4 py-3">Cuerpo de Agua</th>
                            <th class="px-4 py-3 text-right">Nivel Act.</th>
                            <th class="px-4 py-3 text-right">Media Hist.</th>
                            <th class="px-4 py-3 text-right">Cota Alerta</th>
                            <th class="px-4 py-3 text-right">Cota Evac.</th>
                            <th class="px-4 py-3 text-center">Semáforo</th>
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

        <!-- TAB 3: PRECIPITACIONES Y PRONÓSTICO ECMWF -->
        <section id="tab-precipitaciones" class="hidden space-y-6">
            <div class="bg-white border border-slate-200 rounded-xl shadow-sm p-5">
                <h2 class="text-base font-bold text-slate-900 mb-1">Pronóstico Numérico del Tiempo — Modelo ECMWF IFS 0.25°</h2>
                <p class="text-xs text-slate-500 mb-4">Previsión de precipitaciones acumuladas diarias en los puntos de control clave de la cuenca.</p>
                <div class="grid grid-cols-1 md:grid-cols-4 gap-4">
                    {cards_ecmwf}
                </div>
            </div>

            <div class="bg-white border border-slate-200 rounded-xl shadow-sm p-5">
                <h2 class="text-base font-bold text-slate-900 mb-3">Redes Pluviométricas en Tiempo Real (REM San Luis + APA La Pampa)</h2>
                <div class="grid grid-cols-2 md:grid-cols-4 gap-3 text-sm">
                    <div class="p-3 bg-slate-50 rounded-lg border border-slate-200">
                        <span class="text-xs text-slate-500">Estaciones REM San Luis</span>
                        <p class="text-lg font-bold text-slate-800">{len([m for m in meteo_total if 'REM' in m['red']])} en línea</p>
                    </div>
                    <div class="p-3 bg-slate-50 rounded-lg border border-slate-200">
                        <span class="text-xs text-slate-500">Estaciones APA La Pampa</span>
                        <p class="text-lg font-bold text-slate-800">{len([m for m in meteo_total if 'APA' in m['red']])} consolidadas</p>
                    </div>
                    <div class="p-3 bg-slate-50 rounded-lg border border-slate-200">
                        <span class="text-xs text-slate-500">Catálogo Omixom Cba/LP</span>
                        <p class="text-lg font-bold text-slate-800">{len([m for m in meteo_total if 'Omixom' in m['red']])} estaciones</p>
                    </div>
                    <div class="p-3 bg-slate-50 rounded-lg border border-slate-200">
                        <span class="text-xs text-slate-500">Actividad Eléctrica</span>
                        <p class="text-lg font-bold text-amber-600">{len(rayos)} rayos detectados</p>
                    </div>
                </div>
            </div>
        </section>

        <!-- TAB 4: TRASLACIÓN DE ONDA -->
        <section id="tab-traslacion" class="hidden bg-white border border-slate-200 rounded-xl shadow-sm p-6 space-y-6">
            <h2 class="text-base font-bold text-slate-900">Modelo Dinámico de Traslación de Onda — Cuenca Río V</h2>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-6 bg-blue-50/70 border border-blue-200 rounded-xl p-6 text-center">
                <div>
                    <span class="text-xs font-bold text-blue-600 uppercase">Tiempo Estimado a La Pampa</span>
                    <p class="text-3xl font-extrabold text-blue-950 mt-1">{diag_onda['tiempo_viaje_min_dias']:.0f} a {diag_onda['tiempo_viaje_max_dias']:.0f} Días</p>
                    <span class="text-xs text-slate-500">Ventana teórica calculada</span>
                </div>
                <div>
                    <span class="text-xs font-bold text-blue-600 uppercase">Amortiguación Laguna La Margarita</span>
                    <p class="text-3xl font-extrabold text-blue-950 mt-1">{diag_onda['nivel_margarita']:.2f} m</p>
                    <span class="text-xs text-slate-500">{diag_onda['factor_almacenamiento']}</span>
                </div>
                <div>
                    <span class="text-xs font-bold text-blue-600 uppercase">Estado en Cabecera (San Luis)</span>
                    <p class="text-3xl font-extrabold text-emerald-600 mt-1">{diag_onda['estado_alerta']}</p>
                    <span class="text-xs text-slate-500">{diag_onda['origen_alerta']}</span>
                </div>
            </div>
        </section>

        <!-- TAB 5: VISOR CARTOGRÁFICO LEAFLET -->
        <section id="tab-mapa" class="hidden bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden p-2">
            <div id="mapa-container" class="w-full rounded-lg"></div>
        </section>

        <!-- TAB 6: ARCHIVO Y DESCARGAS -->
        <section id="tab-archivo" class="hidden bg-white border border-slate-200 rounded-xl shadow-sm p-6">
            <h2 class="text-base font-bold text-slate-900 mb-4">Informes Ejecutivos & Archivo de Datos</h2>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <a href="{EXCEL_SALIDA}" download class="flex items-center p-4 border border-slate-200 rounded-xl hover:border-emerald-500 hover:bg-emerald-50/30 transition">
                    <i class="fa-solid fa-file-excel text-emerald-600 text-3xl mr-4"></i>
                    <div>
                        <h4 class="font-bold text-slate-800">Planilla Multisolapa Excel</h4>
                        <p class="text-xs text-slate-500">Datos consolidados de niveles INA, REM San Luis, APA La Pampa y cruces.</p>
                    </div>
                </a>
                <a href="{CSV_SALIDA}" download class="flex items-center p-4 border border-slate-200 rounded-xl hover:border-blue-500 hover:bg-blue-50/30 transition">
                    <i class="fa-solid fa-file-csv text-blue-600 text-3xl mr-4"></i>
                    <div>
                        <h4 class="font-bold text-slate-800">Cruce Pluvio-Hidrométrico (CSV)</h4>
                        <p class="text-xs text-slate-500">Archivo liviano para análisis en SIG, Power BI o Python.</p>
                    </div>
                </a>
            </div>
        </section>
    </main>

    <!-- SCRIPTS DE CONTROL -->
    <script>
        var mapaLeaflet = null;
        var datosHidro = {geo_hidro};
        var datosMeteo = {geo_meteo};
        var datosRayos = {geo_rayos};

        function cambiarTab(tabId) {{
            ['panel', 'limnigrafos', 'precipitaciones', 'traslacion', 'mapa', 'archivo'].forEach(function(t) {{
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
                setTimeout(iniciarMapa, 150);
            }}
        }}

        function iniciarMapa() {{
            if (mapaLeaflet !== null) {{
                mapaLeaflet.invalidateSize();
                return;
            }}

            mapaLeaflet = L.map('mapa-container').setView([-34.5, -64.8], 7);

            var capaBase = L.tileLayer('https://{{s}}.basemaps.cartocdn.com/light_all/{{z}}/{{x}}/{{y}}{{r}}.png', {{
                attribution: '&copy; OpenStreetMap, &copy; CARTO'
            }}).addTo(mapaLeaflet);

            var layerHidro = L.featureGroup().addTo(mapaLeaflet);
            var layerMeteo = L.featureGroup().addTo(mapaLeaflet);
            var layerRayos = L.featureGroup().addTo(mapaLeaflet);
            var layerRadares = L.featureGroup().addTo(mapaLeaflet);
            var layerSat = L.featureGroup().addTo(mapaLeaflet);
            var layerRadarComp = L.featureGroup().addTo(mapaLeaflet);

            // Capa Cuerpos de Agua
            datosHidro.forEach(function(h) {{
                var marker = L.circleMarker([h.lat, h.lon], {{
                    radius: 8,
                    fillColor: h.color,
                    color: "#ffffff",
                    weight: 2,
                    opacity: 1,
                    fillOpacity: 0.9
                }}).bindPopup("<b>" + h.nombre + "</b><br>Nivel: " + h.nivel_actual.toFixed(2) + " m (" + h.estado + ")<br>Alerta: " + h.cota_alerta + " m | Evac: " + h.cota_evac + " m");
                layerHidro.addLayer(marker);
            }});

            // Capa Estaciones Meteo
            datosMeteo.forEach(function(m) {{
                var color = m.red.indexOf('San Luis') !== -1 ? '#f59e0b' : (m.red.indexOf('APA') !== -1 ? '#0284c7' : '#8b5cf6');
                var marker = L.circleMarker([m.lat, m.lon], {{
                    radius: 5,
                    fillColor: color,
                    color: "#ffffff",
                    weight: 1.5,
                    fillOpacity: 0.85
                }}).bindPopup("<b>" + m.nombre + "</b> (" + m.red + ")<br>Lluvia 24h: " + m.lluvia_24h_mm.toFixed(1) + " mm");
                layerMeteo.addLayer(marker);
            }});

            // Capa Rayos
            datosRayos.forEach(function(ry) {{
                var marker = L.circleMarker([ry.lat, ry.lon], {{
                    radius: 5,
                    fillColor: "#eab308",
                    color: "#a16207",
                    weight: 1,
                    fillOpacity: 0.95
                }}).bindTooltip("⚡ Rayo: " + ry.hora + " hs");
                layerRayos.addLayer(marker);
            }});

            // Conos de cobertura SINARAME
            L.circle([-36.226, -66.884], {{ radius: 120000, color: "#0284c7", fillOpacity: 0.05, dashArray: "5, 5" }}).bindTooltip("RMA08 Santa Isabel").addTo(layerRadares);
            L.circle([-33.725, -65.386], {{ radius: 120000, color: "#d97706", fillOpacity: 0.05, dashArray: "5, 5" }}).bindTooltip("RMA16 Villa Reynolds").addTo(layerRadares);

            // Capas satelitales en vivo
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

            var overlays = {{
                "💧 Limnígrafos (INA)": layerHidro,
                "🌦️ Redes Meteorológicas": layerMeteo,
                "⚡ Descargas Eléctricas": layerRayos,
                "📡 Cobertura SINARAME": layerRadares,
                "🛰️ Satélite GOES-16 IR": layerSat,
                "🌧️ Radar de Lluvias Compuesto": layerRadarComp
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
# 10. EJECUCIÓN DEL PIPELINE COMPLETO
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

    # 5. Cálculo Traslación de Onda
    diag_onda = calcular_tiempo_viaje_onda(hidro_resumen, rayos_cuenca)

    # 6. Generar Entregables
    generar_entregables_excel_csv(hidro_resumen, total_meteo, diag_onda)
    compilar_portal_web_monitor_zv(hidro_resumen, total_meteo, diag_onda, rayos_cuenca, pronostico_ecmwf, alertas_smn)

    print("=" * 70)
    print("PROCESO COMPLETADO. El archivo 'index.html' ya está listo para publicar.")
    print("=" * 70)