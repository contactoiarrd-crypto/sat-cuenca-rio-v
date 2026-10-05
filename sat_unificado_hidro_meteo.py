# -*- coding: utf-8 -*-
"""
SISTEMA DE ALERTA TEMPRANA TRIPROVINCIAL: CUENCA RÍO V (SAN LUIS - CÓRDOBA - LA PAMPA)
Monitoreo Hidrometeorológico, Pronóstico ECMWF, Parser Oficial CAP/RSS SMN con Polígonos en Mapa y Visor Dinámico.
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

MEDIAS_CLIMATICAS_CUENCA = {
    "Realicó": {"anual_mm": 800, "mes_esperado_mm": 75, "provincia": "La Pampa"},
    "Intendente Alvear": {"anual_mm": 820, "mes_esperado_mm": 78, "provincia": "La Pampa"},
    "Punta Alta (Rancul)": {"anual_mm": 780, "mes_esperado_mm": 72, "provincia": "La Pampa"},
    "General Levalle": {"anual_mm": 750, "mes_esperado_mm": 68, "provincia": "Córdoba"},
    "Jovita": {"anual_mm": 770, "mes_esperado_mm": 70, "provincia": "Córdoba"},
    "Villa Mercedes": {"anual_mm": 680, "mes_esperado_mm": 55, "provincia": "San Luis"}
}

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

LAT_MIN_SL, LAT_MAX_SL = -36.5, -32.0
LON_MIN_SL, LON_MAX_SL = -67.5, -64.5
LAT_MIN_CUENCA, LAT_MAX_CUENCA = -36.8, -32.0
LON_MIN_CUENCA, LON_MAX_CUENCA = -67.8, -62.8

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
    6444: {"nombre": "Trapiche - Hosteria El Trapiche", "rio": "Río Trapiche (Afluente)", "distrito": "San Luis", "tramo": "Cuenca Alta (San Luis)", "lat": -33.105833, "lon": -66.063333},
    6441: {"nombre": "Quinto - Dique Villa Mercedes", "rio": "Río Quinto", "distrito": "San Luis", "tramo": "Cuenca Alta (San Luis)", "lat": -33.653889, "lon": -65.533611},
    6472: {"nombre": "Quinto - Av Circunvalación", "rio": "Río Quinto", "distrito": "San Luis", "tramo": "Cuenca Alta (San Luis)", "lat": -33.739167, "lon": -65.376944},
    6445: {"nombre": "Quinto - Justo Daract", "rio": "Río Quinto", "distrito": "San Luis", "tramo": "Cuenca Alta (San Luis)", "lat": -33.918611, "lon": -65.151667},
    6624: {"nombre": "Aº El Aji - RN Nº35 y RN Nº7", "rio": "Afluente Río Quinto", "distrito": "Córdoba / LP", "tramo": "Cuenca Media (Córdoba)", "lat": -33.931111, "lon": -64.396389},
    6622: {"nombre": "Quinto - R.N. 35", "rio": "Río Quinto", "distrito": "Córdoba / LP", "tramo": "Cuenca Media (Córdoba)", "lat": -34.216389, "lon": -64.386111},
    6623: {"nombre": "Quinto - Canal Devoto RP 4", "rio": "Río Quinto / Canal Devoto", "distrito": "Córdoba / LP", "tramo": "Cuenca Media (Córdoba)", "lat": -33.933056, "lon": -63.449167},
    6391: {"nombre": "Laguna La Margarita", "rio": "Cuenca Río Quinto", "distrito": "Córdoba / LP", "tramo": "Cuenca Media (Córdoba)", "lat": -34.653333, "lon": -63.723056},
    2809: {"nombre": "Quinto - RP Nº26", "rio": "Río Quinto", "distrito": "Córdoba / LP", "tramo": "Ingreso Cuenca Baja (La Pampa)", "lat": -34.762778, "lon": -63.645000}
}

UMBRALES_NOMINALES = {
    "TRAPICHE": {"alerta": 1.80, "evac": 2.30},
    "VILLA MERCEDES": {"alerta": 2.80, "evac": 3.50},
    "CIRCUNVALACION": {"alerta": 1.60, "evac": 2.20},
    "JUSTO DARACT": {"alerta": 2.90, "evac": 3.60},
    "AJI": {"alerta": 1.90, "evac": 2.40},
    "R.N. 35": {"alerta": 2.60, "evac": 3.20},
    "DEVOTO": {"alerta": 2.80, "evac": 3.40},
    "MARGARITA": {"alerta": 2.40, "evac": 3.00},
    "RP Nº26": {"alerta": 2.20, "evac": 2.80},
    "RP 26": {"alerta": 2.20, "evac": 2.80}
}

# =============================================================
# 4. TRASLACIÓN DE ONDA Y CAPACIDAD DE AMORTIGUACIÓN
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
    
    if valor >= c_evac:
        return "#ef4444", "Emergencia / Evacuación", c_alerta, c_evac
    elif valor >= c_alerta:
        return "#f59e0b", "Alerta Preventiva", c_alerta, c_evac
    else:
        return "#10b981", "Vigilancia Preventiva / Calma Hidrológica", c_alerta, c_evac

def calcular_tiempo_viaje_onda(lista_hidro_resumen, lista_rayos=[]):
    dict_h = {h["nombre"]: h for h in lista_hidro_resumen}
    dique_vm = next((h for k, h in dict_h.items() if "Dique Villa Mercedes" in k or "Villa Mercedes" in k), None)
    daract = next((h for k, h in dict_h.items() if "Justo Daract" in k), None)
    margarita = next((h for k, h in dict_h.items() if "Margarita" in k), None)
    rn35 = next((h for k, h in dict_h.items() if "R.N. 35" in k), None)
    rp26 = next((h for k, h in dict_h.items() if "RP Nº26" in k or "RP 26" in k), None)

    origen_alerta = None
    nivel_origen = 0.0
    cota_alerta_origen = 0.0
    estado_alerta_origen = "Vigilancia Preventiva / Calma Hidrológica"
    fecha_deteccion = ahora.strftime("%d/%m/%Y %H:%M")

    for punto in [daract, dique_vm, rn35, rp26]:
        if punto and punto["estado"] in ["Alerta Preventiva", "Emergencia / Evacuación"]:
            origen_alerta = punto["nombre"]
            nivel_origen = punto["nivel_actual"]
            cota_alerta_origen = punto["cota_alerta"]
            estado_alerta_origen = punto["estado"]
            fecha_deteccion = punto["fecha"]
            break

    nivel_margarita = margarita["nivel_actual"] if margarita else 1.43
    cota_alerta_margarita = margarita["cota_alerta"] if margarita else 2.40
    margen_remanente_margarita = round(max(0.0, cota_alerta_margarita - nivel_margarita), 2)
    porcentaje_ocupacion_margarita = min(100, int((nivel_margarita / cota_alerta_margarita) * 100))

    if nivel_margarita < 1.00:
        factor_almacenamiento = "ALTA RETENCIÓN (Bañados deprimidos)"
        ajuste_dias = 4.0
    elif nivel_margarita >= 2.00 or margen_remanente_margarita <= 0.40:
        factor_almacenamiento = "SATURACIÓN CRÍTICA (Efecto vaso lleno)"
        ajuste_dias = -3.0
    else:
        factor_almacenamiento = "AMORTIGUACIÓN REGULAR"
        ajuste_dias = 0.0

    if origen_alerta and ("Villa Mercedes" in origen_alerta or "Trapiche" in origen_alerta):
        t_base_min, t_base_max = 8.0, 12.0
    elif origen_alerta and "Justo Daract" in origen_alerta:
        t_base_min, t_base_max = 7.0, 10.0
    elif origen_alerta and "R.N. 35" in origen_alerta:
        t_base_min, t_base_max = 5.0, 8.0
    elif origen_alerta and "RP Nº26" in origen_alerta:
        t_base_min, t_base_max = 3.0, 6.0
    else:
        t_base_min, t_base_max = 7.0, 10.0

    t_est_min = max(2.0, t_base_min + ajuste_dias)
    t_est_max = max(t_est_min + 1.0, t_base_max + ajuste_dias)

    f_llegada_min = ahora + timedelta(days=t_est_min)
    f_llegada_max = ahora + timedelta(days=t_est_max)

    alerta_activa = estado_alerta_origen in ["Alerta Preventiva", "Emergencia / Evacuación"]
    rayos_cuenca_alta = [r for r in lista_rayos if r.get("lat", 0) > -34.5]
    alerta_convectiva = len(rayos_cuenca_alta) >= 8

    if estado_alerta_origen == "Emergencia / Evacuación":
        protocolo_nivel = "NIVEL ROJO — ACCIÓN INMEDIATA Y PRE-EMERGENCIA"
        protocolo_desc = "Activación del Comité Provincial de Emergencias Hídricas. Alerta vial en RN 35 y pasos bajos de Meridiano V. Relevamiento continuo de bordos de contención y terraplenes en Realicó, Intendente Alvear e Ing. Luiggi. Guardia permanente en compuertas de derivación."
    elif estado_alerta_origen == "Alerta Preventiva" or alerta_convectiva:
        protocolo_nivel = "NIVEL AMARILLO — PRE-ALERTA Y ENLACE TÉCNICO"
        protocolo_desc = "Enlace técnico directo con Recursos Hídricos de Córdoba y San Luis. Seguimiento cada 3 horas de erogaciones y caudales en Presa El Chañar y vertedero de Laguna La Margarita. Verificación preventiva de alcantarillas y luces de puentes en red vial terciaria del norte pampeano."
    else:
        protocolo_nivel = "NIVEL VERDE — VIGILANCIA PREVENTIVA / CALMA HIDROLÓGICA"
        protocolo_desc = "Monitoreo automatizado continuo de rutina. Verificación ordinaria del estado de compuertas y canales de alivio interprovinciales. Sin movilización operativa extraordinaria requerida."

    return {
        "alerta_activa": alerta_activa,
        "alerta_convectiva": alerta_convectiva,
        "origen_alerta": origen_alerta or ("Detección de celdas convectivas" if alerta_convectiva else "Sin anomalía en nacientes"),
        "nivel_origen": nivel_origen,
        "estado_alerta": estado_alerta_origen,
        "fecha_deteccion": fecha_deteccion,
        "nivel_margarita": nivel_margarita,
        "cota_alerta_margarita": cota_alerta_margarita,
        "margen_remanente_margarita": margen_remanente_margarita,
        "porcentaje_ocupacion_margarita": porcentaje_ocupacion_margarita,
        "factor_almacenamiento": factor_almacenamiento,
        "tiempo_viaje_min_dias": t_est_min,
        "tiempo_viaje_max_dias": t_est_max,
        "fecha_arribo_estimada_str": f"{f_llegada_min.strftime('%d/%m')} al {f_llegada_max.strftime('%d/%m/%Y')}",
        "rayos_nacientes": len(rayos_cuenca_alta),
        "protocolo_nivel": protocolo_nivel,
        "protocolo_desc": protocolo_desc
    }

# =============================================================
# 5. MODELO NUMÉRICO ECMWF
# =============================================================
NODOS_ECMWF = [
    {"nombre": "Realicó", "lat": -35.04, "lon": -64.24, "provincia": "La Pampa", "region": "Norte Pampeano (RN 35)"},
    {"nombre": "Intendente Alvear", "lat": -35.24, "lon": -63.59, "provincia": "La Pampa", "region": "Norte Pampeano (Meridiano V)"},
    {"nombre": "Punta Alta (Rancul)", "lat": -35.21, "lon": -64.45, "provincia": "La Pampa", "region": "Norte Pampeano (RP 9)"},
    {"nombre": "Villa Mercedes", "lat": -33.67, "lon": -65.46, "provincia": "San Luis", "region": "Cabecera / Nacientes"},
    {"nombre": "General Levalle", "lat": -34.00, "lon": -63.92, "provincia": "Córdoba", "region": "Cuenca Media (Aporte)"},
    {"nombre": "Jovita", "lat": -34.52, "lon": -63.97, "provincia": "Córdoba", "region": "Cuenca Media-Baja"}
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
                    "region": p["region"],
                    "lat": p["lat"],
                    "lon": p["lon"],
                    "lluvia_hoy": lluvia_hoy,
                    "lluvia_maniana": lluvia_maniana,
                    "lluvia_pasado": lluvia_pasado,
                    "alerta": lluvia_maniana >= 20.0 or lluvia_pasado >= 25.0
                })
        print(f"   -> [ECMWF]: {len(resultados)} nodos procesados.")
    except Exception as e:
        print(f"   [AVISO ECMWF]: {e}")
    return resultados

# =============================================================
# 6. PARSER OFICIAL SAT SMN (API REST JSON Y POLÍGONOS GEOJSON)
# =============================================================

# Diccionario de fenómenos según id numérico del SAT
EVENTOS_SMN_NOMBRES = {
    41: "Tormentas fuertes o severas",
    42: "Vientos fuertes",
    39: "Lluvias abundantes",
    40: "Nevadas",
    45: "Viento Zonda",
    46: "Temperaturas extremas (Frío)",
    47: "Temperaturas extremas (Calor)",
    37: "Condición meteorológica general",
    54: "Visibilidad reducida"
}

def punto_en_poligono(lat, lon, vertices):
    """Algoritmo Ray-Casting para determinar si una coordenada está dentro del polígono."""
    n = len(vertices)
    adentro = False
    p1_lat, p1_lon = vertices[0]
    for i in range(n + 1):
        p2_lat, p2_lon = vertices[i % n]
        if lon > min(p1_lon, p2_lon):
            if lon <= max(p1_lon, p2_lon):
                if lat <= max(p1_lat, p2_lat):
                    if p1_lon != p2_lon:
                        lat_interseccion = (lon - p1_lon) * (p2_lat - p1_lat) / (p2_lon - p1_lon) + p1_lat
                    if p1_lat == p2_lat or lat <= lat_interseccion:
                        adentro = not adentro
        p1_lat, p1_lon = p2_lat, p2_lon
    return adentro

def poligono_interseca_cuenca(vertices_poligono):
    """Verifica si el polígono del SMN toca la Cuenca del Río V."""
    for v_lat, v_lon in vertices_poligono:
        if LAT_MIN_CUENCA <= v_lat <= LAT_MAX_CUENCA and LON_MIN_CUENCA <= v_lon <= LON_MAX_CUENCA:
            return True

    localidades_testigo = [
        (-33.67, -65.46),  # Villa Mercedes
        (-33.86, -65.18),  # Justo Daract
        (-34.52, -63.97),  # Jovita
        (-34.00, -63.92),  # General Levalle
        (-35.04, -64.24),  # Realicó
        (-35.24, -63.59),  # Intendente Alvear
        (-35.08, -64.51),  # Rancul
    ]
    for loc_lat, loc_lon in localidades_testigo:
        if punto_en_poligono(loc_lat, loc_lon, vertices_poligono):
            return True

    return False

def obtener_alertas_smn_cuenca(pronostico_ecmwf=[]):
    """
    Consume la API REST del SAT (https://ws1.smn.gob.ar/v1/warning/alert/area)
    y cruza las alertas a 24, 48 y 72 hs con los polígonos vectoriales de la cuenca.
    """
    print("2. Consultando API REST oficial del SAT (ws1.smn.gob.ar)...", flush=True)
    alertas = []
    poligonos_cap = []

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://www.smn.gob.ar/alertas"
    }

    url_api_alertas = "https://ws1.smn.gob.ar/v1/warning/alert/area?mode=alert&compact=true"
    
    # 1. Obtenemos el estado de alertas de la API
    datos_alertas = []
    try:
        r = session.get(url_api_alertas, headers=headers, timeout=12)
        if r.status_code == 200:
            datos_alertas = r.json()
    except Exception as e:
        print(f"   [AVISO API SMN]: {e}")

    # 2. Obtenemos o cargamos las geometrías de las áreas
    # Si tenés el archivo de geometrías guardado localmente en tu repositorio, lo carga directo
    geometrias_por_area = {}
    archivo_geo_cache = "geometrias_sat_smn.json"
    
    if os.path.exists(archivo_geo_cache):
        try:
            with open(archivo_geo_cache, "r", encoding="utf-8") as f:
                geometrias_por_area = json.load(f)
        except Exception:
            pass

    # Si no existe caché local o está vacía, intentamos descargar la capa base del SMN
    if not geometrias_por_area:
        url_geometrias = "https://www.smn.gob.ar/alertas/archivos/alertas_dia1.json"
        try:
            r_geo = session.get(url_geometrias, headers=headers, timeout=12)
            if r_geo.status_code == 200:
                geo_json = r_geo.json()
                for feat in geo_json.get("features", []):
                    gid = feat.get("properties", {}).get("gid")
                    geom = feat.get("geometry", {})
                    if gid and geom:
                        geometrias_por_area[str(gid)] = geom
        except Exception:
            pass

    # 3. Procesamos cada área y evaluamos coincidencia temporal y espacial
    for item in datos_alertas:
        area_id = str(item.get("area_id"))
        warnings = item.get("warnings", [])

        # Chequear si tiene alerta en las próximas 72 hs (max_level >= 2)
        alertas_futuras = [w for w in warnings if w.get("max_level", 1) >= 2]
        if not alertas_futuras:
            continue

        # Obtener geometría y convertir a coordenadas [lat, lon] para Leaflet
        geom = geometrias_por_area.get(area_id)
        vertices_lat_lon = []
        if geom:
            coords = geom.get("coordinates", [])
            # Si es MultiPolygon o Polygon, aplanamos al contorno exterior
            while len(coords) > 0 and isinstance(coords[0], list) and isinstance(coords[0][0], list):
                coords = coords[0]
            # Las coordenadas GeoJSON vienen [lon, lat], las invertimos a [lat, lon]
            for pt in coords:
                if len(pt) >= 2:
                    vertices_lat_lon.append([float(pt[1]), float(pt[0])])

        # Verificamos si toca la Cuenca del Río V (o si es una de las áreas conocidas de la cuenca)
        AREAS_CONOCIDAS_CUENCA = ["3343", "3358", "3362", "3363", "3365", "3366", "3378", "768", "807", "808", "809", "810", "811", "821", "824"]
        
        toca_cuenca = (area_id in AREAS_CONOCIDAS_CUENCA) or (vertices_lat_lon and poligono_interseca_cuenca(vertices_lat_lon))

        if toca_cuenca:
            for w in alertas_futuras:
                fecha_alerta = w.get("date")
                nivel_num = w.get("max_level", 1)

                # Definir nivel y color según la escala SMN
                if nivel_num >= 4:
                    nivel_str, color_hex = "Naranja / Rojo", "#dc2626"
                elif nivel_num == 3:
                    nivel_str, color_hex = "Naranja", "#ea580c"
                else:
                    nivel_str, color_hex = "Amarillo", "#f59e0b"

                # Identificar fenómenos activos
                eventos_activos = [EVENTOS_SMN_NOMBRES.get(ev.get("id"), "Tormentas") for ev in w.get("events", []) if ev.get("max_level", 1) >= 2]
                fenomeno_str = ", ".join(set(eventos_activos)) if eventos_activos else "Alerta meteorológica preventiva"

                # Info para la tabla y panel de control
                alertas.append({
                    "zona": f"Cuenca Río V - Área {area_id}",
                    "plazo": f"Fecha: {fecha_alerta}",
                    "fenomeno": fenomeno_str,
                    "nivel": nivel_str,
                    "descripcion": f"Alerta {nivel_str} emitida por el SMN para el día {fecha_alerta}. Fenómeno esperado: {fenomeno_str}."
                })

                # Polígono para dibujar en el mapa de Leaflet
                if vertices_lat_lon:
                    poligonos_cap.append({
                        "coordenadas": vertices_lat_lon,
                        "color": color_hex,
                        "nivel": nivel_str,
                        "evento": fenomeno_str,
                        "plazo": f"Válido: {fecha_alerta}",
                        "descripcion": f"Área {area_id}: {fenomeno_str} ({nivel_str})"
                    })

    print(f"   -> [SAT SMN]: {len(alertas)} alerta(s) y {len(poligonos_cap)} polígono(s) interceptados para la cuenca.")
    return alertas, poligonos_cap
