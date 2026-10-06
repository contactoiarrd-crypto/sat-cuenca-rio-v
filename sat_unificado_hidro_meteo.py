# -*- coding: utf-8 -*-
"""
SISTEMA DE ALERTA TEMPRANA TRIPROVINCIAL: CUENCA RÍO V (SAN LUIS - CÓRDOBA - LA PAMPA)
Monitoreo Hidrometeorológico, Pronóstico ECMWF IFS, Parser Oficial SAT SMN,
Monitoreo de Cuerpos de Agua INA, Traslación de Onda y Visor Cartográfico para Toma de Decisiones.
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
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

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

LAT_MIN_CUENCA, LAT_MAX_CUENCA = -36.8, -32.0
LON_MIN_CUENCA, LON_MAX_CUENCA = -67.8, -62.8

session = requests.Session()
retries = Retry(total=2, backoff_factor=0.5, status_forcelist=[500, 502, 503, 504])
session.mount("https://", HTTPAdapter(max_retries=retries))
session.mount("http://", HTTPAdapter(max_retries=retries))

# =============================================================
# 2. CATÁLOGO REDES METEOROLÓGICAS (OMIXOM - CÓRDOBA Y LA PAMPA)
# =============================================================
ESTACIONES_OMIXOM = [
    {"id": "OMX_CBA_1", "nombre": "General Levalle", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.0000, "lon": -63.9163, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_2", "nombre": "Río Bamba", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.0537, "lon": -63.7332, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_3", "nombre": "Jovita", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.5194, "lon": -63.9683, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_4", "nombre": "Villa Valeria", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.3427, "lon": -64.9290, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_5", "nombre": "Nicolás Bruzzone", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.4391, "lon": -64.3424, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_6", "nombre": "Melo", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.3452, "lon": -63.4377, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_7", "nombre": "Huanchillas", "departamento": "Juárez Celman", "provincia": "Córdoba", "lat": -33.6665, "lon": -63.6395, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_8", "nombre": "Hipólito Bouchard", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.7060, "lon": -63.5030, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_9", "nombre": "General Roca", "departamento": "General Roca", "provincia": "Córdoba", "lat": -33.9948, "lon": -65.0754, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_10", "nombre": "Serrano", "departamento": "Roque Sáenz Peña", "provincia": "Córdoba", "lat": -34.4629, "lon": -63.5309, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_11", "nombre": "Coronel Moldes", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.6481, "lon": -64.5950, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_12", "nombre": "Viamonte", "departamento": "Unión", "provincia": "Córdoba", "lat": -33.7429, "lon": -63.0996, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_13", "nombre": "Chaján", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.5508, "lon": -65.0059, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_14", "nombre": "Villa Rossi", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.2949, "lon": -63.2654, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_15", "nombre": "Presa El Chañar", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.9608, "lon": -65.0551, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_16", "nombre": "Huinca Renancó", "departamento": "General Roca", "provincia": "Córdoba", "lat": -34.8208, "lon": -64.3738, "red": "Omixom Córdoba"},
    {"id": "OMX_CBA_17", "nombre": "Vicuña Mackenna", "departamento": "Río Cuarto", "provincia": "Córdoba", "lat": -33.9754, "lon": -64.3642, "red": "Omixom Córdoba"},
    {"id": "OMX_LP_1", "nombre": "El Tala - La Veneta", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.3119, "lon": -64.7144, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_2", "nombre": "Realicó", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.0576, "lon": -64.2129, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_3", "nombre": "Trilí", "departamento": "Quemú Quemú", "provincia": "La Pampa", "lat": -35.9036, "lon": -63.6429, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_4", "nombre": "Ingeniero Luiggi", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.4714, "lon": -64.6024, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_5", "nombre": "Conhelo", "departamento": "Conhelo", "provincia": "La Pampa", "lat": -35.9895, "lon": -64.5954, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_6", "nombre": "Arata", "departamento": "Trenel", "provincia": "La Pampa", "lat": -35.6391, "lon": -64.3564, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_7", "nombre": "Pichi Huinca", "departamento": "Rancul", "provincia": "La Pampa", "lat": -35.6482, "lon": -64.7699, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_8", "nombre": "Rancul", "departamento": "Rancul", "provincia": "La Pampa", "lat": -35.0883, "lon": -64.5082, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_9", "nombre": "Coronel Hilario Lagos", "departamento": "Chapaleufú", "provincia": "La Pampa", "lat": -35.0344, "lon": -63.9111, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_10", "nombre": "Colonia Barón", "departamento": "Quemú Quemú", "provincia": "La Pampa", "lat": -36.1508, "lon": -63.8550, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_11", "nombre": "Intendente Alvear", "departamento": "Chapaleufú", "provincia": "La Pampa", "lat": -35.3182, "lon": -63.6054, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_12", "nombre": "Alta Italia", "departamento": "Realicó", "provincia": "La Pampa", "lat": -35.3317, "lon": -64.1191, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_13", "nombre": "Winifreda", "departamento": "Conhelo", "provincia": "La Pampa", "lat": -36.2229, "lon": -64.2487, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_14", "nombre": "General Pico", "departamento": "Maracó", "provincia": "La Pampa", "lat": -35.6969, "lon": -63.6207, "red": "Omixom La Pampa"},
    {"id": "OMX_LP_15", "nombre": "Eduardo Castex", "departamento": "Conhelo", "provincia": "La Pampa", "lat": -35.9160, "lon": -64.2956, "red": "Omixom La Pampa"}
]

# =============================================================
# 3. CATÁLOGO CUERPOS DE AGUA Y COTAS FÍSICAS (INA / BDH)
# =============================================================
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
    "TRAPICHE": {"alerta": 1.80, "evac": 2.30, "base": 0.85},
    "VILLA MERCEDES": {"alerta": 2.80, "evac": 3.50, "base": 1.20},
    "CIRCUNVALACION": {"alerta": 1.60, "evac": 2.20, "base": 0.90},
    "JUSTO DARACT": {"alerta": 2.90, "evac": 3.60, "base": 1.10},
    "AJI": {"alerta": 1.90, "evac": 2.40, "base": 0.70},
    "R.N. 35": {"alerta": 2.60, "evac": 3.20, "base": 1.15},
    "DEVOTO": {"alerta": 2.80, "evac": 3.40, "base": 1.05},
    "MARGARITA": {"alerta": 2.40, "evac": 3.00, "base": 1.42},
    "RP Nº26": {"alerta": 2.20, "evac": 2.80, "base": 1.10},
    "RP 26": {"alerta": 2.20, "evac": 2.80, "base": 1.10}
}

def consultar_estacion_ina(est_id, info):
    """Consulta la API del INA para una estación con recuperación de fallos controlada."""
    f_desde = (ahora - timedelta(days=3)).strftime("%Y-%m-%d")
    f_hasta = (ahora + timedelta(days=1)).strftime("%Y-%m-%d")
    
    url = f"https://alerta.ina.gob.ar/pub/datos/series.php?estacion_id={est_id}&var_id=2&timestart={f_desde}&timeend={f_hasta}&formato=json"
    
    nivel_actual = None
    fecha_obs = "Sin reporte reciente"
    
    try:
        r = session.get(url, timeout=5, verify=False)
        if r.status_code == 200:
            datos = r.json()
            if isinstance(datos, list) and len(datos) > 0:
                ultima_valida = next((d for d in reversed(datos) if d.get("valor") is not None), None)
                if ultima_valida:
                    nivel_actual = round(float(ultima_valida["valor"]), 2)
                    fecha_obs = ultima_valida.get("timestart", ahora.strftime("%Y-%m-%d %H:%M"))
    except Exception:
        pass

    # Umbrales
    nombre_u = info["nombre"].upper()
    c_alerta, c_evac, c_base = 2.50, 3.20, 1.00
    for k, u in UMBRALES_NOMINALES.items():
        if k in nombre_u:
            c_alerta, c_evac, c_base = u["alerta"], u["evac"], u["base"]
            break

    if nivel_actual is None:
        nivel_actual = c_base
        fecha_obs = ahora.strftime("%d/%m %H:%M") + " (Cota Nominal)"

    if nivel_actual >= c_evac:
        color = "#ef4444"
        estado = "Emergencia / Evacuación"
    elif nivel_actual >= c_alerta:
        color = "#f59e0b"
        estado = "Alerta Preventiva"
    else:
        color = "#10b981"
        estado = "Vigilancia / Normal"

    return {
        "id": est_id,
        "nombre": info["nombre"],
        "rio": info["rio"],
        "distrito": info["distrito"],
        "tramo": info["tramo"],
        "lat": info["lat"],
        "lon": info["lon"],
        "nivel_actual": nivel_actual,
        "cota_alerta": c_alerta,
        "cota_evac": c_evac,
        "color": color,
        "estado": estado,
        "fecha": fecha_obs
    }

def obtener_datos_ina():
    print("1. Consultando niveles en cuerpos de agua (INA)...", flush=True)
    resumen_hidro = []
    with ThreadPoolExecutor(max_workers=5) as executor:
        futuros = {executor.submit(consultar_estacion_ina, eid, info): eid for eid, info in ESTACIONES_INA_CATALOGO.items()}
        for f in as_completed(futuros):
            resumen_hidro.append(f.result())
    
    # Ordenar por cuenca (de alta a baja)
    resumen_hidro.sort(key=lambda x: x["lat"], reverse=True)
    print(f"   -> [INA]: {len(resumen_hidro)} cuerpos de agua monitoreados.")
    return resumen_hidro

# =============================================================
# 4. TRASLACIÓN DE ONDA Y CAPACIDAD DE AMORTIGUACIÓN
# =============================================================
def calcular_traslacion_onda(resumen_hidro):
    dict_h = {h["nombre"]: h for h in resumen_hidro}
    margarita = next((h for k, h in dict_h.items() if "Margarita" in k), None)
    
    puntos_criticos = [h for h in resumen_hidro if h["estado"] in ["Alerta Preventiva", "Emergencia / Evacuación"]]
    
    nivel_margarita = margarita["nivel_actual"] if margarita else 1.42
    cota_alerta_margarita = margarita["cota_alerta"] if margarita else 2.40
    margen_remanente = round(max(0.0, cota_alerta_margarita - nivel_margarita), 2)
    porc_ocupacion = min(100, int((nivel_margarita / cota_alerta_margarita) * 100))

    if nivel_margarita < 1.10:
        factor_almacenamiento = "ALTA RETENCIÓN (Bañados deprimidos)"
        ajuste_dias = 3.5
    elif nivel_margarita >= 2.00 or margen_remanente <= 0.40:
        factor_almacenamiento = "SATURACIÓN CRÍTICA (Efecto vaso lleno)"
        ajuste_dias = -2.5
    else:
        factor_almacenamiento = "AMORTIGUACIÓN REGULAR"
        ajuste_dias = 0.0

    t_base_min, t_base_max = 6.0, 9.0
    origen_alerta = "Sin anomalía en nacientes"
    estado_alerta = "Vigilancia / Normal"
    fecha_deteccion = ahora.strftime("%d/%m/%Y %H:%M")

    if puntos_criticos:
        pt = puntos_criticos[0]
        origen_alerta = pt["nombre"]
        estado_alerta = pt["estado"]
        fecha_deteccion = pt["fecha"]
        if "Trapiche" in origen_alerta or "Villa Mercedes" in origen_alerta:
            t_base_min, t_base_max = 8.0, 11.0
        elif "Justo Daract" in origen_alerta:
            t_base_min, t_base_max = 7.0, 10.0
        elif "R.N. 35" in origen_alerta:
            t_base_min, t_base_max = 4.0, 7.0
        elif "RP Nº26" in origen_alerta:
            t_base_min, t_base_max = 2.0, 4.0

    t_est_min = max(1.5, t_base_min + ajuste_dias)
    t_est_max = max(t_est_min + 1.0, t_base_max + ajuste_dias)

    f_llegada_min = ahora + timedelta(days=t_est_min)
    f_llegada_max = ahora + timedelta(days=t_est_max)

    if estado_alerta == "Emergencia / Evacuación":
        protocolo_nivel = "NIVEL ROJO — ACCIÓN INMEDIATA Y PRE-EMERGENCIA"
        protocolo_desc = "Activación del Comité Interprovincial de Cuenca. Alerta vial en RN 35 y Meridiano V. Relevamiento continuo de bordos de contención y compuertas en Realicó e Intendente Alvear."
    elif estado_alerta == "Alerta Preventiva":
        protocolo_nivel = "NIVEL AMARILLO — PRE-ALERTA Y ENLACE TÉCNICO"
        protocolo_desc = "Enlace operativo entre Recursos Hídricos de Córdoba, San Luis y La Pampa. Monitoreo intensivo de vertederos y alcantarillas terciarias."
    else:
        protocolo_nivel = "NIVEL VERDE — VIGILANCIA PREVENTIVA / CALMA"
        protocolo_desc = "Monitoreo ordinario en nacientes y tramos regulados. Sin requerimiento de maniobras extraordinarias."

    return {
        "alerta_activa": len(puntos_criticos) > 0,
        "origen_alerta": origen_alerta,
        "estado_alerta": estado_alerta,
        "fecha_deteccion": fecha_deteccion,
        "nivel_margarita": nivel_margarita,
        "cota_alerta_margarita": cota_alerta_margarita,
        "margen_remanente": margen_remanente,
        "porc_ocupacion": porc_ocupacion,
        "factor_almacenamiento": factor_almacenamiento,
        "tiempo_viaje_min_dias": t_est_min,
        "tiempo_viaje_max_dias": t_est_max,
        "fecha_arribo_estimada": f"{f_llegada_min.strftime('%d/%m')} al {f_llegada_max.strftime('%d/%m/%Y')}",
        "protocolo_nivel": protocolo_nivel,
        "protocolo_desc": protocolo_desc
    }

# =============================================================
# 5. PRONÓSTICO METEOROLÓGICO ECMWF IFS (OPEN-METEO)
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
    print("2. Consultando pronóstico ECMWF IFS...", flush=True)
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
        r = session.get(url, timeout=8)
        if r.status_code == 200:
            datos = r.json()
            if not isinstance(datos, list):
                datos = [datos]
            for idx, p in enumerate(NODOS_ECMWF):
                d_met = datos[idx].get("daily", {})
                lluvias = d_met.get("precipitation_sum", [])
                
                ll_hoy = round(lluvias[0] if len(lluvias) > 0 and lluvias[0] is not None else 0.0, 1)
                ll_man = round(lluvias[1] if len(lluvias) > 1 and lluvias[1] is not None else 0.0, 1)
                ll_pas = round(lluvias[2] if len(lluvias) > 2 and lluvias[2] is not None else 0.0, 1)

                resultados.append({
                    "nodo": p["nombre"],
                    "provincia": p["provincia"],
                    "region": p["region"],
                    "lat": p["lat"],
                    "lon": p["lon"],
                    "lluvia_hoy": ll_hoy,
                    "lluvia_maniana": ll_man,
                    "lluvia_pasado": ll_pas,
                    "alerta": ll_man >= 25.0 or ll_pas >= 30.0
                })
        print(f"   -> [ECMWF]: {len(resultados)} nodos procesados.")
    except Exception as e:
        print(f"   [AVISO ECMWF]: {e}")
    return resultados

# =============================================================
# 6. PARSER OFICIAL SAT SMN (REST API + CAPAS VECTORIALES)
# =============================================================
DEPARTAMENTOS_CUENCA = [
    "GENERAL PEDERNERA", "CORONEL PRINGLES", "GOBERNADOR DUPUY",
    "GENERAL ROCA", "PRESIDENTE ROQUE SAENZ PENA", "RIO CUARTO", "JUAREZ CELMAN",
    "REALICO", "CHAPALEUFU", "RANCUL", "MARACO", "TRENEL", "CONHELO", "QUEMU QUEMU"
]

EVENTOS_SMN_NOMBRES = {
    41: "Tormentas fuertes o severas",
    42: "Vientos fuertes",
    39: "Lluvias abundantes",
    40: "Nevadas",
    45: "Viento Zonda",
    46: "Temperaturas extremas (Frío)",
    47: "Temperaturas extremas (Calor)",
    37: "Condición general",
    54: "Visibilidad reducida"
}

def obtener_alertas_smn():
    print("3. Consultando SAT oficial del SMN...", flush=True)
    alertas_tabla = []
    poligonos_leaflet = []

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "application/json"
    }

    url_api = "https://ws1.smn.gob.ar/v1/warning/alert/area?mode=alert&compact=true"
    url_geometrias = "https://www.smn.gob.ar/alertas/archivos/alertas_dia1.json"

    geometrias_dict = {}
    try:
        r_geo = session.get(url_geometrias, headers=headers, timeout=6)
        if r_geo.status_code == 200:
            for feat in r_geo.json().get("features", []):
                gid = feat.get("properties", {}).get("gid")
                if gid:
                    geometrias_dict[str(gid)] = feat.get("geometry", {})
    except Exception:
        pass

    try:
        r_alertas = session.get(url_api, headers=headers, timeout=8)
        if r_alertas.status_code == 200:
            for area in r_alertas.json():
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
                        if nivel_num >= 4:
                            nivel_str, color_hex = "Rojo", "#ef4444"
                        elif nivel_num == 3:
                            nivel_str, color_hex = "Naranja", "#ea580c"
                        else:
                            nivel_str, color_hex = "Amarillo", "#f59e0b"

                        eventos_list = [EVENTOS_SMN_NOMBRES.get(ev.get("id"), "Tormentas") for ev in w.get("events", []) if ev.get("max_level", 1) >= 2]
                        fenomeno = ", ".join(set(eventos_list)) if eventos_list else "Tormentas / Lluvias"

                        alertas_tabla.append({
                            "zona": area.get("name", f"Zona {area_id}"),
                            "fecha": w.get("date", "Hoy/Mañana"),
                            "fenomeno": fenomeno,
                            "nivel": nivel_str,
                            "color": color_hex,
                            "descripcion": f"Alerta {nivel_str} emitida por el SMN. Fenómeno: {fenomeno}."
                        })

                        # Extraer polígono
                        geom = geometrias_dict.get(area_id)
                        if geom:
                            coords = geom.get("coordinates", [])
                            # Manejo seguro de Polygons y MultiPolygons
                            anillos = []
                            if geom.get("type") == "MultiPolygon":
                                for poly in coords:
                                    if poly and len(poly) > 0:
                                        anillos.append([[pt[1], pt[0]] for pt in poly[0] if len(pt) >= 2])
                            elif geom.get("type") == "Polygon":
                                if coords and len(coords) > 0:
                                    anillos.append([[pt[1], pt[0]] for pt in coords[0] if len(pt) >= 2])

                            for anillo in anillos:
                                poligonos_leaflet.append({
                                    "coordenadas": anillo,
                                    "color": color_hex,
                                    "nivel": nivel_str,
                                    "evento": fenomeno,
                                    "zona": area.get("name", f"Área {area_id}")
                                })
    except Exception as e:
        print(f"   [AVISO SMN]: {e}")

    print(f"   -> [SAT SMN]: {len(alertas_tabla)} alerta(s) registradas para la región.")
    return alertas_tabla, poligonos_leaflet

# =============================================================
# 7. GENERADOR DE REPORTES (EXCEL Y CSV)
# =============================================================
def exportar_datos(hidro_datos, ecmwf_datos, smn_datos):
    # 1. Exportar CSV
    filas_csv = []
    for h in hidro_datos:
        filas_csv.append({
            "Tipo": "Cuerpo de Agua (INA)",
            "Nombre": h["nombre"],
            "Ubicacion": h["tramo"],
            "Latitud": h["lat"],
            "Longitud": h["lon"],
            "Nivel_Actual_m": h["nivel_actual"],
            "Cota_Alerta_m": h["cota_alerta"],
            "Cota_Evacuacion_m": h["cota_evac"],
            "Estado": h["estado"],
            "Actualizacion": h["fecha"]
        })
    df_csv = pd.DataFrame(filas_csv)
    df_csv.to_csv(CSV_SALIDA, index=False, encoding="utf-8-sig")

    # 2. Exportar Excel estilizado si openpyxl está presente
    if OPENPYXL_AVAILABLE:
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Cuerpos de Agua INA"
        
        headers = ["Estación / Punto", "Tramo", "Nivel (m)", "Cota Alerta", "Cota Evac", "Estado", "Última Observación"]
        ws1.append(headers)
        
        fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        font_header = Font(color="FFFFFF", bold=True)
        for col_num in range(1, len(headers) + 1):
            cell = ws1.cell(row=1, column=col_num)
            cell.fill = fill_header
            cell.font = font_header
