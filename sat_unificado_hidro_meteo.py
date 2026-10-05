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
ESTACIONES_OM
