import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import time
import glob
import asyncio
import aiohttp
import pycountry
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.chart import BarChart, Reference
import fitz  # PyMuPDF
from datetime import datetime, timezone, timedelta
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Configuración general y estilos de la página (ZONA HORARIA QUITO)
ZONA_HORARIA_QUITO = timezone(timedelta(hours=-5))

def obtener_hora_quito():
    """Obtiene la fecha y hora actual exacta en Quito, Ecuador (UTC-5)."""
    return datetime.now(ZONA_HORARIA_QUITO)

# Configuración de la página
st.set_page_config(
    page_title="Dashboard WAF & Sustentación de Tesis - UCG",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS institucional (Diseño Azul F5 / Modo Oscuro)
st.markdown("""
    <style>
    .stApp { background-color: #0F172A; color: #F8FAFC; }
    .kpi-card { border-radius: 12px; padding: 16px 12px; text-align: center; color: white; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.08); }
    .kpi-tot { background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%); }
    .kpi-alto { background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%); }
    .kpi-med { background: linear-gradient(135deg, #D97706 0%, #B45309 100%); }
    .kpi-bajo { background: linear-gradient(135deg, #16A34A 0%, #15803D 100%); }
    .kpi-title { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px; opacity: 0.95; }
    .kpi-number { font-size: 34px; font-weight: 800; margin-top: 2px; }
    .section-header { background-color: #1E293B; padding: 12px 20px; border-radius: 10px; border-left: 5px solid #0284C7; margin-bottom: 15px; }
    h1, h2, h3 { color: #F8FAFC !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    [data-testid="stSidebar"] { background-color: #0284C7; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; }
    .author-credit { font-size: 11px; color: #94A3B8; font-style: italic; margin-top: -10px; margin-bottom: 15px; }
    </style>
""", unsafe_allow_html=True)

ARCHIVO_EXCEL = "base_visitas.xlsx"
ARCHIVO_LOGS = "registro_ingresos.csv"

# Configuración de APIs para Escaneo de IPs
IPINFO_API_TOKEN = "1610cc1e3d24a1"
VIRUSTOTAL_API_KEY = "2087ece6d1897627c3f204176af5cd0e413a7e4e111cff06b992308a90f5db50"
ABUSEIPDB_API_KEY = "ffafa47c3907b66766d7a3af9446530de589e6ea7d525780691b867962aef24c64b68020b358bf33"

RETRY_ATTEMPTS = 3
RETRY_DELAY = 2

CATEGORIAS_ABUSEIPDB = {
    3: "Phishing",