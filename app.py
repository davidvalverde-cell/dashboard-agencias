import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
import time

# ==========================================
# CONFIGURACIÓN DE ARCHIVOS DE PERSISTENCIA
# ==========================================
EXCEL_PATH = "base_visitas.xlsx"
LOG_FILE = "historial_ingresos.csv"

st.set_page_config(page_title="Dashboard de Agencias y Auditoría", layout="wide")


# ==========================================
# 1. MANEJO DE REGISTRO / LOG DE USUARIOS (ACUMULATIVO)
# ==========================================
def registrar_ingreso(usuario: str, rol: str):
    """
    Registra el inicio de sesión en modo 'append' sin borrar jamás el historial previo.
    """
    try:
        nuevo_registro = pd.DataFrame([{
            "Usuario": usuario,
            "Rol": rol,
            "FechaHora": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }])
        
        # Si el archivo no existe, escribe cabecera; si existe, añade al final
        header_needed = not os.path.exists(LOG_FILE)
        nuevo_registro.to_csv(LOG_FILE, mode='a', header=header_needed, index=False)
    except Exception as e:
        st.error(f"Error al registrar log de auditoría: {e}")

def cargar_historial_logs() -> pd.DataFrame:
    """
    Carga el historial acumulado de auditoría de usuarios.
    """
    if os.path.exists(LOG_FILE) and os.path.getsize(LOG_FILE) > 0:
        try:
            df = pd.read_csv(LOG_FILE)
            df['FechaHora'] = pd.to_datetime(df['FechaHora'])
            return df
        except Exception:
            return pd.DataFrame(columns=["Usuario", "Rol", "FechaHora"])
    return pd.DataFrame(columns=["Usuario", "Rol", "FechaHora"])


# ==========================================
# 2. LECTURA Y GUARDADO SEGURO (ATÓMICO) DE EXCEL
# ==========================================
@st.cache_data(ttl=1)
def cargar_datos(path=EXCEL_PATH):
    """
    Carga los datos del Excel de forma segura evitando lecturas sobre archivos en escritura.
    """
    # Si el archivo no existe, crea una base inicial de prueba
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        df_init = pd.DataFrame({
            "Región": ["Norte", "Centro", "Sur"],
            "Zona": ["Juan Perez", "Carlos Mendoza", "Luis Fernandez"],
            "Código": [184, 189, 188],
            "NombreAgencia": ["184-PIAZZA L", "189-GUAYAQUIL S", "188-QUITO CENTRO"],
            "Anio": [2025, 2025, 2025],
            "Mes": [9, 9, 9],
            "Dia": [7, 7, 7],
            "Semana": ["Semana 1", "Semana 1", "Semana 1"],
            "Hora": [10, 11, 12],
            "Terminal": ["CAJA01", "I4V1", "CAJA02"],
            "NombreMaq": ["CAJACAJA01", "CAJAI4V1", "CAJACAJA02"],
            "TipoTran": ["CONSULTA", "DEPÓSITO", "RETIRO"],
            "FechaVisita": ["2025-09-07", "2025-09-07", "2025-09-07"],
            "Calidad": [2.5, 2.8, 2.2],
            "ProcesosTransaccionales": [2.4, 2.9, 2.1],
            "ManejoEfectivoControl": [2.6, 2.7, 2.9],
            "TalentoCultura": [2.8, 2.5, 2.6]
        })
        df_init.to_excel(path, index=False)

    # Reintentos de lectura en caso de bloqueo momentáneo de disco
    for _ in range(3):
        try:
            df = pd.read_excel(path)
            col_agencia = "NombreAgencia" if "NombreAgencia" in df.columns else df.columns[3]
            col_region = "Región" if "Región" in df.columns else df.columns[0]
            col_coord = "Zona" if "Zona" in df.columns else df.columns[1]
            col_fecha = "FechaVisita" if "FechaVisita" in df.columns else df.columns[12]
            cols_eval = [c for c in ["Calidad", "ProcesosTransaccionales", "ManejoEfectivoControl", "TalentoCultura"] if c in df.columns]
            return df, col_agencia, col_region, col_coord, col_fecha, cols_eval
        except Exception:
            time.sleep(0.3)
            
    # Lectura final de respaldo
    df = pd.read_excel(path)
    return df, df.columns[3], df.columns[0], df.columns[1], df.columns[12], []

def guardar_datos_seguro(df: pd.DataFrame, path=EXCEL_PATH) -> bool:
    """
    Guarda los datos usando un archivo temporal y reemplazo atómico para EVITAR el error al pulsar Guardar varias veces.
    """
    temp_path = f"{path}.tmp"
    try:
        # 1. Guardar primero en archivo temporal
        df.to_excel(temp_path, index=False)
        
        # 2. Reemplazo seguro de archivo
        if os.path.exists(temp_path):
            os.replace(temp_path, path)
            
        # 3. Limpiar caché de Streamlit
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"No se pudo guardar el archivo: {e}")
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass
        return False


# ==========================================
# 3. SESIÓN DEL USUARIO Y REGISTRO AUTOMÁTICO
# ==========================================
if "usuario_actual" not in st.session_state:
    st.session_state["usuario_actual"] = "gabriela.armendariz"
    st.session_state["rol_actual"] = "Administrador"
    
if "ingreso_registrado" not in st.session_state:
    registrar_ingreso(st.session_state["usuario_actual"], st.session_state["rol_actual"])
    st.session_state["ingreso_registrado"] = True


# ==========================================
# 4. CARGA INICIAL DE DATOS
# ==========================================
df_base, col_agencia, col_region, col_coord, col_fecha, cols_eval = cargar_datos()


# ==========================================
# 5. ESTRUCTURA DE PESTAÑAS (INTERFAZ)
# ==========================================
tab1, tab2 = st.tabs(["📁 Carga y Gestión de Matrices", "📋 Log de Usuarios Registrados"])

# ------------------------------------------
# PESTAÑA 1: GESTIÓN DE MATRICES
# ------------------------------------------
with tab1:
    st.title("📁 Carga y Edición de Matrices de Visitas")
    st.write("Edite directamente los registros o cargue una nueva matriz. Los cambios se guardan de forma permanente.")

    # Edición directa de la base de datos
    df_editado = st.data_editor(
        df_base,
        use_container_width=True,
        num_rows="dynamic",
        key="editor_matrices"
    )

    st.divider()
    col_g1, col_g2 = st.columns([1, 4])

    with col_g1:
        # BOTÓN GUARDAR SIN ERRORES MULTI-CLIC
        if st.button("💾 Guardar Cambios", type="primary", use_container_width=True):
            exito = guardar_datos_seguro(df_editado)
            if exito:
                st.success("✅ ¡Datos e historial guardados exitosamente!")
                time.sleep(0.5)
                st.rerun()

    with col_g2:
        st.info("💡 Al presionar 'Guardar Cambios', el archivo de matriz y el registro histórico se conservan de forma permanente.")

# ------------------------------------------
# PESTAÑA 2: AUDITORÍA DE INGRESOS CON WIDGETS
# ------------------------------------------
with tab2:
    st.title("📋 Historial de Ingresos de Usuarios")
    st.caption("A continuación se muestra el registro acumulado con fecha y hora de cada usuario que ha iniciado sesión.")

    df_logs = cargar_historial_logs()

    if not df_logs.empty:
        # --- WIDGETS DINÁMICOS (KPIs) ---
        st.subheader("📊 Indicadores Clave de Acceso")

        total_ingresos = len(df_logs)
        usuarios_unicos = df_logs['Usuario'].nunique()
        hoy_str = date.today().strftime("%Y-%m-%d")
        ingresos_hoy = df_logs[df_logs['FechaHora'].dt.strftime("%Y-%m-%d") == hoy_str].shape[0]

        k1, k2, k3 = st.columns(3)
        k1.metric(label="Total de Ingresos Históricos", value=f"{total_ingresos:,}")
        k2.metric(label="Usuarios Únicos Registrados", value=f"{usuarios_unicos:,}")
        k3.metric(label="Ingresos Realizados Hoy", value=f"{ingresos_hoy:,}")

        st.divider()

        # --- FILTROS DE AUDITORÍA ---
        cf1, cf2 = st.columns(2)
        with cf1:
            filtro_usr = st.text_input("🔍 Buscar por Usuario:", "")
        with cf2:
            f_min = df_logs['FechaHora'].min().date() if not df_logs.empty else date.today()
            rango_f = st.date_input("📅 Rango de Fechas:", value=(f_min, date.today()))

        # Aplicación de filtros
        df_filtrado = df_logs.copy()
        if filtro_usr:
            df_filtrado = df_filtrado[df_filtrado['Usuario'].str.contains(filtro_usr, case=False, na=False)]

        if isinstance(rango_f, tuple) and len(rango_f) == 2:
            df_filtrado = df_filtrado[
                (df_filtrado['FechaHora'].dt.date >= rango_f[0]) & 
                (df_filtrado['FechaHora'].dt.date <= rango_f[1])
            ]

        # Ordenar más recientes primero
        df_filtrado = df_filtrado.sort_values(by="FechaHora", ascending=False)
        df_display = df_filtrado.copy()
        df_display['FechaHora'] = df_display['FechaHora'].dt.strftime("%Y-%m-%d %H:%M:%S")

        # --- TABLA DE LOGS ---
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Usuario": st.column_config.TextColumn("Usuario", width="medium"),
                "Rol": st.column_config.TextColumn("Rol", width="medium"),
                "FechaHora": st.column_config.TextColumn("FechaHora", width="medium")
            }
        )
    else:
        st.info("No hay registros de ingreso acumulados hasta el momento.")