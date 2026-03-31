import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime, timedelta

# Configuración de la página
st.set_page_config(page_title="Dashboard de Excelencia Operativa", layout="wide", page_icon="📈")

# --- FUNCIONES DE CARGA Y PROCESAMIENTO ---
@st.cache_data
def generar_datos_prueba():
    """Genera datos de prueba positivos si no se sube un archivo para demostrar el flujo."""
    np.random.seed(42)
    fechas_ingreso = [datetime(2025, 1, 1) + timedelta(days=int(x)) for x in np.random.randint(0, 365, 500)]
    datos = []
    for f in fechas_ingreso:
        # Simulamos un equipo eficiente: cierres entre 2 y 20 días hábiles
        dias_resolucion = np.random.randint(2, 20)
        fecha_cierre = f + timedelta(days=dias_resolucion + (dias_resolucion // 5) * 2) 
        datos.append({
            "ID_Caso": f"CASO-{np.random.randint(100000, 999999)}",
            "Fecha_Ingreso": f,
            "Fecha_Cierre": fecha_cierre,
            "Tipologia": np.random.choice(["GPS", "Jornada", "Documentación", "Velocidad"]),
            "Estado": "Cerrado"
        })
    df = pd.DataFrame(datos)
    return df

def procesar_datos(df):
    """Calcula los tiempos de ciclo y métricas base."""
    df['Fecha_Ingreso'] = pd.to_datetime(df['Fecha_Ingreso'])
    df['Fecha_Cierre'] = pd.to_datetime(df['Fecha_Cierre'])
    
    # Extraer periodos para filtros
    df['Mes'] = df['Fecha_Cierre'].dt.to_period('M').astype(str)
    df['Trimestre'] = df['Fecha_Cierre'].dt.to_period('Q').astype(str)
    df['Año'] = df['Fecha_Cierre'].dt.year.astype(str)
    
    # Calcular Lead Time (Días Hábiles)
    # Filtramos nulos para evitar errores en busday_count
    df_cerrados = df.dropna(subset=['Fecha_Ingreso', 'Fecha_Cierre']).copy()
    
    # Convertir a formato fecha (sin hora) para numpy
    fechas_in = df_cerrados['Fecha_Ingreso'].values.astype('datetime64[D]')
    fechas_out = df_cerrados['Fecha_Cierre'].values.astype('datetime64[D]')
    
    df_cerrados['Dias_Gestion'] = np.busday_count(fechas_in, fechas_out)
    
    # SLA Compliance (Meta: <= 15 días)
    df_cerrados['Cumple_SLA'] = df_cerrados['Dias_Gestion'] <= 15
    
    return df_cerrados

# --- INTERFAZ DE USUARIO ---
st.title("📊 Panel de Control de Alto Rendimiento")
st.markdown("Visibilizando la agilidad, el flujo continuo y la excelencia resolutiva del equipo.")

# Carga de archivos
archivo_subido = st.sidebar.file_uploader("Cargar Planilla Maestra (CSV/Excel)", type=["csv", "xlsx"])

if archivo_subido is not None:
    try:
        if archivo_subido.name.endswith('.csv'):
            df_raw = pd.read_csv(archivo_subido)
        else:
            df_raw = pd.read_excel(archivo_subido)
        st.sidebar.success("¡Archivo cargado con éxito!")
    except Exception as e:
        st.sidebar.error(f"Error al leer el archivo: {e}")
        df_raw = generar_datos_prueba()
else:
    st.sidebar.info("Usando datos de demostración. Sube tu archivo para ver datos reales.")
    df_raw = generar_datos_prueba()

df_procesado = procesar_datos(df_raw)

# --- FILTROS DE TIEMPO ---
st.sidebar.header("Filtros de Análisis")
tipo_periodo = st.sidebar.radio("Seleccionar Agrupación Temporal:", ["Mensual", "Trimestral", "Anual"])

if tipo_periodo == "Mensual":
    periodos_disponibles = sorted(df_procesado['Mes'].unique(), reverse=True)
    columna_filtro = 'Mes'
elif tipo_periodo == "Trimestral":
    periodos_disponibles = sorted(df_procesado['Trimestre'].unique(), reverse=True)
    columna_filtro = 'Trimestre'
else:
    periodos_disponibles = sorted(df_procesado['Año'].unique(), reverse=True)
    columna_filtro = 'Año'

periodo_seleccionado = st.sidebar.selectbox("Seleccionar Periodo:", periodos_disponibles)

# Filtrar el dataframe
df_filtrado = df_procesado[df_procesado[columna_filtro] == periodo_seleccionado]

# --- CÁLCULO DE KPIs ---
total_casos_cerrados = len(df_filtrado)
lead_time_promedio = df_filtrado['Dias_Gestion'].mean()
sla_compliance = (df_filtrado['Cumple_SLA'].sum() / total_casos_cerrados * 100) if total_casos_cerrados > 0 else 0

# --- VISUALIZACIÓN DE KPIs ---
st.subheader(f"Resultados del Periodo: {periodo_seleccionado}")

col1, col2, col3 = st.columns(3)
with col1:
    st.metric(label="✅ Índice de Entrega (Throughput)", value=f"{total_casos_cerrados} Casos", help="Total de casos cerrados exitosamente.")
with col2:
    st.metric(label="⚡ Tiempo Medio de Resolución", value=f"{lead_time_promedio:.1f} Días", help="Promedio de días hábiles desde ingreso a cierre.")
with col3:
    st.metric(label="⭐ Tasa de Resolución Óptima (SLA)", value=f"{sla_compliance:.1f}%", help="Porcentaje de casos resueltos en 15 días o menos.")

st.divider()

# --- GRÁFICOS DE RENDIMIENTO ---
col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    st.markdown("### Agilidad por Tipología")
    df_agrupado_tipo = df_filtrado.groupby('Tipologia')['Dias_Gestion'].mean().reset_index()
    fig1 = px.bar(df_agrupado_tipo, x='Tipologia', y='Dias_Gestion', text_auto='.1f',
                  title="Tiempo Promedio de Cierre por Tipo de Caso",
                  labels={'Dias_Gestion': 'Días Hábiles', 'Tipologia': 'Tipo de Caso'},
                  color_discrete_sequence=['#2ecc71'])
    st.plotly_chart(fig1, use_container_width=True)

with col_chart2:
    st.markdown("### Cumplimiento del Estándar (SLA)")
    df_sla = df_filtrado['Cumple_SLA'].value_counts().reset_index()
    df_sla.columns = ['Cumple Meta (<= 15 días)', 'Cantidad']
    df_sla['Cumple Meta (<= 15 días)'] = df_sla['Cumple Meta (<= 15 días)'].map({True: 'Óptimo', False: 'Fuera de Plazo'})
    fig2 = px.pie(df_sla, values='Cantidad', names='Cumple Meta (<= 15 días)', 
                  title="Distribución de Casos según SLA",
                  color='Cumple Meta (<= 15 días)',
                  color_discrete_map={'Óptimo': '#27ae60', 'Fuera de Plazo': '#f39c12'})
    st.plotly_chart(fig2, use_container_width=True)

# --- REPORTES Y DESCARGAS (Botones de Valor) ---
st.divider()
st.subheader("📥 Reportes de Gestión")
st.markdown("Descarga los datos procesados para respaldar los logros del equipo.")

# Función para convertir dataframe a CSV descargable
@st.cache_data
def convertir_df(df):
    return df.to_csv(index=False).encode('utf-8')

csv = convertir_df(df_filtrado)

col_d1, col_d2 = st.columns(2)

with col_d1:
    st.download_button(
        label="Descargar Detalle del Periodo (CSV)",
        data=csv,
        file_name=f"Reporte_Gestion_{periodo_seleccionado}.csv",
        mime="text/csv",
        use_container_width=True
    )

with col_d2:
    # Generar consolidado histórico para descarga
    csv_historico = convertir_df(df_procesado)
    st.download_button(
        label="Descargar Base Histórica Completa (CSV)",
        data=csv_historico,
        file_name="Reporte_Historico_Alto_Rendimiento.csv",
        mime="text/csv",
        use_container_width=True
    )
