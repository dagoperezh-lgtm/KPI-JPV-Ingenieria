import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
from datetime import datetime, timedelta
from docx import Document

# --- SECCIÓN 1: CONFIGURACIÓN Y MOTOR DE CARGA ---
# Configuración de la página
st.set_page_config(page_title="Dashboard de Excelencia Operativa", layout="wide", page_icon="📈")

@st.cache_data
def generar_datos_prueba():
    """Genera datos de prueba robustos si no se sube un archivo."""
    np.random.seed(42)
    fechas_ingreso = [datetime(2025, 1, 1) + timedelta(days=int(x)) for x in np.random.randint(0, 365, 500)]
    datos = []
    for f in fechas_ingreso:
        dias_resolucion = np.random.randint(2, 20)
        fecha_cierre = f + timedelta(days=dias_resolucion + (dias_resolucion // 5) * 2) 
        dias_descargo = np.random.randint(1, dias_resolucion) if dias_resolucion > 1 else 1
        fecha_descargos = f + timedelta(days=dias_descargo)
        retrocesos = np.random.choice([0, 0, 0, 1, 2])
        datos.append({
            "ID_Caso": f"CASO-{np.random.randint(100000, 999999)}",
            "Fecha_Ingreso": f,
            "Fecha_Descargos": fecha_descargos,
            "Fecha_Cierre": fecha_cierre,
            "Tipologia": np.random.choice(["GPS", "Jornada", "Documentación", "Velocidad"]),
            "Retrocesos": retrocesos,
            "Estado": "Cerrado"
        })
    return pd.DataFrame(datos)

# Interfaz lateral de carga
st.sidebar.title("Configuración")
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
    st.sidebar.info("Usando datos de demostración.")
    df_raw = generar_datos_prueba()

# --- SECCIÓN 2: MOTOR DE CÁLCULO (PROCESAMIENTO DE KPIs) ---
def procesar_datos(df):
    """Calcula los tiempos de ciclo y métricas base."""
    df['Fecha_Ingreso'] = pd.to_datetime(df['Fecha_Ingreso'])
    
    # Manejo seguro de columnas faltantes
    if 'Fecha_Descargos' not in df.columns:
        df['Fecha_Descargos'] = df['Fecha_Ingreso']
    if 'Retrocesos' not in df.columns:
        df['Retrocesos'] = 0
        
    df['Fecha_Descargos'] = pd.to_datetime(df['Fecha_Descargos'])
    df['Fecha_Cierre'] = pd.to_datetime(df['Fecha_Cierre'])
    
    # Extracción de periodos
    df['Mes'] = df['Fecha_Cierre'].dt.to_period('M').astype(str)
    df['Trimestre'] = df['Fecha_Cierre'].dt.to_period('Q').astype(str)
    df['Año'] = df['Fecha_Cierre'].dt.year.astype(str)
    df['Mes_Ingreso'] = df['Fecha_Ingreso'].dt.to_period('M').astype(str)
    df['Trimestre_Ingreso'] = df['Fecha_Ingreso'].dt.to_period('Q').astype(str)
    df['Año_Ingreso'] = df['Fecha_Ingreso'].dt.year.astype(str)
    
    # Cálculos de tiempo en días hábiles
    df_cerrados = df.dropna(subset=['Fecha_Ingreso', 'Fecha_Cierre']).copy()
    fechas_in = df_cerrados['Fecha_Ingreso'].values.astype('datetime64[D]')
    fechas_out = df_cerrados['Fecha_Cierre'].values.astype('datetime64[D]')
    df_cerrados['Dias_Gestion'] = np.busday_count(fechas_in, fechas_out)
    
    fechas_desc = df_cerrados['Fecha_Descargos'].values.astype('datetime64[D]')
    fechas_desc = np.where(fechas_desc > fechas_out, fechas_in, fechas_desc)
    df_cerrados['Dias_Activos'] = np.busday_count(fechas_desc, fechas_out)
    
    # Reglas de negocio (SLA y Calidad)
    df_cerrados['Cumple_SLA'] = df_cerrados['Dias_Gestion'] <= 15
    df_cerrados['First_Time_Right'] = df_cerrados['Retrocesos'] == 0
    
    return df_cerrados

df_procesado = procesar_datos(df_raw)

# Filtros de interfaz
st.sidebar.header("Filtros de Análisis")
tipo_periodo = st.sidebar.radio("Seleccionar Agrupación:", ["Mensual", "Trimestral", "Anual"])

if tipo_periodo == "Mensual":
    periodos_disp = sorted(df_procesado['Mes'].unique(), reverse=True)
    col_cierre, col_ingreso = 'Mes', 'Mes_Ingreso'
elif tipo_periodo == "Trimestral":
    periodos_disp = sorted(df_procesado['Trimestre'].unique(), reverse=True)
    col_cierre, col_ingreso = 'Trimestre', 'Trimestre_Ingreso'
else:
    periodos_disp = sorted(df_procesado['Año'].unique(), reverse=True)
    col_cierre, col_ingreso = 'Año', 'Año_Ingreso'

periodo_seleccionado = st.sidebar.selectbox("Seleccionar Periodo:", periodos_disp)

# Cruce de datos
df_filtrado_cierre = df_procesado[df_procesado[col_cierre] == periodo_seleccionado]
df_filtrado_ingreso = df_procesado[df_procesado[col_ingreso] == periodo_seleccionado]

# Resultados de KPIs matemáticos
lead_time_promedio = df_filtrado_cierre['Dias_Gestion'].mean()
cycle_time_promedio = df_filtrado_cierre['Dias_Activos'].mean()
total_casos_cerrados = len(df_filtrado_cierre)
sla_compliance = (df_filtrado_cierre['Cumple_SLA'].sum() / total_casos_cerrados * 100) if total_casos_cerrados > 0 else 0
first_time_right = (df_filtrado_cierre['First_Time_Right'].sum() / total_casos_cerrados * 100) if total_casos_cerrados > 0 else 0
throughput = total_casos_cerrados
total_ingresos_periodo = len(df_filtrado_ingreso)
tasa_traccion = (throughput / total_ingresos_periodo * 100) if total_ingresos_periodo > 0 else 100

# --- SECCIÓN 3: MOTOR DE REPORTE VISUAL (DASHBOARD) ---
st.title("📊 Panel de Control de Alto Rendimiento")
st.markdown("Visibilizando la agilidad, el flujo continuo y la excelencia resolutiva del equipo.")
st.subheader(f"Resultados del Periodo: {periodo_seleccionado}")

# Fila 1: KPIs Capacidad
st.markdown("##### 1. Capacidad y Salud del Portafolio")
col1, col2, col3 = st.columns(3)
with col1:
    st.metric(label="✅ Índice de Entrega (Throughput)", value=f"{throughput} Casos")
with col2:
    st.metric(label="📥 Ingresos del Periodo", value=f"{total_ingresos_periodo} Casos")
with col3:
    st.metric(label="⚖️ Tasa de Tracción", value=f"{tasa_traccion:.1f}%", help="Cierres vs Ingresos.")

# Fila 2: KPIs Agilidad
st.markdown("##### 2. Agilidad y Flujo Continuo")
col4, col5, col6 = st.columns(3)
with col4:
    st.metric(label="⚡ Tiempo Medio Resolución (Lead Time)", value=f"{lead_time_promedio:.1f} Días")
with col5:
    st.metric(label="⏱️ Cycle Time Activo (Post-Descargos)", value=f"{cycle_time_promedio:.1f} Días")

# Fila 3: KPIs Calidad
st.markdown("##### 3. Excelencia y Cumplimiento")
col7, col8, col9 = st.columns(3)
with col7:
    st.metric(label="⭐ Resolución Óptima (SLA)", value=f"{sla_compliance:.1f}%")
with col8:
    st.metric(label="🎯 Calidad en Origen (First-Time Right)", value=f"{first_time_right:.1f}%")

st.divider()

# Gráficos
col_chart1, col_chart2 = st.columns(2)
with col_chart1:
    st.markdown("### Agilidad por Tipología")
    if not df_filtrado_cierre.empty:
        df_agrupado_tipo = df_filtrado_cierre.groupby('Tipologia')['Dias_Gestion'].mean().reset_index()
        fig1 = px.bar(df_agrupado_tipo, x='Tipologia', y='Dias_Gestion', text_auto='.1f',
                      title="Tiempo Promedio de Cierre por Tipo de Caso", color_discrete_sequence=['#2ecc71'])
        st.plotly_chart(fig1, use_container_width=True)

with col_chart2:
    st.markdown("### Cumplimiento del Estándar (SLA)")
    if not df_filtrado_cierre.empty:
        df_sla = df_filtrado_cierre['Cumple_SLA'].value_counts().reset_index()
        df_sla.columns = ['Cumple Meta', 'Cantidad']
        df_sla['Cumple Meta'] = df_sla['Cumple Meta'].map({True: 'Óptimo', False: 'Fuera de Plazo'})
        fig2 = px.pie(df_sla, values='Cantidad', names='Cumple Meta', title="Distribución de Casos según SLA",
                      color='Cumple Meta', color_discrete_map={'Óptimo': '#27ae60', 'Fuera de Plazo': '#f39c12'})
        st.plotly_chart(fig2, use_container_width=True)

# --- SECCIÓN 4: MOTOR DE REPORTES (EXCEL Y WORD) ---
import matplotlib.pyplot as plt
from docx.shared import Inches

st.divider()
st.subheader("📥 Reportes de Gestión Ejecutiva")
st.markdown("Descarga los informes formales listos para presentar o archivar.")

# Función generadora de Excel en memoria
@st.cache_data
def generar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Datos')
    return output.getvalue()

# Función generadora de Gráficos estáticos para Word
def generar_grafico_tendencia(fechas, valores, titulo, ylabel, color, legend=None, valores2=None, color2=None, legend2=None):
    plt.figure(figsize=(8, 4))
    plt.plot(fechas, valores, marker='o', color=color, label=legend, linewidth=2)
    if valores2 is not None:
        plt.plot(fechas, valores2, marker='s', color=color2, label=legend2, linewidth=2)
    plt.title(titulo, fontsize=12, fontweight='bold')
    plt.ylabel(ylabel, fontsize=10)
    plt.xticks(rotation=45, ha='right', fontsize=8)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    if legend or legend2:
        plt.legend()
    plt.tight_layout()
    
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=150)
    plt.close()
    img_stream.seek(0)
    return img_stream

# Función generadora de Word con contexto Histórico
def generar_word_historico(tipo_per, periodo_actual, df_completo):
    doc = Document()
    doc.add_heading(f'Reporte Ejecutivo de Gestión', 0)
    doc.add_paragraph(f'Tipo de Análisis: {tipo_per}')
    doc.add_paragraph(f'Periodo Evaluado: {periodo_actual} y Contexto Histórico')
    
    # 0. Agrupar y preparar datos históricos
    col_periodo = 'Mes' if tipo_per == 'Mensual' else 'Trimestre' if tipo_per == 'Trimestral' else 'Año'
    col_periodo_in = 'Mes_Ingreso' if tipo_per == 'Mensual' else 'Trimestre_Ingreso' if tipo_per == 'Trimestral' else 'Año_Ingreso'
    
    historico_cierres = df_completo.dropna(subset=['Fecha_Cierre']).groupby(col_periodo).agg(
        Throughput=('ID_Caso', 'count'),
        Lead_Time=('Dias_Gestion', 'mean'),
        Cycle_Time=('Dias_Activos', 'mean'),
        SLA_Cumple=('Cumple_SLA', 'sum'),
        FTR_Cumple=('First_Time_Right', 'sum')
    ).reset_index()
    
    historico_cierres['SLA_Compliance'] = (historico_cierres['SLA_Cumple'] / historico_cierres['Throughput']) * 100
    historico_cierres['FTR'] = (historico_cierres['FTR_Cumple'] / historico_cierres['Throughput']) * 100
    
    historico_ingresos = df_completo.dropna(subset=['Fecha_Ingreso']).groupby(col_periodo_in).agg(
        Ingresos=('ID_Caso', 'count')
    ).reset_index().rename(columns={col_periodo_in: col_periodo})
    
    historico = pd.merge(historico_cierres, historico_ingresos, on=col_periodo, how='outer').fillna(0)
    historico = historico.sort_values(col_periodo).tail(12) # Últimos 12 periodos para claridad visual
    historico['Tasa_Traccion'] = np.where(historico['Ingresos'] > 0, (historico['Throughput'] / historico['Ingresos']) * 100, 100)
    
    # Extraer valores exactos del periodo actual
    datos_actuales = historico[historico[col_periodo] == periodo_actual]
    if not datos_actuales.empty:
        t_put = datos_actuales['Throughput'].values[0]
        lead_t = datos_actuales['Lead_Time'].values[0]
        cycle_t = datos_actuales['Cycle_Time'].values[0]
        sla_comp = datos_actuales['SLA_Compliance'].values[0]
        ftr = datos_actuales['FTR'].values[0]
        traccion = datos_actuales['Tasa_Traccion'].values[0]
    else:
        t_put, lead_t, cycle_t, sla_comp, ftr, traccion = 0, 0, 0, 0, 0, 0

    # --- SECCIÓN 1: CAPACIDAD ---
    doc.add_heading('1. Capacidad y Salud del Portafolio', level=1)
    doc.add_paragraph('Esta sección evalúa el volumen de trabajo gestionado y el equilibrio operativo del área, visibilizando la capacidad para mantener el portafolio fluido sin cuellos de botella.')
    
    doc.add_paragraph(f'• Índice de Entrega (Throughput): {int(t_put)} casos cerrados en {periodo_actual}.', style='List Bullet')
    doc.add_paragraph('Definición: Refleja el volumen exacto de casos cerrados exitosamente. Permite visualizar la capacidad operativa real y la productividad constante del equipo de analistas.')
    
    doc.add_paragraph(f'• Tasa de Tracción: {traccion:.1f}% en {periodo_actual}.', style='List Bullet')
    doc.add_paragraph('Definición: Mide la relación matemática entre los casos que ingresan y los que se logran resolver. Un porcentaje cercano o superior al 100% indica que el equipo procesa a un ritmo saludable, evitando atrasos.')
    
    img_capacidad = generar_grafico_tendencia(
        historico[col_periodo], historico['Throughput'], 
        'Tendencia Histórica: Índice de Entrega (Cierres) vs Ingresos', 'Cantidad de Casos', 
        '#27ae60', 'Casos Cerrados (Throughput)', historico['Ingresos'], '#2980b9', 'Casos Ingresados'
    )
    doc.add_picture(img_capacidad, width=Inches(6.0))

    # --- SECCIÓN 2: AGILIDAD ---
    doc.add_heading('2. Agilidad y Flujo Continuo', level=1)
    doc.add_paragraph('Esta sección mide la velocidad de respuesta del equipo, mostrando cómo ha evolucionado el tiempo total del caso frente al tiempo efectivo de análisis técnico.')
    
    doc.add_paragraph(f'• Tiempo Medio de Resolución (Lead Time): {lead_t:.1f} días hábiles en {periodo_actual}.', style='List Bullet')
    doc.add_paragraph('Definición: Promedio de días hábiles transcurridos desde que un caso ingresa hasta su resolución final. Visibiliza la agilidad global y la eficiencia del proceso.')
    
    doc.add_paragraph(f'• Cycle Time Activo (Post-Descargos): {cycle_t:.1f} días hábiles en {periodo_actual}.', style='List Bullet')
    doc.add_paragraph('Definición: Promedio de días que toma el análisis desde que se reciben los antecedentes hasta el cierre. Es el indicador puro de agilidad técnica interna, aislando tiempos de espera de terceros.')
    
    img_agilidad = generar_grafico_tendencia(
        historico[col_periodo], historico['Lead_Time'], 
        'Tendencia Histórica: Lead Time vs Cycle Time', 'Días Hábiles Promedio', 
        '#e74c3c', 'Lead Time Integral', historico['Cycle_Time'], '#f39c12', 'Cycle Time Activo'
    )
    doc.add_picture(img_agilidad, width=Inches(6.0))

    # --- SECCIÓN 3: CALIDAD ---
    doc.add_heading('3. Excelencia y Cumplimiento', level=1)
    doc.add_paragraph('Esta sección refleja el rigor técnico histórico, la calidad de la revisión inicial y el nivel de servicio respecto a los estándares establecidos.')
    
    doc.add_paragraph(f'• Resolución Óptima (SLA Compliance): {sla_comp:.1f}% en {periodo_actual}.', style='List Bullet')
    doc.add_paragraph('Definición: Porcentaje de casos cerrados cumpliendo estrictamente con el estándar de tiempo objetivo (<= 15 días hábiles). Métrica principal de nivel de servicio.')
    
    doc.add_paragraph(f'• Calidad en Origen (First-Time Right): {ftr:.1f}% en {periodo_actual}.', style='List Bullet')
    doc.add_paragraph('Definición: Porcentaje de casos que fluyeron sin devoluciones o reprocesos. Demuestra prolijidad en el ingreso de datos y madurez de los controles iniciales.')
    
    img_calidad = generar_grafico_tendencia(
        historico[col_periodo], historico['SLA_Compliance'], 
        'Tendencia Histórica: Cumplimiento SLA y Calidad en Origen (%)', 'Porcentaje de Cumplimiento (%)', 
        '#8e44ad', 'SLA Compliance (%)', historico['FTR'], '#16a085', 'First-Time Right (%)'
    )
    doc.add_picture(img_calidad, width=Inches(6.0))

    doc.add_heading('Conclusión Estratégica', level=1)
    doc.add_paragraph('La visibilización de estas tendencias históricas permite identificar ciclos de mejora, asegurar el flujo continuo y respaldar formalmente la gestión de excelencia mantenida por el equipo a lo largo del tiempo.')
    
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

col_d1, col_d2, col_d3 = st.columns(3)

# Botón 1: Data del Mes (Excel)
with col_d1:
    excel_data = generar_excel(df_filtrado_cierre)
    st.download_button(
        label="📊 Data del Periodo (Excel)",
        data=excel_data,
        file_name=f"Data_Gestion_{periodo_seleccionado}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# Botón 2: Reporte Formal (Word)
with col_d2:
    word_data = generar_word_historico(
        tipo_periodo, 
        periodo_seleccionado, 
        df_procesado  # Se pasa toda la base de datos para construir la historia
    )
    st.download_button(
        label="📝 Reporte Ejecutivo (Word)",
        data=word_data,
        file_name=f"Reporte_Ejecutivo_{tipo_periodo}_{periodo_seleccionado}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

# Botón 3: Backup Total (Excel)
with col_d3:
    excel_historico = generar_excel(df_procesado)
    st.download_button(
        label="📁 Backup Histórico Total (Excel)",
        data=excel_historico,
        file_name="Historico_Alto_Rendimiento.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
