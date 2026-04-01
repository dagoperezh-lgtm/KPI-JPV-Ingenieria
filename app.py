import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt

# --- SECCIÓN 1: CONFIGURACIÓN Y MOTOR DE CARGA ---
st.set_page_config(page_title="Dashboard de Gestión: Procesos y Tendencias", layout="wide", page_icon="⚙️")

@st.cache_data
def generar_datos_prueba():
    """Genera datos de prueba robustos para visualizar tendencias reales."""
    np.random.seed(42)
    hoy = datetime(2026, 4, 1) 
    areas = ["Ingeniería y Energía", "Equipos Móviles"]
    liquidadores = ["Carlos Mendoza", "Ana Rojas", "Luis Silva", "Marta Pérez"]
    estados_principales = ["Ingreso", "Instrucción y Análisis", "Resolución", "Cerrado"]
    
    subestados_map = {
        "Ingreso": ["Recepción Inicial", "Esperando Antecedentes Básicos"],
        "Instrucción y Análisis": ["En Análisis Técnico", "Esperando Presupuestos", "Esperando Descargos", "Revisión de Evidencia"],
        "Resolución": ["Redacción de Informe", "Para Revisión de Jefatura", "Pendiente de Firma"],
        "Cerrado": ["Cierre Administrativo", "Resolución Emitida"]
    }

    fechas_ingreso = [hoy - timedelta(days=int(x)) for x in np.random.randint(2, 1000, 1500)]
    datos = []
    
    for f in fechas_ingreso:
        area = np.random.choice(areas, p=[0.3, 0.7]) 
        liquidador = np.random.choice(liquidadores)
        
        dias_desde_ingreso = (hoy - f).days
        prob_cerrado = min(0.95, dias_desde_ingreso / 45)
        es_cerrado = np.random.random() < prob_cerrado
        
        if es_cerrado:
            estado = "Cerrado"
            subestado = np.random.choice(subestados_map["Cerrado"])
            base_resolucion = np.random.randint(15, 60) if area == "Ingeniería y Energía" else np.random.randint(5, 30)
            dias_resolucion = min(base_resolucion, dias_desde_ingreso)
            fecha_cierre = f + timedelta(days=max(1, dias_resolucion))
            fecha_ultimo_cambio = fecha_cierre
        else:
            estado = np.random.choice(estados_principales[:-1])
            subestado = np.random.choice(subestados_map[estado])
            fecha_cierre = pd.NaT
            dias_estancado = np.random.randint(0, min(dias_desde_ingreso + 1, 60))
            fecha_ultimo_cambio = hoy - timedelta(days=dias_estancado)

        datos.append({
            "ID_Caso": f"CASO-{np.random.randint(100000, 999999)}",
            "Area_Negocio": area,
            "Liquidador": liquidador,
            "Estado_Actual": estado,
            "Subestado_Actual": subestado,
            "Fecha_Ingreso": f,
            "Fecha_Ultimo_Cambio": fecha_ultimo_cambio,
            "Fecha_Cierre": fecha_cierre
        })
        
    return pd.DataFrame(datos)

st.sidebar.title("Configuración y Carga")
archivo_subido = st.sidebar.file_uploader("Cargar Reporte de Casos (CSV/Excel)", type=["csv", "xlsx"])

# --- MOTOR DE HOMOLOGACIÓN DE DATOS ---
if archivo_subido is not None:
    try:
        # Selector para saltar filas en blanco del Excel
        filas_saltar = st.sidebar.number_input("Filas a saltar (Encabezado desfasado)", min_value=0, max_value=20, value=0)
        
        if archivo_subido.name.endswith('.csv'):
            df_crudo = pd.read_csv(archivo_subido, skiprows=filas_saltar)
        else:
            df_crudo = pd.read_excel(archivo_subido, skiprows=filas_saltar)
            
        columnas_reales = df_crudo.columns.tolist()
        st.sidebar.success("Archivo leído. Mapea las columnas clave:")
        
        # Mapeo dinámico para evitar KeyErrors
        col_id = st.sidebar.selectbox("Columna ID Caso", columnas_reales, index=0)
        col_area = st.sidebar.selectbox("Columna Área de Negocio", columnas_reales, index=1 if len(columnas_reales)>1 else 0)
        col_liq = st.sidebar.selectbox("Columna Liquidador", columnas_reales, index=2 if len(columnas_reales)>2 else 0)
        col_estado = st.sidebar.selectbox("Columna Estado", columnas_reales, index=3 if len(columnas_reales)>3 else 0)
        col_subestado = st.sidebar.selectbox("Columna Subestado", columnas_reales, index=4 if len(columnas_reales)>4 else 0)
        col_fecha_in = st.sidebar.selectbox("Columna Fecha Ingreso", columnas_reales, index=5 if len(columnas_reales)>5 else 0)
        col_fecha_mod = st.sidebar.selectbox("Columna Último Cambio", columnas_reales, index=6 if len(columnas_reales)>6 else 0)
        col_fecha_out = st.sidebar.selectbox("Columna Fecha Cierre", columnas_reales, index=7 if len(columnas_reales)>7 else 0)

        df_raw = df_crudo.rename(columns={
            col_id: "ID_Caso",
            col_area: "Area_Negocio",
            col_liq: "Liquidador",
            col_estado: "Estado_Actual",
            col_subestado: "Subestado_Actual",
            col_fecha_in: "Fecha_Ingreso",
            col_fecha_mod: "Fecha_Ultimo_Cambio",
            col_fecha_out: "Fecha_Cierre"
        })
        
    except Exception as e:
        st.sidebar.error(f"Error al procesar el archivo: {e}")
        df_raw = generar_datos_prueba()
else:
    st.sidebar.info("Usando datos de demostración.")
    df_raw = generar_datos_prueba()

# --- SECCIÓN 2: MOTOR DE CÁLCULO Y SEGMENTACIÓN ---
def procesar_datos_integrales(df):
    df['Fecha_Ingreso'] = pd.to_datetime(df['Fecha_Ingreso'], errors='coerce')
    df['Fecha_Ultimo_Cambio'] = pd.to_datetime(df['Fecha_Ultimo_Cambio'], errors='coerce')
    df['Fecha_Cierre'] = pd.to_datetime(df['Fecha_Cierre'], errors='coerce')
    
    # Rellenar fechas vacías en casos abiertos con el ingreso
    df['Fecha_Ultimo_Cambio'] = df['Fecha_Ultimo_Cambio'].fillna(df['Fecha_Ingreso'])
    
    df['Es_Abierto'] = (df['Estado_Actual'] != 'Cerrado') & df['Fecha_Cierre'].isna()
    
    # Extraer periodos
    df['Mes_Cierre'] = df['Fecha_Cierre'].dt.to_period('M').astype(str)
    df['Trimestre_Cierre'] = df['Fecha_Cierre'].dt.to_period('Q').astype(str)
    df['Año_Cierre'] = df['Fecha_Cierre'].dt.year.astype(str).replace('nan', 'Pendiente')
    
    # Cálculos para casos Abiertos (WIP actual)
    df_abiertos = df[df['Es_Abierto']].copy()
    if not df_abiertos.empty:
        fechas_cambio = df_abiertos['Fecha_Ultimo_Cambio'].values.astype('datetime64[D]')
        fecha_hoy = np.datetime64('today')
        df_abiertos['Dias_En_Subestado'] = np.busday_count(fechas_cambio, fecha_hoy)
        
        condiciones = [
            (df_abiertos['Dias_En_Subestado'] <= 15),
            (df_abiertos['Dias_En_Subestado'] > 15) & (df_abiertos['Dias_En_Subestado'] <= 30),
            (df_abiertos['Dias_En_Subestado'] > 30)
        ]
        opciones = ['0-15 Días', '16-30 Días', '+30 Días (Crítico)']
        df_abiertos['Tramo_Aging'] = np.select(condiciones, opciones, default='Desconocido')
    
    # Cálculos para casos Cerrados
    df_cerrados = df[~df['Es_Abierto']].copy()
    if not df_cerrados.empty:
        fechas_in = df_cerrados['Fecha_Ingreso'].values.astype('datetime64[D]')
        fechas_out = df_cerrados['Fecha_Cierre'].values.astype('datetime64[D]')
        df_cerrados['Lead_Time_Total'] = np.busday_count(fechas_in, fechas_out)
        
    return df_abiertos, df_cerrados, df

df_abiertos, df_cerrados, df_master = procesar_datos_integrales(df_raw)

# --- FILTROS DE INTERFAZ TEMPORAL ---
st.sidebar.header("Filtros de Tendencias")
st.sidebar.markdown("*Aplica a la pestaña de Históricos.*")
tipo_periodo = st.sidebar.radio("Agrupación Temporal:", ["Mensual", "Trimestral", "Anual"])

if tipo_periodo == "Mensual":
    periodos_disp = sorted(df_cerrados['Mes_Cierre'].unique(), reverse=True) if not df_cerrados.empty else []
    col_cierre = 'Mes_Cierre'
elif tipo_periodo == "Trimestral":
    periodos_disp = sorted(df_cerrados['Trimestre_Cierre'].unique(), reverse=True) if not df_cerrados.empty else []
    col_cierre = 'Trimestre_Cierre'
else:
    periodos_disp = sorted(df_cerrados['Año_Cierre'].unique(), reverse=True) if not df_cerrados.empty else []
    col_cierre = 'Año_Cierre'

periodos_limpios = [p for p in periodos_disp if p != 'NaT' and p != 'Pendiente']
periodo_seleccionado = st.sidebar.selectbox("Seleccionar Periodo Final:", periodos_limpios) if periodos_limpios else None

# --- SECCIÓN 3: MOTOR DE REPORTE VISUAL (DASHBOARD) ---
st.title("📊 Panel de Gestión Integral")
st.markdown("Monitor de control de casos en curso (WIP) y análisis de tendencias históricas.")

tab_energia, tab_moviles, tab_tendencias = st.tabs(["⚡ WIP: Ingeniería y Energía", "🚜 WIP: Equipos Móviles", "📈 Tendencias e Históricos"])

def renderizar_panel_area(df_area_abiertos, area_nombre):
    if df_area_abiertos.empty:
        st.success(f"No hay casos activos en {area_nombre}.")
        return
    
    st.markdown(f"**Total Casos en Curso (WIP): {len(df_area_abiertos)}**")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown("#### Tasa de Envejecimiento (Aging)")
        aging_count = df_area_abiertos['Tramo_Aging'].value_counts().reset_index()
        aging_count.columns = ['Tramo', 'Cantidad']
        fig_aging = px.pie(aging_count, values='Cantidad', names='Tramo', hole=0.4,
                           color='Tramo', color_discrete_map={'0-15 Días': '#2ecc71', '16-30 Días': '#f1c40f', '+30 Días (Crítico)': '#e74c3c'})
        fig_aging.update_layout(margin=dict(t=0, b=0, l=0, r=0))
        st.plotly_chart(fig_aging, use_container_width=True)

    with col2:
        st.markdown("#### Carga por Liquidador (Backlog)")
        carga_liq = df_area_abiertos.groupby('Liquidador').size().reset_index(name='Casos')
        carga_liq = carga_liq.sort_values('Casos', ascending=True)
        fig_carga = px.bar(carga_liq, x='Casos', y='Liquidador', orientation='h', text='Casos', color_discrete_sequence=['#3498db'])
        fig_carga.update_layout(margin=dict(t=0, b=0, l=0, r=0))
        st.plotly_chart(fig_carga, use_container_width=True)

    st.markdown("#### Matriz Quirúrgica de Subestados (Cuellos de Botella)")
    residencia_subestado = df_area_abiertos.groupby(['Estado_Actual', 'Subestado_Actual']).agg(
        Casos_Detenidos=('ID_Caso', 'count'), Dias_Promedio_Estancado=('Dias_En_Subestado', 'mean')
    ).reset_index().sort_values('Dias_Promedio_Estancado', ascending=False)
    
    fig_matrix = px.scatter(residencia_subestado, x='Subestado_Actual', y='Dias_Promedio_Estancado', 
                            size='Casos_Detenidos', color='Estado_Actual', text='Casos_Detenidos',
                            size_max=40, labels={'Dias_Promedio_Estancado': 'Días Promedio Atrapado'})
    fig_matrix.update_traces(textposition='top center')
    st.plotly_chart(fig_matrix, use_container_width=True)

with tab_energia:
    # Ajustar el nombre del área según cómo venga en el Excel real (Sensible a mayúsculas)
    renderizar_panel_area(df_abiertos[df_abiertos['Area_Negocio'].str.contains('Ingeniería', case=False, na=False)], 'Ingeniería y Energía')

with tab_moviles:
    renderizar_panel_area(df_abiertos[df_abiertos['Area_Negocio'].str.contains('Móviles', case=False, na=False)], 'Equipos Móviles')

with tab_tendencias:
    st.subheader(f"Análisis Retrospectivo y Tendencias ({tipo_periodo})")
    if not df_cerrados.empty and periodo_seleccionado:
        df_cierre_periodo = df_cerrados[df_cerrados[col_cierre] == periodo_seleccionado]
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Total Resoluciones del Periodo", len(df_cierre_periodo))
        with c2:
            lt_energia = df_cierre_periodo[df_cierre_periodo['Area_Negocio'].str.contains('Ingeniería', case=False, na=False)]['Lead_Time_Total'].mean()
            st.metric("Lead Time Ingeniería (Días)", f"{lt_energia:.1f}" if pd.notna(lt_energia) else "N/A")
        with c3:
            lt_moviles = df_cierre_periodo[df_cierre_periodo['Area_Negocio'].str.contains('Móviles', case=False, na=False)]['Lead_Time_Total'].mean()
            st.metric("Lead Time Móviles (Días)", f"{lt_moviles:.1f}" if pd.notna(lt_moviles) else "N/A")
            
        st.divider()
        
        st.markdown("#### Tendencias a lo largo del tiempo")
        historico_agrupado = df_cerrados.groupby([col_cierre, 'Area_Negocio']).agg(
            Volumen=('ID_Caso', 'count'),
            Lead_Time=('Lead_Time_Total', 'mean')
        ).reset_index().sort_values(col_cierre)
        
        ultimos_periodos = sorted([p for p in historico_agrupado[col_cierre].unique() if p != 'NaT'])[-12:]
        historico_filtrado = historico_agrupado[historico_agrupado[col_cierre].isin(ultimos_periodos)]

        col_t1, col_t2 = st.columns(2)
        with col_t1:
            fig_vol = px.line(historico_filtrado, x=col_cierre, y='Volumen', color='Area_Negocio', markers=True,
                              title="Volumen de Resolución (Throughput) Histórico")
            st.plotly_chart(fig_vol, use_container_width=True)
            
        with col_t2:
            fig_lt_hist = px.line(historico_filtrado, x=col_cierre, y='Lead_Time', color='Area_Negocio', markers=True,
                                  title="Evolución del Lead Time Promedio")
            st.plotly_chart(fig_lt_hist, use_container_width=True)

# --- SECCIÓN 4: MOTOR DE REPORTES EXPORTABLES (EXCEL Y WORD) ---
st.divider()
st.subheader("📥 Generación de Reportes Formales")

@st.cache_data
def generar_excel_completo(df_master, df_abiertos):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_master.to_excel(writer, index=False, sheet_name='Base_Completa')
        df_abiertos.to_excel(writer, index=False, sheet_name='WIP_Abiertos')
        if not df_abiertos.empty:
            pivot_liq = pd.pivot_table(df_abiertos, values='ID_Caso', index=['Area_Negocio', 'Liquidador'], columns='Subestado_Actual', aggfunc='count', fill_value=0)
            pivot_liq.to_excel(writer, sheet_name='Matriz_Carga')
    return output.getvalue()

def generar_grafico_mpl(df, x_col, y_col, titulo, ylabel, color):
    plt.figure(figsize=(7, 3.5))
    plt.plot(df[x_col], df[y_col], marker='o', color=color, linewidth=2)
    plt.title(titulo, fontsize=10, fontweight='bold')
    plt.ylabel(ylabel, fontsize=9)
    plt.xticks(rotation=45, ha='right', fontsize=8)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=120)
    plt.close()
    img_stream.seek(0)
    return img_stream

def generar_word_reporte(df_abiertos, df_cerrados, periodo_sel, col_cierre):
    doc = Document()
    doc.add_heading('Reporte Ejecutivo de Operaciones e Ingeniería', 0)
    doc.add_paragraph(f'Periodo de Corte: {periodo_sel}')
    
    doc.add_heading('1. Estado Actual del Portafolio (WIP)', level=1)
    doc.add_paragraph(f'Total de casos en curso a la fecha: {len(df_abiertos)} casos.')
    
    if not df_abiertos.empty:
        aging = df_abiertos['Tramo_Aging'].value_counts()
        doc.add_paragraph(f"• Casos en Riesgo (+30 Días Estancados): {aging.get('+30 Días (Crítico)', 0)}")
    
    doc.add_heading(f'2. Cierres y Tendencias', level=1)
    if not df_cerrados.empty and periodo_sel:
        datos_periodo = df_cerrados[df_cerrados[col_cierre] == periodo_sel]
        doc.add_paragraph(f'Volumen resuelto en el periodo: {len(datos_periodo)} casos.')
        
        tendencia = df_cerrados.groupby(col_cierre)['Lead_Time_Total'].mean().reset_index().tail(6)
        if len(tendencia) > 1:
            img_trend = generar_grafico_mpl(tendencia, col_cierre, 'Lead_Time_Total', 'Evolución Lead Time (Últimos periodos)', 'Días Promedio', '#2980b9')
            doc.add_picture(img_trend, width=Inches(6.0))
            
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

col_d1, col_d2 = st.columns(2)

with col_d1:
    excel_data = generar_excel_completo(df_master, df_abiertos)
    st.download_button(label="📊 Descargar Matriz Operativa (Excel)", data=excel_data,
                       file_name=f"Matriz_Operativa_{datetime.now().strftime('%Y%m%d')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

with col_d2:
    if periodo_seleccionado:
        word_data = generar_word_reporte(df_abiertos, df_cerrados, periodo_seleccionado, col_cierre)
        st.download_button(label="📝 Generar Reporte de Gerencia (Word)", data=word_data,
                           file_name=f"Reporte_Gerencia_{periodo_seleccionado}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    else:
        st.info("Sube datos para habilitar el reporte en Word.")
        
