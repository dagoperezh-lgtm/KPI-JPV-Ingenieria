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
    np.random.seed(42)
    hoy = datetime(2026, 4, 1) 
    areas = ["Ingeniería y Energía", "Equipos Móviles"]
    liquidadores = ["Carlos Mendoza", "Ana Rojas", "Luis Silva", "Marta Pérez"]
    estados_principales = ["Ingreso", "Instrucción y Análisis", "Resolución", "Cerrado"]
    
    fechas_ingreso = [hoy - timedelta(days=int(x)) for x in np.random.randint(2, 1000, 1500)]
    datos = []
    
    for f in fechas_ingreso:
        area = np.random.choice(areas, p=[0.3, 0.7]) 
        liquidador = np.random.choice(liquidadores)
        dias_desde_ingreso = (hoy - f).days
        es_cerrado = np.random.random() < 0.8
        estado = "Cerrado" if es_cerrado else "En Análisis"
        
        datos.append({
            "ID_Caso": f"CASO-{np.random.randint(100000, 999999)}",
            "Area_Negocio": area,
            "Liquidador": liquidador,
            "Estado_Actual": estado,
            "Subestado_Actual": "Subestado de prueba",
            "Fecha_Ingreso": f,
            "Fecha_Cierre": f + timedelta(days=np.random.randint(5, 60)) if es_cerrado else pd.NaT,
            "Días desde asignación": np.random.randint(1, 100),
            "Días desde contacto": np.random.randint(1, 10),
            "Días entre inspección e asignación": np.random.randint(1, 15),
            "Días informe inicial - inspección": np.random.randint(1, 20),
            "Días análisis contractual - asignación": np.random.randint(5, 45),
            "Días informe final despachado - análisis contractual enviado": np.random.randint(1, 30),
            "Perdida bruta (en moneda del caso)": np.random.randint(1000, 50000),
            "Monto asegurado (en moneda del caso)": np.random.randint(50000, 1000000),
            "Gastos (UF)": np.random.uniform(5, 50),
            "Honorarios (UF)": np.random.uniform(10, 100)
        })
    return pd.DataFrame(datos)

def buscar_indice_columna(columnas, palabras_clave):
    for i, col in enumerate(columnas):
        if str(col).strip().lower() in palabras_clave:
            return i
    for i, col in enumerate(columnas):
        for palabra in palabras_clave:
            if palabra in str(col).strip().lower():
                return i
    return 0

st.sidebar.title("Configuración y Carga")
archivo_subido = st.sidebar.file_uploader("Cargar Reporte de Casos (CSV/Excel)", type=["csv", "xlsx"])

if archivo_subido is not None:
    try:
        filas_saltar = st.sidebar.number_input("Filas a saltar (Encabezado desfasado)", min_value=0, max_value=20, value=5)
        
        if archivo_subido.name.endswith('.csv'):
            df_crudo = pd.read_csv(archivo_subido, skiprows=filas_saltar, low_memory=False)
        else:
            xl = pd.ExcelFile(archivo_subido)
            hoja_seleccionada = st.sidebar.selectbox("Selecciona la pestaña de tu Excel", xl.sheet_names)
            df_crudo = pd.read_excel(archivo_subido, sheet_name=hoja_seleccionada, skiprows=filas_saltar)
            
        columnas_reales = df_crudo.columns.tolist()
        st.sidebar.success("¡Archivo detectado! Mapeo automático activado.")
        
        idx_id = buscar_indice_columna(columnas_reales, ['número de caso', 'numero de caso', 'id'])
        idx_area = buscar_indice_columna(columnas_reales, ['división', 'division', 'área de negocio'])
        idx_liq = buscar_indice_columna(columnas_reales, ['ajustador senior', 'liquidador'])
        idx_estado = buscar_indice_columna(columnas_reales, ['estado'])
        idx_subestado = buscar_indice_columna(columnas_reales, ['sub estado', 'subestado'])
        idx_in = buscar_indice_columna(columnas_reales, ['creado en', 'fecha de denuncio'])
        idx_out = buscar_indice_columna(columnas_reales, ['fecha de cierre', 'fecha cierre'])

        col_id = st.sidebar.selectbox("Columna ID Caso", columnas_reales, index=idx_id)
        col_area = st.sidebar.selectbox("Columna División", columnas_reales, index=idx_area)
        col_liq = st.sidebar.selectbox("Columna Ajustador Senior", columnas_reales, index=idx_liq)
        col_estado = st.sidebar.selectbox("Columna Estado", columnas_reales, index=idx_estado)
        col_subestado = st.sidebar.selectbox("Columna Sub estado", columnas_reales, index=idx_subestado)
        col_fecha_in = st.sidebar.selectbox("Columna Creado en", columnas_reales, index=idx_in)
        col_fecha_out = st.sidebar.selectbox("Columna Fecha Cierre", columnas_reales, index=idx_out)

        # Renombramos solo lo fundamental para el sistema, el resto pasa tal cual
        df_raw = df_crudo.rename(columns={
            col_id: "ID_Caso",
            col_area: "Area_Negocio",
            col_liq: "Liquidador",
            col_estado: "Estado_Actual",
            col_subestado: "Subestado_Actual",
            col_fecha_in: "Fecha_Ingreso",
            col_fecha_out: "Fecha_Cierre"
        })
        
    except Exception as e:
        st.sidebar.error(f"Error al procesar el archivo: {e}")
        df_raw = generar_datos_prueba()
else:
    st.sidebar.info("Usando datos de demostración para visualización.")
    df_raw = generar_datos_prueba()

# --- SECCIÓN 2: MOTOR DE CÁLCULO ESTRICTO ---
def procesar_datos_integrales(df):
    # BLINDAJE ANTI-KEYERROR: Asegura que las columnas existan aunque el Excel falle
    for col in ['Estado_Actual', 'Subestado_Actual', 'Area_Negocio', 'Liquidador']:
        if col not in df.columns:
            df[col] = 'Desconocido'
    for col in ['Fecha_Ingreso', 'Fecha_Cierre']:
        if col not in df.columns:
            df[col] = pd.NaT

    df['Estado_Actual'] = df['Estado_Actual'].fillna('Desconocido').astype(str).str.strip().str.upper()
    df['Subestado_Actual'] = df['Subestado_Actual'].fillna('Desconocido').astype(str).str.strip().str.upper()
    df['Area_Negocio'] = df['Area_Negocio'].fillna('Sin Área').astype(str).str.strip()
    
    # 1. FILTRO: Omitir casos Rechazados
    df = df[~df['Estado_Actual'].str.contains('RECHAZADO|RECHAZO', case=False, na=False)]
    df = df[~df['Subestado_Actual'].str.contains('RECHAZADO|RECHAZO', case=False, na=False)]
    
    # 2. DETECCIÓN Y FILTRO DE TIEMPOS DE RESIDENCIA (Columnas "Días")
    cols_dias = [col for col in df.columns if 'Días' in col or 'Dias' in col]
    for c in cols_dias:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        # FILTRO: Eliminar errores lógicos (outliers de > 1500 días)
        df = df[df[c] < 1500]

    # Transformación de Fechas Clave
    df['Fecha_Ingreso'] = pd.to_datetime(df['Fecha_Ingreso'], errors='coerce')
    df['Fecha_Cierre'] = pd.to_datetime(df['Fecha_Cierre'], errors='coerce')
    
    df['Es_Abierto'] = ~df['Estado_Actual'].str.contains('CERRADO') & df['Fecha_Cierre'].isna()
    
    df['Mes_Cierre'] = df['Fecha_Cierre'].dt.to_period('M').astype(str)
    df['Trimestre_Cierre'] = df['Fecha_Cierre'].dt.to_period('Q').astype(str)
    df['Año_Cierre'] = df['Fecha_Cierre'].dt.year.astype(str).replace('nan', 'Pendiente')
    
    df_abiertos = df[df['Es_Abierto']].copy()
    df_cerrados = df[~df['Es_Abierto']].copy()
            
    return df_abiertos, df_cerrados, df, cols_dias

df_abiertos, df_cerrados, df_master, columnas_de_dias = procesar_datos_integrales(df_raw)

# --- FILTROS DE INTERFAZ TEMPORAL ---
st.sidebar.header("Filtros de Tendencias")
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

periodos_limpios = [p for p in periodos_disp if p != 'NaT' and p != 'Pendiente' and p != 'nan']
periodo_seleccionado = st.sidebar.selectbox("Seleccionar Periodo Final:", periodos_limpios) if periodos_limpios else None

# --- SECCIÓN 3: MOTOR DE REPORTE VISUAL (DASHBOARD) ---
st.title("📊 Panel de Gestión: Tiempos de Residencia Reales")
st.markdown("Monitor de control basado exclusivamente en los registros de tiempo de tu sistema.")

tab_energia, tab_moviles, tab_tendencias = st.tabs(["⚡ WIP: Ingeniería y Energía", "🚜 WIP: Equipos Móviles", "📈 Cierres e Históricos"])

def renderizar_panel_area(df_area_abiertos, area_nombre, cols_dias):
    if df_area_abiertos.empty:
        st.success(f"No hay casos activos detectados para la división {area_nombre}.")
        return
    
    st.markdown(f"**Total Casos en Curso (WIP): {len(df_area_abiertos)}**")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown(f"#### Carga por Ajustador Senior ({area_nombre})")
        carga_liq = df_area_abiertos.groupby('Liquidador').size().reset_index(name='Casos Asignados')
        carga_liq = carga_liq.sort_values('Casos Asignados', ascending=True)
        fig_carga = px.bar(carga_liq, x='Casos Asignados', y='Liquidador', orientation='h', text='Casos Asignados', color_discrete_sequence=['#3498db'])
        st.plotly_chart(fig_carga, use_container_width=True)

    with col2:
        st.markdown("#### Tiempos Promedio de Residencia (Casos Activos)")
        st.markdown("Promedio de días registrados por el sistema en las etapas de vida del caso.")
        
        # Filtramos las columnas que realmente tienen datos para graficar los promedios
        promedios_dias = []
        for c in cols_dias:
            promedio = df_area_abiertos[c].mean()
            if promedio > 0:
                promedios_dias.append({'Etapa (Columna del Sistema)': c, 'Días Promedio': promedio})
                
        if promedios_dias:
            df_promedios = pd.DataFrame(promedios_dias).sort_values('Días Promedio', ascending=True)
            fig_dias = px.bar(df_promedios, x='Días Promedio', y='Etapa (Columna del Sistema)', orientation='h', text_auto='.1f', color_discrete_sequence=['#e74c3c'])
            st.plotly_chart(fig_dias, use_container_width=True)
        else:
            st.info("No hay datos de 'Días' registrados para los casos activos de esta división.")

with tab_energia:
    renderizar_panel_area(df_abiertos[df_abiertos['Area_Negocio'].str.contains('Ingeniería', case=False, na=False)], 'Ingeniería y Energía', columnas_de_dias)

with tab_moviles:
    renderizar_panel_area(df_abiertos[df_abiertos['Area_Negocio'].str.contains('Móvil|Movil', case=False, na=False)], 'Equipos Móviles', columnas_de_dias)

with tab_tendencias:
    st.subheader(f"Análisis Retrospectivo ({tipo_periodo})")
    if not df_cerrados.empty and periodo_seleccionado:
        df_cierre_periodo = df_cerrados[df_cerrados[col_cierre] == periodo_seleccionado]
        
        c1, c2 = st.columns(2)
        with c1:
            st.metric("Total Resoluciones del Periodo", len(df_cierre_periodo))
            vol_energia = len(df_cierre_periodo[df_cierre_periodo['Area_Negocio'].str.contains('Ingeniería', case=False, na=False)])
            st.markdown(f"**Ingeniería:** {vol_energia} casos")
        with c2:
            st.metric("Total Ajustadores Involucrados", df_cierre_periodo['Liquidador'].nunique())
            vol_moviles = len(df_cierre_periodo[df_cierre_periodo['Area_Negocio'].str.contains('Móvil|Movil', case=False, na=False)])
            st.markdown(f"**Móviles:** {vol_moviles} casos")
            
        st.divider()
        
        st.markdown("#### Tendencias Históricas de Cierre (Volumen)")
        historico_agrupado = df_cerrados.groupby([col_cierre, 'Area_Negocio']).agg(Volumen=('ID_Caso', 'count')).reset_index().sort_values(col_cierre)
        ultimos_periodos = sorted([p for p in historico_agrupado[col_cierre].unique() if p != 'NaT' and p != 'nan'])[-12:]
        historico_filtrado = historico_agrupado[historico_agrupado[col_cierre].isin(ultimos_periodos)]

        fig_vol = px.line(historico_filtrado, x=col_cierre, y='Volumen', color='Area_Negocio', markers=True)
        st.plotly_chart(fig_vol, use_container_width=True)
    else:
        st.info("No hay casos cerrados con fechas válidas para mostrar tendencias históricas.")

# --- SECCIÓN 4: MOTOR DE REPORTES EXPORTABLES (EXCEL Y WORD) ---
st.divider()
st.subheader("📥 Generación de Reportes Formales")

@st.cache_data
def generar_excel_completo(df_master, df_abiertos):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_master.to_excel(writer, index=False, sheet_name='Base_Filtrada_Limpia')
        df_abiertos.to_excel(writer, index=False, sheet_name='WIP_Abiertos')
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
    doc.add_paragraph(f'Periodo de Análisis y Corte: {periodo_sel}')
    
    doc.add_heading('1. Estado Actual del Portafolio (WIP)', level=1)
    doc.add_paragraph(f'Total de casos en curso a la fecha: {len(df_abiertos)} casos.')
    doc.add_paragraph('Nota: Este reporte excluye todos los casos clasificados como "Rechazados" y anomalías del sistema para asegurar la integridad de las métricas.')
    
    doc.add_heading(f'2. Cierres y Tendencias', level=1)
    if not df_cerrados.empty and periodo_sel:
        datos_periodo = df_cerrados[df_cerrados[col_cierre] == periodo_sel]
        doc.add_paragraph(f'Volumen resuelto en el periodo final: {len(datos_periodo)} casos.')
        
        tendencia = df_cerrados.groupby(col_cierre).size().reset_index(name='Volumen').tail(6)
        if len(tendencia) > 1:
            img_trend = generar_grafico_mpl(tendencia, col_cierre, 'Volumen', 'Evolución de Resoluciones (Últimos periodos)', 'Cantidad de Casos', '#27ae60')
            doc.add_picture(img_trend, width=Inches(6.0))
            
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

col_d1, col_d2 = st.columns(2)

with col_d1:
    excel_data = generar_excel_completo(df_master, df_abiertos)
    st.download_button(label="📊 Descargar Base Limpia (Excel)", data=excel_data,
                       file_name=f"Base_Datos_Limpia_{datetime.now().strftime('%Y%m%d')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

with col_d2:
    if periodo_seleccionado and not df_cerrados.empty:
        word_data = generar_word_reporte(df_abiertos, df_cerrados, periodo_seleccionado, col_cierre)
        st.download_button(label="📝 Generar Reporte de Gerencia (Word)", data=word_data,
                           file_name=f"Reporte_Gerencia_{periodo_seleccionado}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    else:
        st.info("No hay información suficiente para generar el reporte Word.")
