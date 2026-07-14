"""
Tablero Gerencial — División Ingeniería y Equipo Móvil.

App de solo lectura que reemplaza el llenado manual semanal del Excel
TABLERO_ING: lee la Base Maestra y los planes semanales que el equipo ya
carga en OpsControl (mismo Google Sheet) y arma automáticamente la misma
grilla (Stock, Ajuste, IFL, Top 5, Gestión Comercial) por División /
Ajustador / Tramo UF.

Lo único que se ingresa manualmente aquí son las Metas (cuota semanal de
entregables y meta anual/mensual de producción), porque esos números no
existen en ningún caso del sistema.
"""
import io
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

import tablero_calculo as calc
import tablero_datos as datos
import tablero_metas as metas_mod

st.set_page_config(page_title="Tablero Gerencial — Ingeniería y Equipo Móvil", layout="wide", page_icon="🛠️")


# ---------------------------------------------------------
# DATOS DE DEMOSTRACIÓN (cuando no hay conexión a Google Sheets)
# ---------------------------------------------------------
def _generar_demo():
    np.random.seed(7)
    ajustadores = {
        "Francisco Silva": "INGENIERÍA Y ENERGÍA",
        "Camilo Abarzúa": "INGENIERÍA Y ENERGÍA",
        "Nélson Canio": "INGENIERÍA Y ENERGÍA",
        "Mauricio Muñoz": "EQUIPO MOVIL",
        "Leopoldo Soto": "EQUIPO MOVIL",
    }
    companias = ["Mapfre", "Southbridge", "Suramericana", "Chubb", "Consorcio", "HDI"]
    corredoras = ["Marsh S.A.", "Aon Risk Services", "Ureta y Fernández", "BancoEstado Corredores"]

    filas = []
    hoy = datetime.now()
    for i in range(90):
        aj = np.random.choice(list(ajustadores.keys()))
        fila = {
            "Número de caso": f"{200000 + i}",
            "Nickname": f"Caso {i}",
            "División": ajustadores[aj],
            "Compañía de seguros": np.random.choice(companias),
            "Corredora": np.random.choice(corredoras),
            "Ajustador senior": aj,
            "Asegurado": f"Asegurado {i}",
            "Estado": np.random.choice(["Ajuste", "IFL", "Cerrado"], p=[0.5, 0.3, 0.2]),
            "Sub estado": "En Proceso",
            "Creado en": (hoy - pd.Timedelta(days=int(np.random.randint(1, 400)))).strftime("%Y-%m-%d"),
            "Divisa": "UF",
            "Perdida bruta (en moneda del caso)": int(np.random.choice([500, 2500, 8000])),
        }
        for extra in range(53):
            fila[f"col_{extra}"] = ""
        fila["Honorarios"] = round(np.random.uniform(5, 120), 2)
        filas.append(fila)
    df_maestro = pd.DataFrame(filas)

    tareas = []
    acciones_ifl = ["Preparar Informe - Informe Final de Liquidación", "Preparar Informe - Carta de Cobertura (Rechazo)"]
    acciones_ajuste = ["Preparar Informe - Informe Intermedio", "Preparar Informe - Preliminar Extendido", "En Ajuste - Revisión de cobertura"]
    for _, row in df_maestro[df_maestro["Estado"] != "Cerrado"].sample(40, random_state=3).iterrows():
        es_ifl = np.random.random() < 0.3
        tareas.append({
            "id_transaccion": f"demo-{row['Número de caso']}",
            "tipo_actividad": np.random.choice(["Programada", "Actividad Adicional"], p=[0.8, 0.2]),
            "categoria": "Operativa",
            "numero_caso": row["Número de caso"],
            "asegurado": row["Asegurado"],
            "tramo_uf": "<= 1000 UF",
            "honorarios_estimados": row["Honorarios"],
            "accion": np.random.choice(acciones_ifl if es_ifl else acciones_ajuste),
            "fecha_compromiso": hoy.strftime("%Y-%m-%d"),
            "estado_cumplimiento": np.random.choice(["Realizado", "Pendiente"], p=[0.6, 0.4]),
            "fecha_ejecucion": hoy.strftime("%Y-%m-%d %H:%M:%S"),
            "Ajustador": row["Ajustador senior"],
        })
    for aj in ajustadores:
        tareas.append({
            "id_transaccion": f"demo-com-{aj}", "tipo_actividad": "Programada", "categoria": "Gestión Comercial",
            "numero_caso": "N/A", "asegurado": "N/A", "tramo_uf": "N/A", "honorarios_estimados": 0.0,
            "accion": "Reunión de seguimiento con corredora", "fecha_compromiso": hoy.strftime("%Y-%m-%d"),
            "estado_cumplimiento": "Realizado", "fecha_ejecucion": hoy.strftime("%Y-%m-%d %H:%M:%S"), "Ajustador": aj,
        })
    df_plan = pd.DataFrame(tareas)
    return df_maestro, df_plan


# ---------------------------------------------------------
# SIDEBAR: ESTADO DE CONEXIÓN Y SELECCIÓN DE SEMANA
# ---------------------------------------------------------
st.sidebar.title("🛠️ Tablero Gerencial")
st.sidebar.caption("Fuente de datos: OpsControl (Base Maestra + Planes Semanales)")

modo_demo = False
df_maestro, fecha_actualizacion = datos.load_base_maestra()
if df_maestro is None:
    st.sidebar.warning("⚠️ Sin conexión a Google Sheets o Base Maestra vacía. Mostrando datos de demostración.")
    modo_demo = True
else:
    st.sidebar.success(f"✅ Base Maestra actualizada: {fecha_actualizacion or 'sin fecha'} ({len(df_maestro)} casos)")

st.sidebar.markdown("---")
offset_semana = st.sidebar.selectbox(
    "Semana a reportar:",
    options=[0, -1, -2, -3],
    format_func=lambda o: {0: "Semana Actual", -1: "Semana Anterior", -2: "Hace 2 semanas", -3: "Hace 3 semanas"}[o],
)
hoy = datetime.now()
dias_actual = datos.dias_habiles_semana(offset_semana, hoy)
dias_previa = datos.dias_habiles_semana(offset_semana - 1, hoy)
week_id_actual = datos.get_week_identifier(offset_semana, hoy)
week_id_previa = datos.get_week_identifier(offset_semana - 1, hoy)

if modo_demo:
    df_maestro, df_plan_actual = _generar_demo()
    df_plan_previa = df_plan_actual.sample(frac=0.7, random_state=1)
else:
    df_plan_actual = datos.load_planes_semana(week_id_actual)
    df_plan_previa = datos.load_planes_semana(week_id_previa)

metas = metas_mod.load_metas()
metas_semanales_dict = metas_mod.metas_semanales_como_dict(metas)

st.title("🛠️ Tablero Gerencial — Ingeniería y Equipo Móvil")
st.caption(f"Semana actual: {datos.rango_semana_legible(offset_semana, hoy)}  ·  Semana anterior: {datos.rango_semana_legible(offset_semana - 1, hoy)}")

tab_grilla, tab_avance, tab_top5, tab_comercial, tab_metas = st.tabs([
    "📋 Grilla Semanal", "📈 Avance Producción / Meta", "🏆 Top 5", "🤝 Gestión Comercial", "🎯 Configurar Metas",
])


# ---------------------------------------------------------
# TAB 1: GRILLA SEMANAL (Stock / Ajuste / IFL) — Semana Anterior vs Actual
# ---------------------------------------------------------
def _formatear_grilla(grilla):
    if grilla.empty:
        return grilla
    df = grilla.copy()
    for c in ["Stock_Hon_UF", "Ajuste_HonpUF", "Ajuste_HonrUF", "IFL_HonpUF", "IFL_HonrUF"]:
        df[c] = df[c].round(2)
    for c in ["Asignados", "Stock_Q", "Ajuste_Qp", "Ajuste_Qr", "IFL_Qp", "IFL_Qr"]:
        df[c] = df[c].astype(int)
    df["Ajuste_Meta"] = df["Ajuste_Meta"].apply(lambda v: "-" if pd.isna(v) else int(v))
    df["IFL_Meta"] = df["IFL_Meta"].apply(lambda v: "-" if pd.isna(v) else int(v))
    return df.rename(columns={
        "Division": "División", "Ajustador": "Ajustador", "Tramo": "Tramo",
        "Asignados": "Asignados", "Stock_Q": "Q", "Dias_prom": "Días prom", "Stock_Hon_UF": "Hon UF",
        "Ajuste_Qp": "Ajuste: Q p", "Ajuste_Meta": "Ajuste: Meta", "Ajuste_HonpUF": "Ajuste: Hon p UF",
        "Ajuste_Qr": "Ajuste: Q r", "Ajuste_HonrUF": "Ajuste: Hon r UF",
        "IFL_Qp": "IFL: Q p", "IFL_Meta": "IFL: Meta", "IFL_HonpUF": "IFL: Hon p UF",
        "IFL_Qr": "IFL: Q r", "IFL_HonrUF": "IFL: Hon r UF",
    })


with tab_grilla:
    for etiqueta, dias, df_plan in [
        ("Semana Anterior", dias_previa, df_plan_previa),
        (f"Semana Actual ({datos.rango_semana_legible(offset_semana, hoy)})", dias_actual, df_plan_actual),
    ]:
        st.subheader(etiqueta)
        grilla = calc.construir_grilla(df_maestro, df_plan, dias, metas_semanales_dict)
        if grilla.empty:
            st.info("Sin casos ni planes registrados para este período.")
            continue
        for division in sorted(grilla["Division"].unique()):
            with st.expander(f"**{division}**", expanded=True):
                sub = _formatear_grilla(grilla[grilla["Division"] == division].drop(columns=["Division"]))
                st.dataframe(sub, use_container_width=True, hide_index=True)
                c1, c2, c3 = st.columns(3)
                c1.metric("Stock total (Q)", int(sub["Q"].sum()))
                c2.metric("Ajuste realizado (Hon UF)", f"{sub['Ajuste: Hon r UF'].sum():,.2f}")
                c3.metric("IFL realizado (Hon UF)", f"{sub['IFL: Hon r UF'].sum():,.2f}")
        st.markdown("---")

    grilla_export = calc.construir_grilla(df_maestro, df_plan_actual, dias_actual, metas_semanales_dict)
    if not grilla_export.empty:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            _formatear_grilla(calc.construir_grilla(df_maestro, df_plan_previa, dias_previa, metas_semanales_dict)).to_excel(writer, index=False, sheet_name="Semana Anterior")
            _formatear_grilla(grilla_export).to_excel(writer, index=False, sheet_name="Semana Actual")
        st.download_button(
            "📥 Descargar Tablero (Excel)", data=buffer.getvalue(),
            file_name=f"Tablero_{week_id_actual}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ---------------------------------------------------------
# TAB 2: AVANCE PRODUCCIÓN / META
# ---------------------------------------------------------
with tab_avance:
    st.subheader("Avance Producción vs Meta")
    st.caption("Producción = honorarios de casos con Informe Final de Liquidación / Carta de Cobertura (Rechazo) marcados como Realizado, acumulado del año en curso.")

    if modo_demo:
        df_tareas_anio = df_plan_actual
    else:
        df_tareas_anio = datos.load_todas_las_tareas_realizadas()

    avance = calc.avance_produccion(df_maestro, df_tareas_anio, hoy)
    metas_anuales_df = metas_mod.metas_anuales_a_df(metas)

    if avance.empty:
        st.info("Aún no hay producción (IFL realizados) registrada este año.")
    else:
        combinado = avance.merge(metas_anuales_df.rename(columns={"division": "Division"}), on="Division", how="left")
        combinado["meta_anual_uf"] = pd.to_numeric(combinado.get("meta_anual_uf"), errors="coerce")
        combinado["meta_mensual_uf"] = pd.to_numeric(combinado.get("meta_mensual_uf"), errors="coerce")
        combinado["% Cump Anual"] = (combinado["Produccion_Acumulada_UF"] / combinado["meta_anual_uf"] * 100).round(1)
        combinado["% Cump Mes"] = (combinado["Produccion_Mes_UF"] / combinado["meta_mensual_uf"] * 100).round(1)

        for _, fila in combinado.iterrows():
            st.markdown(f"#### {fila['Division']}")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Producción Acumulada (UF)", f"{fila['Produccion_Acumulada_UF']:,.2f}")
            c2.metric("Meta Anual (UF)", f"{fila['meta_anual_uf']:,.0f}" if pd.notna(fila["meta_anual_uf"]) else "Sin definir")
            c3.metric("Producción Mes (UF)", f"{fila['Produccion_Mes_UF']:,.2f}")
            c4.metric("Meta Mes (UF)", f"{fila['meta_mensual_uf']:,.0f}" if pd.notna(fila["meta_mensual_uf"]) else "Sin definir")
            if pd.notna(fila["% Cump Anual"]):
                st.progress(min(int(fila["% Cump Anual"]), 100), text=f"{fila['% Cump Anual']:.1f}% de la meta anual")
        st.caption("Si falta la Meta Anual/Mensual de una división, defínela en la pestaña 'Configurar Metas'.")


# ---------------------------------------------------------
# TAB 3: TOP 5
# ---------------------------------------------------------
with tab_top5:
    st.subheader("Top 5 por División")
    col_ing, col_mov = st.columns(2)
    for columna, filtro, titulo in [(col_ing, "Ingenier", "Ingeniería y Energía"), (col_mov, "Móvil|Movil", "Equipo Móvil")]:
        with columna:
            st.markdown(f"##### {titulo}")
            top_comp, top_corr = calc.calcular_top5(df_maestro, filtro)
            st.markdown("**Compañías de seguros**")
            st.dataframe(top_comp.rename_axis("Compañía").reset_index(name="Casos vigentes"), use_container_width=True, hide_index=True)
            st.markdown("**Corredoras**")
            st.dataframe(top_corr.rename_axis("Corredora").reset_index(name="Casos vigentes"), use_container_width=True, hide_index=True)


# ---------------------------------------------------------
# TAB 4: GESTIÓN COMERCIAL
# ---------------------------------------------------------
with tab_comercial:
    st.subheader("Bitácora de Gestión Comercial — Semana Actual")
    bitacora = calc.bitacora_gestion_comercial(df_plan_actual)
    if bitacora.empty:
        st.info("Sin gestiones comerciales registradas esta semana.")
    else:
        st.dataframe(bitacora, use_container_width=True, hide_index=True)


# ---------------------------------------------------------
# TAB 5: CONFIGURACIÓN DE METAS (manual)
# ---------------------------------------------------------
with tab_metas:
    st.subheader("Meta semanal por Ajustador y Tramo (cuota de entregables)")
    st.caption("Cantidad de entregables Ajuste / IFL que se espera que cada ajustador cumpla por semana en cada tramo. No se calcula de ningún caso: es un objetivo que define la gerencia.")
    df_metas_sem = metas_mod.metas_semanales_a_df(metas)
    metas_sem_editado = st.data_editor(
        df_metas_sem, num_rows="dynamic", use_container_width=True, key="editor_metas_semanales",
        column_config={
            "ajustador": st.column_config.TextColumn("Ajustador"),
            "tramo": st.column_config.SelectboxColumn("Tramo", options=["<= 1000 UF", "> 1000 Y <= 5000 UF", "> 5000 UF (MCL)", "<= 200.000 USD", "> 200.000 USD (MCL)"]),
            "tipo": st.column_config.SelectboxColumn("Tipo", options=["Ajuste", "IFL"]),
            "valor": st.column_config.NumberColumn("Meta (Q)", min_value=0, step=1),
        },
    )

    st.markdown("---")
    st.subheader("Meta Anual y Mensual de Producción por División (UF)")
    df_metas_anio = metas_mod.metas_anuales_a_df(metas)
    metas_anio_editado = st.data_editor(
        df_metas_anio, num_rows="dynamic", use_container_width=True, key="editor_metas_anuales",
        column_config={
            "division": st.column_config.TextColumn("División"),
            "meta_anual_uf": st.column_config.NumberColumn("Meta Anual (UF)", min_value=0),
            "meta_mensual_uf": st.column_config.NumberColumn("Meta Mensual (UF)", min_value=0),
        },
    )

    if st.button("💾 Guardar Metas", type="primary"):
        nuevas_metas = {
            "metas_semanales": metas_mod.metas_semanales_desde_df(metas_sem_editado),
            "metas_anuales": metas_mod.metas_anuales_desde_df(metas_anio_editado),
        }
        metas_mod.save_metas(nuevas_metas)
        st.success("✅ Metas guardadas correctamente.")
        st.rerun()
