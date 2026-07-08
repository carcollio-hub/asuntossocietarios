# -*- coding: utf-8 -*-
"""
Dashboard SCD - Gestiones (Boletines, Anticipos, Aporte Viaje, Pagos)
Fuente de datos: archivo Excel con una pestaña por gestión.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ============================================================
# CONFIGURACIÓN DE PÁGINA
# ============================================================
st.set_page_config(
    page_title="Dashboard SCD | Gestiones",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# PALETA DE COLORES Y ESTILO
# ============================================================
COLOR_PRIMARY = "#1B3B6F"     # azul profundo
COLOR_SECONDARY = "#2A9D8F"   # verde azulado
COLOR_ACCENT = "#E76F51"      # coral
COLOR_GOLD = "#E9C46A"        # dorado
COLOR_SAGE = "#8AB17D"        # verde salvia
COLOR_PLUM = "#6D597A"        # ciruela

PALETTE = [COLOR_PRIMARY, COLOR_SECONDARY, COLOR_ACCENT, COLOR_GOLD, COLOR_SAGE, COLOR_PLUM]
COLOR_MAP_SEXO = {"Hombre": COLOR_PRIMARY, "Mujer": COLOR_ACCENT}

MESES_ORDEN = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
               "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

RANGO_EDAD_ORDEN = ["0 a 11", "12 a 18", "19 a 30", "31 a 40", "41 a 50", "51 a 60", "61 a 100"]

PLOTLY_TEMPLATE = "plotly_white"

st.markdown(f"""
<style>
    .main .block-container {{
        padding-top: 1.5rem;
        max-width: 1300px;
    }}
    h1, h2, h3 {{
        color: {COLOR_PRIMARY};
    }}
    div[data-testid="stMetric"] {{
        background-color: #F7F9FB;
        border: 1px solid #E3E8EE;
        border-left: 5px solid {COLOR_SECONDARY};
        border-radius: 10px;
        padding: 14px 16px 8px 16px;
    }}
    div[data-testid="stMetricLabel"] {{
        font-weight: 600;
        color: #4A5568;
    }}
    section[data-testid="stSidebar"] {{
        background-color: #F7F9FB;
    }}
    .stTabs [data-baseweb="tab-list"] {{
        gap: 4px;
    }}
    .stTabs [data-baseweb="tab"] {{
        background-color: #F0F2F6;
        border-radius: 8px 8px 0 0;
        padding: 8px 16px;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {COLOR_PRIMARY};
        color: white;
    }}
    .seccion-titulo {{
        color: {COLOR_PRIMARY};
        border-bottom: 2px solid {COLOR_SECONDARY};
        padding-bottom: 6px;
        margin-top: 10px;
    }}
</style>
""", unsafe_allow_html=True)


# ============================================================
# CARGA Y LIMPIEZA DE DATOS
# ============================================================
def _limpiar_nombres_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.replace("\xa0", " ").strip() for c in df.columns]
    return df


def _normalizar_fecha(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        if col.lower().startswith("fecha") and col != "Fecha":
            df = df.rename(columns={col: "Fecha"})
            break
    return df


def _normalizar_tipo_conexo(valor):
    if pd.isna(valor):
        return valor
    v = str(valor).strip().upper()
    if "AGRUP" in v:
        return "Agrupación"
    if "SOLIST" in v:
        return "Solista"
    return str(valor).strip()


def _normalizar_tipo_pago(valor):
    if pd.isna(valor):
        return valor
    v = str(valor).strip()
    key = v.lower()
    mapping = {
        "aporte viaje": "Aporte Viaje",
        "pago derechos": "Pago Derechos",
    }
    return mapping.get(key, v)


@st.cache_data(show_spinner="Cargando datos...")
def cargar_datos(file) -> dict:
    xl = pd.ExcelFile(file)
    hojas = {}

    for nombre in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=nombre)
        df = _limpiar_nombres_columnas(df)
        df = _normalizar_fecha(df)

        if "ID Python" in df.columns:
            df = df.drop(columns=["ID Python"])

        if "Mes" in df.columns:
            df["Mes"] = pd.Categorical(df["Mes"], categories=MESES_ORDEN, ordered=True)

        if "Rango edad" in df.columns:
            cats = [c for c in RANGO_EDAD_ORDEN if c in df["Rango edad"].unique()]
            df["Rango edad"] = pd.Categorical(df["Rango edad"], categories=cats, ordered=True)

        hojas[nombre] = df

    # --- Limpieza específica: Boletines Conexos ---
    if "Boletines Conexos" in hojas:
        df = hojas["Boletines Conexos"]
        if "Tipo" in df.columns:
            df["Tipo"] = df["Tipo"].apply(_normalizar_tipo_conexo)
        hojas["Boletines Conexos"] = df

    # --- Limpieza específica: Aporte viaje ---
    if "Aporte viaje" in hojas:
        df = hojas["Aporte viaje"]
        if "Categ" in df.columns:
            df = df.rename(columns={"Categ": "Categoría"})
        hojas["Aporte viaje"] = df

    # --- Limpieza específica: Pagos ---
    if "Pagos" in hojas:
        df = hojas["Pagos"]
        if "Tipo de Pago" in df.columns:
            df["Tipo de Pago"] = df["Tipo de Pago"].apply(_normalizar_tipo_pago)
        if "Categoría" in df.columns:
            df["Categoría"] = df["Categoría"].replace({"Admin. Inactivo": "Administrado Inactivo"})
        if "Monto" in df.columns:
            df["Monto"] = pd.to_numeric(df["Monto"], errors="coerce")
        hojas["Pagos"] = df

    return hojas


# ============================================================
# FILTROS GLOBALES
# ============================================================
def obtener_opciones_filtro(hojas: dict, columna: str) -> list:
    valores = set()
    for df in hojas.values():
        if columna in df.columns:
            valores.update([v for v in df[columna].dropna().unique().tolist()])
    return sorted(valores, key=lambda x: str(x))


def aplicar_filtros(df: pd.DataFrame, sexo, region, edad, anio) -> pd.DataFrame:
    df_f = df.copy()
    if sexo and "Sexo" in df_f.columns:
        df_f = df_f[df_f["Sexo"].isin(sexo)]
    if region and "Región" in df_f.columns:
        df_f = df_f[df_f["Región"].isin(region)]
    if edad and "Rango edad" in df_f.columns:
        df_f = df_f[df_f["Rango edad"].astype(str).isin(edad)]
    if anio and "Año" in df_f.columns:
        df_f = df_f[df_f["Año"].isin(anio)]
    return df_f


# ============================================================
# FUNCIONES AUXILIARES DE VISUALIZACIÓN
# ============================================================
def titulo_seccion(texto: str):
    st.markdown(f"<h3 class='seccion-titulo'>{texto}</h3>", unsafe_allow_html=True)


def grafico_evolucion_mensual(df: pd.DataFrame, titulo: str, color=COLOR_PRIMARY):
    if df.empty or "Año" not in df.columns or "Mes" not in df.columns:
        st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")
        return
    conteo = df.groupby(["Año", "Mes"], observed=True).size().reset_index(name="Cantidad")
    conteo["Periodo"] = conteo["Año"].astype(str) + " - " + conteo["Mes"].astype(str)
    conteo = conteo.sort_values(["Año", "Mes"])
    fig = px.line(conteo, x="Periodo", y="Cantidad", markers=True, template=PLOTLY_TEMPLATE,
                  color_discrete_sequence=[color])
    fig.update_layout(title=titulo, xaxis_title="Periodo", yaxis_title="N° de registros",
                       height=380, xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)


def grafico_barras_categoria(df: pd.DataFrame, columna: str, titulo: str, top_n=None, horizontal=False):
    if df.empty or columna not in df.columns:
        st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")
        return
    conteo = df[columna].value_counts(dropna=True).reset_index()
    conteo.columns = [columna, "Cantidad"]
    if top_n:
        conteo = conteo.head(top_n)
    orient = "h" if horizontal else "v"
    x_col, y_col = ("Cantidad", columna) if horizontal else (columna, "Cantidad")
    fig = px.bar(conteo, x=x_col, y=y_col, orientation=orient, template=PLOTLY_TEMPLATE,
                 color=columna, color_discrete_sequence=PALETTE, text="Cantidad")
    fig.update_traces(textposition="outside", showlegend=False)
    fig.update_layout(title=titulo, height=380)
    if horizontal:
        fig.update_layout(yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig, use_container_width=True)


def grafico_torta(df: pd.DataFrame, columna: str, titulo: str, color_map=None):
    if df.empty or columna not in df.columns:
        st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")
        return
    conteo = df[columna].value_counts(dropna=True).reset_index()
    conteo.columns = [columna, "Cantidad"]
    fig = px.pie(conteo, names=columna, values="Cantidad", template=PLOTLY_TEMPLATE, hole=0.45,
                 color=columna, color_discrete_map=color_map,
                 color_discrete_sequence=PALETTE)
    fig.update_traces(textinfo="percent+label")
    fig.update_layout(title=titulo, height=380)
    st.plotly_chart(fig, use_container_width=True)


def seccion_demografia(df: pd.DataFrame, key_prefix: str):
    titulo_seccion("Perfil demográfico")
    c1, c2, c3 = st.columns(3)
    with c1:
        grafico_torta(df, "Sexo", "Distribución por sexo", color_map=COLOR_MAP_SEXO)
    with c2:
        if not df.empty and "Rango edad" in df.columns:
            conteo = df["Rango edad"].value_counts(dropna=True).reindex(RANGO_EDAD_ORDEN).dropna().reset_index()
            conteo.columns = ["Rango edad", "Cantidad"]
            fig = px.bar(conteo, x="Rango edad", y="Cantidad", template=PLOTLY_TEMPLATE,
                         color_discrete_sequence=[COLOR_SECONDARY], text="Cantidad")
            fig.update_traces(textposition="outside")
            fig.update_layout(title="Distribución por rango de edad", height=380)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")
    with c3:
        grafico_barras_categoria(df, "Región", "Top regiones", top_n=8, horizontal=True)


# ============================================================
# ENCABEZADO
# ============================================================
st.title("📊 Dashboard SCD — Gestiones de Socios")
st.caption("Boletines Autorales · Boletines Conexos · Anticipos Reajustables · Aporte Viaje · Pagos")

DATA_PATH = "Datos_Python.xlsx"

archivo_subido = None
try:
    hojas = cargar_datos(DATA_PATH)
except FileNotFoundError:
    st.sidebar.warning("No se encontró 'Datos_Python.xlsx' en la carpeta del proyecto.")
    archivo_subido = st.sidebar.file_uploader("Sube el archivo Excel", type=["xlsx"])
    if archivo_subido is None:
        st.info("⬅️ Sube el archivo Excel en la barra lateral para comenzar.")
        st.stop()
    hojas = cargar_datos(archivo_subido)

# ============================================================
# SIDEBAR — FILTROS GLOBALES
# ============================================================
st.sidebar.header("🔎 Filtros")

opciones_sexo = obtener_opciones_filtro(hojas, "Sexo")
opciones_region = obtener_opciones_filtro(hojas, "Región")
opciones_edad = [e for e in RANGO_EDAD_ORDEN if e in obtener_opciones_filtro(hojas, "Rango edad")]
opciones_anio = obtener_opciones_filtro(hojas, "Año")

filtro_sexo = st.sidebar.multiselect("Sexo", opciones_sexo, default=[])
filtro_region = st.sidebar.multiselect("Región", opciones_region, default=[])
filtro_edad = st.sidebar.multiselect("Rango edad", opciones_edad, default=[])
filtro_anio = st.sidebar.multiselect("Año", opciones_anio, default=[])

st.sidebar.caption("Si no seleccionas ninguna opción en un filtro, se consideran todos los valores.")
st.sidebar.divider()
st.sidebar.caption(f"Última actualización de datos: {pd.Timestamp.today().strftime('%d-%m-%Y')}")

# Aplicar filtros a cada hoja
hojas_filtradas = {
    nombre: aplicar_filtros(df, filtro_sexo, filtro_region, filtro_edad, filtro_anio)
    for nombre, df in hojas.items()
}

# ============================================================
# TABS PRINCIPALES
# ============================================================
tab_resumen, tab_autorales, tab_conexos, tab_anticipos, tab_viaje, tab_pagos = st.tabs([
    "🏠 Resumen General",
    "📝 Boletines Autorales",
    "🎵 Boletines Conexos",
    "💰 Anticipos Reajustables",
    "✈️ Aporte Viaje",
    "💳 Pagos",
])

# ------------------------------------------------------------
# TAB: RESUMEN GENERAL
# ------------------------------------------------------------
with tab_resumen:
    titulo_seccion("Volumen de gestiones")
    cols = st.columns(5)
    nombres_kpi = ["Boletines Autorales", "Boletines Conexos", "Anticipos Reajustables", "Aporte viaje", "Pagos"]
    for c, nombre in zip(cols, nombres_kpi):
        with c:
            st.metric(nombre, f"{len(hojas_filtradas[nombre]):,}".replace(",", "."))

    st.divider()
    titulo_seccion("Registros por gestión")
    resumen = pd.DataFrame({
        "Gestión": nombres_kpi,
        "Registros": [len(hojas_filtradas[n]) for n in nombres_kpi],
    })
    fig = px.bar(resumen, x="Gestión", y="Registros", template=PLOTLY_TEMPLATE,
                 color="Gestión", color_discrete_sequence=PALETTE, text="Registros")
    fig.update_traces(textposition="outside", showlegend=False)
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)

    st.divider()
    titulo_seccion("Evolución mensual combinada")
    frames = []
    for nombre in nombres_kpi:
        df = hojas_filtradas[nombre]
        if not df.empty and "Año" in df.columns and "Mes" in df.columns:
            conteo = df.groupby(["Año", "Mes"], observed=True).size().reset_index(name="Cantidad")
            conteo["Gestión"] = nombre
            frames.append(conteo)
    if frames:
        combinado = pd.concat(frames, ignore_index=True)
        combinado["Periodo"] = combinado["Año"].astype(str) + " - " + combinado["Mes"].astype(str)
        combinado = combinado.sort_values(["Año", "Mes"])
        fig = px.line(combinado, x="Periodo", y="Cantidad", color="Gestión", markers=True,
                      template=PLOTLY_TEMPLATE, color_discrete_sequence=PALETTE)
        fig.update_layout(height=420, xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")

    st.divider()
    # Perfil demográfico combinado (todas las gestiones concatenadas)
    demografia_frames = []
    for nombre in nombres_kpi:
        df = hojas_filtradas[nombre][["Sexo", "Región", "Rango edad"]].copy() \
            if all(c in hojas_filtradas[nombre].columns for c in ["Sexo", "Región", "Rango edad"]) \
            else pd.DataFrame(columns=["Sexo", "Región", "Rango edad"])
        demografia_frames.append(df)
    demografia_general = pd.concat(demografia_frames, ignore_index=True) if demografia_frames else pd.DataFrame()
    seccion_demografia(demografia_general, "resumen")

# ------------------------------------------------------------
# TAB: BOLETINES AUTORALES
# ------------------------------------------------------------
with tab_autorales:
    df = hojas_filtradas["Boletines Autorales"]

    titulo_seccion("Indicadores clave")
    c1, c2, c3, c4 = st.columns(4)
    total = len(df)
    pct_ingresado = (df["Estado Declaración"].eq("Obra ingresada").mean() * 100) if total else 0
    pct_transcripcion = (df["Solicita Transcripción"].eq("Si").mean() * 100) if total else 0
    pct_ia = (df["Uso IA"].eq("Si").mean() * 100) if total else 0
    c1.metric("Total declaraciones", f"{total:,}".replace(",", "."))
    c2.metric("Obra ingresada", f"{pct_ingresado:.1f}%")
    c3.metric("Solicita transcripción", f"{pct_transcripcion:.1f}%")
    c4.metric("Declara uso de IA", f"{pct_ia:.1f}%")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        grafico_evolucion_mensual(df, "Evolución mensual de declaraciones", color=COLOR_PRIMARY)
    with col2:
        grafico_barras_categoria(df, "Categoría", "Declaraciones por categoría de socio")

    col3, col4 = st.columns(2)
    with col3:
        grafico_torta(df, "Estado Declaración", "Estado de la declaración")
    with col4:
        grafico_torta(df, "Uso IA", "Declaración de uso de IA")

    st.divider()
    seccion_demografia(df, "autorales")

# ------------------------------------------------------------
# TAB: BOLETINES CONEXOS
# ------------------------------------------------------------
with tab_conexos:
    df = hojas_filtradas["Boletines Conexos"]

    titulo_seccion("Indicadores clave")
    c1, c2, c3, c4 = st.columns(4)
    total = len(df)
    pct_ingresado = (df["Estado"].eq("Conexo ingresado").mean() * 100) if total else 0
    pct_agrupacion = (df["Tipo"].eq("Agrupación").mean() * 100) if total else 0
    tipo_uso_top = df["Tipo de uso"].mode().iloc[0] if total and not df["Tipo de uso"].mode().empty else "—"
    c1.metric("Total declaraciones", f"{total:,}".replace(",", "."))
    c2.metric("Conexo ingresado", f"{pct_ingresado:.1f}%")
    c3.metric("Son agrupación", f"{pct_agrupacion:.1f}%")
    c4.metric("Tipo de uso más frecuente", tipo_uso_top)

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        grafico_evolucion_mensual(df, "Evolución mensual de declaraciones", color=COLOR_SECONDARY)
    with col2:
        grafico_barras_categoria(df, "Categoría", "Declaraciones por categoría de socio")

    col3, col4 = st.columns(2)
    with col3:
        grafico_torta(df, "Tipo", "Solista vs. Agrupación")
    with col4:
        grafico_barras_categoria(df, "Tipo de uso", "Distribución por tipo de uso", horizontal=True)

    st.divider()
    seccion_demografia(df, "conexos")

# ------------------------------------------------------------
# TAB: ANTICIPOS REAJUSTABLES
# ------------------------------------------------------------
with tab_anticipos:
    df = hojas_filtradas["Anticipos Reajustables"]

    titulo_seccion("Indicadores clave")
    c1, c2, c3 = st.columns(3)
    total = len(df)
    pct_aprobado = (df["Estado solicitud"].eq("Aprobado").mean() * 100) if total else 0
    pct_firma = (df["Firma Contrato de Cesión fiduciaria"].eq("Si").mean() * 100) if total else 0
    c1.metric("Total solicitudes", f"{total:,}".replace(",", "."))
    c2.metric("Solicitudes aprobadas", f"{pct_aprobado:.1f}%")
    c3.metric("Firma contrato de cesión", f"{pct_firma:.1f}%")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        grafico_evolucion_mensual(df, "Evolución mensual de solicitudes", color=COLOR_ACCENT)
    with col2:
        grafico_torta(df, "Estado solicitud", "Estado de la solicitud")

    col3, col4 = st.columns(2)
    with col3:
        grafico_barras_categoria(df, "Categoría", "Solicitudes por categoría de socio")
    with col4:
        if not df.empty and "Estado solicitud" in df.columns and "Firma Contrato de Cesión fiduciaria" in df.columns:
            cruzado = df.groupby(["Estado solicitud", "Firma Contrato de Cesión fiduciaria"], observed=True) \
                .size().reset_index(name="Cantidad")
            fig = px.bar(cruzado, x="Estado solicitud", y="Cantidad", color="Firma Contrato de Cesión fiduciaria",
                         template=PLOTLY_TEMPLATE, barmode="group", color_discrete_sequence=PALETTE)
            fig.update_layout(title="Estado de solicitud vs. firma de contrato", height=380)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")

    st.divider()
    seccion_demografia(df, "anticipos")

# ------------------------------------------------------------
# TAB: APORTE VIAJE
# ------------------------------------------------------------
with tab_viaje:
    df = hojas_filtradas["Aporte viaje"]

    titulo_seccion("Indicadores clave")
    c1, c2, c3 = st.columns(3)
    total = len(df)
    monto_total = df["Monto Liquido"].sum() if total else 0
    monto_prom = df["Monto Liquido"].mean() if total else 0
    c1.metric("Total solicitudes", f"{total:,}".replace(",", "."))
    c2.metric("Monto líquido total", f"${monto_total:,.0f}".replace(",", "."))
    c3.metric("Monto líquido promedio", f"${monto_prom:,.0f}".replace(",", "."))

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        grafico_evolucion_mensual(df, "Evolución mensual de solicitudes", color=COLOR_GOLD)
    with col2:
        grafico_barras_categoria(df, "País(es) Destino", "Top 10 países de destino", top_n=10, horizontal=True)

    col3, col4 = st.columns(2)
    with col3:
        if not df.empty and "Monto Liquido" in df.columns:
            fig = px.histogram(df, x="Monto Liquido", nbins=25, template=PLOTLY_TEMPLATE,
                               color_discrete_sequence=[COLOR_SAGE])
            fig.update_layout(title="Distribución del monto líquido", height=380)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")
    with col4:
        grafico_barras_categoria(df, "Categoría", "Solicitudes por categoría de socio")

    st.divider()
    seccion_demografia(df, "viaje")

# ------------------------------------------------------------
# TAB: PAGOS
# ------------------------------------------------------------
with tab_pagos:
    df = hojas_filtradas["Pagos"]

    titulo_seccion("Indicadores clave")
    c1, c2, c3 = st.columns(3)
    total = len(df)
    monto_total = df["Monto"].sum() if total else 0
    monto_prom = df["Monto"].mean() if total else 0
    c1.metric("Total de pagos", f"{total:,}".replace(",", "."))
    c2.metric("Monto total pagado", f"${monto_total:,.0f}".replace(",", "."))
    c3.metric("Monto promedio", f"${monto_prom:,.0f}".replace(",", "."))

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        grafico_evolucion_mensual(df, "Evolución mensual de pagos (cantidad)", color=COLOR_PLUM)
    with col2:
        if not df.empty and "Año" in df.columns and "Mes" in df.columns and "Monto" in df.columns:
            monto_mensual = df.groupby(["Año", "Mes"], observed=True)["Monto"].sum().reset_index()
            monto_mensual["Periodo"] = monto_mensual["Año"].astype(str) + " - " + monto_mensual["Mes"].astype(str)
            monto_mensual = monto_mensual.sort_values(["Año", "Mes"])
            fig = px.line(monto_mensual, x="Periodo", y="Monto", markers=True, template=PLOTLY_TEMPLATE,
                          color_discrete_sequence=[COLOR_ACCENT])
            fig.update_layout(title="Evolución mensual del monto pagado", height=380, xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")

    col3, col4 = st.columns(2)
    with col3:
        if not df.empty and "Tipo de Pago" in df.columns:
            conteo = df["Tipo de Pago"].value_counts().reset_index()
            conteo.columns = ["Tipo de Pago", "Cantidad"]
            fig = px.bar(conteo, x="Cantidad", y="Tipo de Pago", orientation="h", template=PLOTLY_TEMPLATE,
                         color_discrete_sequence=[COLOR_PRIMARY], text="Cantidad")
            fig.update_traces(textposition="outside")
            fig.update_layout(title="Cantidad de pagos por tipo", height=460,
                              yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")
    with col4:
        if not df.empty and "Tipo de Pago" in df.columns and "Monto" in df.columns:
            monto_tipo = df.groupby("Tipo de Pago", observed=True)["Monto"].sum().reset_index()
            monto_tipo = monto_tipo.sort_values("Monto", ascending=True)
            fig = px.bar(monto_tipo, x="Monto", y="Tipo de Pago", orientation="h", template=PLOTLY_TEMPLATE,
                         color_discrete_sequence=[COLOR_SECONDARY])
            fig.update_layout(title="Monto total pagado por tipo", height=460)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")

    st.divider()
    grafico_barras_categoria(df, "Categoría", "Pagos por categoría de socio")

    st.divider()
    seccion_demografia(df, "pagos")
