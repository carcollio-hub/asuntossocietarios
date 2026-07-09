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
    page_title="Panel de Gestiones - Gerencia de Asuntos Societarios | SCD",
    page_icon="🔖",
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

# Colores suaves y claros, especiales para la distribución por sexo
COLOR_MAP_SEXO = {"Hombre": "#7FB3D5", "Mujer": "#F5B7A5"}

MESES_ORDEN = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
               "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

RANGO_EDAD_ORDEN = ["0 a 11", "12 a 18", "19 a 30", "31 a 40", "41 a 50", "51 a 60", "61 a 100"]

MES_A_TRIMESTRE = {
    "enero": "T1", "febrero": "T1", "marzo": "T1",
    "abril": "T2", "mayo": "T2", "junio": "T2",
    "julio": "T3", "agosto": "T3", "septiembre": "T3",
    "octubre": "T4", "noviembre": "T4", "diciembre": "T4",
}
TRIMESTRE_ORDEN = ["T1", "T2", "T3", "T4"]

PLOTLY_TEMPLATE = "plotly_white"


def formatear_moneda_exacta(valor) -> str:
    """Formato $1.234.567 (separador de miles con punto, sin decimales)."""
    if valor is None or pd.isna(valor):
        return "$0"
    return f"${valor:,.0f}".replace(",", ".")


def formatear_moneda_compacta(valor) -> str:
    """Formato abreviado y legible para montos grandes: $3,99 mil millones / $290,8 millones."""
    if valor is None or pd.isna(valor):
        return "$0"
    valor = float(valor)
    signo = "-" if valor < 0 else ""
    valor = abs(valor)
    if valor >= 1_000_000_000:
        return f"{signo}${valor / 1_000_000_000:.2f} mil millones"
    if valor >= 1_000_000:
        return f"{signo}${valor / 1_000_000:.1f} millones"
    return f"{signo}{formatear_moneda_exacta(valor)}"

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


def _limpiar_monto_pagos(valor):
    """Convierte el campo Monto a numérico. Maneja casos como 'CLP 161.476' (texto con
    prefijo de moneda y punto como separador de miles), que antes se perdían como nulos."""
    if pd.isna(valor):
        return None
    if isinstance(valor, str):
        texto = valor.upper().replace("CLP", "").replace(".", "").replace(",", "").strip()
        if texto == "":
            return None
        try:
            return float(texto)
        except ValueError:
            return None
    return float(valor)


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

        columnas_unnamed = [c for c in df.columns if str(c).startswith("Unnamed")]
        if columnas_unnamed:
            df = df.drop(columns=columnas_unnamed)

        if "Mes" in df.columns:
            df["Mes"] = pd.Categorical(df["Mes"], categories=MESES_ORDEN, ordered=True)
            df["Trimestre"] = df["Mes"].astype(str).map(MES_A_TRIMESTRE)
            df["Trimestre"] = pd.Categorical(df["Trimestre"], categories=TRIMESTRE_ORDEN, ordered=True)

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
            df["Monto"] = df["Monto"].apply(_limpiar_monto_pagos)
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


def aplicar_filtros(df: pd.DataFrame, sexo, region, edad, anio, trimestre) -> pd.DataFrame:
    df_f = df.copy()
    if sexo and "Sexo" in df_f.columns:
        df_f = df_f[df_f["Sexo"].isin(sexo)]
    if region and "Región" in df_f.columns:
        df_f = df_f[df_f["Región"].isin(region)]
    if edad and "Rango edad" in df_f.columns:
        df_f = df_f[df_f["Rango edad"].astype(str).isin(edad)]
    if anio and "Año" in df_f.columns:
        df_f = df_f[df_f["Año"].isin(anio)]
    if trimestre and "Trimestre" in df_f.columns:
        df_f = df_f[df_f["Trimestre"].astype(str).isin(trimestre)]
    return df_f


# ============================================================
# FUNCIONES AUXILIARES DE VISUALIZACIÓN
# ============================================================
def delta_ultimo_trimestre(df: pd.DataFrame):
    """Retorna la variación % del último trimestre disponible vs el trimestre anterior, o None si no aplica."""
    if df.empty or "Año" not in df.columns or "Mes" not in df.columns:
        return None
    temp = _agregar_periodo_trimestral(df)
    conteo = temp.groupby(["Año", "Trimestre", "Periodo"], observed=True).size().reset_index(name="Cantidad")
    conteo = conteo.sort_values(["Año", "Trimestre"])
    if len(conteo) < 2:
        return None
    actual = conteo.iloc[-1]["Cantidad"]
    anterior = conteo.iloc[-2]["Cantidad"]
    if anterior == 0:
        return None
    variacion = (actual - anterior) / anterior * 100
    return f"{variacion:+.1f}% vs trimestre anterior"


def titulo_seccion(texto: str):
    st.markdown(f"<h3 class='seccion-titulo'>{texto}</h3>", unsafe_allow_html=True)


def _agregar_periodo_trimestral(df: pd.DataFrame) -> pd.DataFrame:
    temp = df.copy()
    temp["Trimestre"] = temp["Mes"].astype(str).map(MES_A_TRIMESTRE)
    temp["Periodo"] = temp["Año"].astype(str) + " " + temp["Trimestre"]
    temp["Trimestre"] = pd.Categorical(temp["Trimestre"], categories=TRIMESTRE_ORDEN, ordered=True)
    return temp


def grafico_evolucion_trimestral(df: pd.DataFrame, titulo: str, color=COLOR_PRIMARY, y_title="N° de registros"):
    if df.empty or "Año" not in df.columns or "Mes" not in df.columns:
        st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")
        return
    temp = _agregar_periodo_trimestral(df)
    conteo = temp.groupby(["Año", "Trimestre", "Periodo"], observed=True).size().reset_index(name="Cantidad")
    conteo = conteo.sort_values(["Año", "Trimestre"])
    max_val = conteo["Cantidad"].max() if not conteo.empty else 0
    fig = px.bar(conteo, x="Periodo", y="Cantidad", template=PLOTLY_TEMPLATE,
                 color_discrete_sequence=[color], text="Cantidad")
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(title=titulo, xaxis_title="Trimestre", yaxis_title=y_title,
                       height=380, yaxis=dict(range=[0, max_val * 1.2]))
    st.plotly_chart(fig, use_container_width=True)


def grafico_barras_categoria(df: pd.DataFrame, columna: str, titulo: str, top_n=None, horizontal=False,
                              mostrar_porcentaje=False, color_unico=None):
    if df.empty or columna not in df.columns:
        st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")
        return
    conteo = df[columna].value_counts(dropna=True).reset_index()
    conteo.columns = [columna, "Cantidad"]
    total_general = conteo["Cantidad"].sum()
    if top_n:
        conteo = conteo.head(top_n)
    if mostrar_porcentaje and total_general:
        conteo["Etiqueta"] = conteo["Cantidad"].astype(str) + " (" + \
            (conteo["Cantidad"] / total_general * 100).round(1).astype(str) + "%)"
    else:
        conteo["Etiqueta"] = conteo["Cantidad"].astype(str)

    orient = "h" if horizontal else "v"
    x_col, y_col = ("Cantidad", columna) if horizontal else (columna, "Cantidad")
    max_val = conteo["Cantidad"].max() if not conteo.empty else 0

    if color_unico:
        fig = px.bar(conteo, x=x_col, y=y_col, orientation=orient, template=PLOTLY_TEMPLATE,
                     color_discrete_sequence=[color_unico], text="Etiqueta")
    else:
        fig = px.bar(conteo, x=x_col, y=y_col, orientation=orient, template=PLOTLY_TEMPLATE,
                     color=columna, color_discrete_sequence=PALETTE, text="Etiqueta")
    fig.update_traces(textposition="outside", showlegend=False, cliponaxis=False)
    fig.update_layout(title=titulo, height=380, uniformtext_minsize=10, uniformtext_mode="hide")
    if horizontal:
        fig.update_layout(yaxis={"categoryorder": "total ascending"})
        fig.update_xaxes(range=[0, max_val * 1.22])
    else:
        fig.update_yaxes(range=[0, max_val * 1.22])
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
st.title("🔖 Panel Operativo de Gestiones")
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
CLAVES_FILTROS = ["filtro_sexo", "filtro_region", "filtro_edad", "filtro_anio", "filtro_trimestre"]


def _resetear_filtros():
    for clave in CLAVES_FILTROS:
        st.session_state[clave] = []


with st.sidebar:
    st.markdown(f"""
    <div style="background-color:{COLOR_PRIMARY}; padding:14px 16px; border-radius:10px; margin-bottom:14px;">
        <span style="color:white; font-size:1.15rem; font-weight:700;">🔎 Filtros del dashboard</span><br>
        <span style="color:#DCE6F1; font-size:0.82rem;">Se aplican a todas las pestañas y gráficos.</span>
    </div>
    """, unsafe_allow_html=True)

    opciones_sexo = obtener_opciones_filtro(hojas, "Sexo")
    opciones_region = obtener_opciones_filtro(hojas, "Región")
    opciones_edad = [e for e in RANGO_EDAD_ORDEN if e in obtener_opciones_filtro(hojas, "Rango edad")]
    opciones_anio = obtener_opciones_filtro(hojas, "Año")
    opciones_trimestre = [t for t in TRIMESTRE_ORDEN if t in obtener_opciones_filtro(hojas, "Trimestre")]

    st.markdown("**👤 Sexo**")
    filtro_sexo = st.multiselect("Sexo", opciones_sexo, default=[], key="filtro_sexo",
                                  placeholder="Todos", label_visibility="collapsed")

    st.markdown("**📍 Región**")
    filtro_region = st.multiselect("Región", opciones_region, default=[], key="filtro_region",
                                    placeholder="Todas", label_visibility="collapsed")

    st.markdown("**🎂 Rango de edad**")
    filtro_edad = st.multiselect("Rango edad", opciones_edad, default=[], key="filtro_edad",
                                  placeholder="Todos", label_visibility="collapsed")

    col_anio, col_trim = st.columns(2)
    with col_anio:
        st.markdown("**📅 Año**")
        filtro_anio = st.multiselect("Año", opciones_anio, default=[], key="filtro_anio",
                                      placeholder="Todos", label_visibility="collapsed")
    with col_trim:
        st.markdown("**📆 Trimestre**")
        filtro_trimestre = st.multiselect("Trimestre", opciones_trimestre, default=[], key="filtro_trimestre",
                                           placeholder="Todos", label_visibility="collapsed")

    st.button("🔄 Restablecer filtros", use_container_width=True, on_click=_resetear_filtros)
    st.caption("Si no seleccionas ninguna opción en un filtro, se consideran todos los valores.")

# Aplicar filtros a cada hoja
hojas_filtradas = {
    nombre: aplicar_filtros(df, filtro_sexo, filtro_region, filtro_edad, filtro_anio, filtro_trimestre)
    for nombre, df in hojas.items()
}

with st.sidebar:
    st.divider()
    st.markdown("**📊 Registros filtrados**")
    for _nombre in ["Boletines Autorales", "Boletines Conexos", "Anticipos Reajustables", "Aporte viaje", "Pagos"]:
        st.caption(f"{_nombre}: **{len(hojas_filtradas[_nombre]):,}**".replace(",", "."))
    st.divider()
    st.caption(f"Última actualización de datos: {pd.Timestamp.today().strftime('%d-%m-%Y')}")



# ============================================================
# TABS PRINCIPALES
# ============================================================
tab_resumen, tab_autorales, tab_conexos, tab_anticipos, tab_viaje, tab_pagos, tab_fondo, tab_convenios = st.tabs([
    "🏠 Resumen General",
    "📝 Boletines Autorales",
    "🎵 Boletines Conexos",
    "💰 Anticipos Reajustables",
    "✈️ Aporte Viaje",
    "💳 Pagos",
    "🚨 Fondo Emergencia",
    "🤝 Convenios",
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
    titulo_seccion("Evolución trimestral por gestión")
    st.caption("Cada gestión se muestra por separado (mismo formato usado en cada pestaña), ya que sus volúmenes tienen escalas muy distintas entre sí.")
    colores_gestion = {
        "Boletines Autorales": COLOR_PRIMARY,
        "Boletines Conexos": COLOR_SECONDARY,
        "Anticipos Reajustables": COLOR_ACCENT,
        "Aporte viaje": COLOR_GOLD,
        "Pagos": COLOR_PLUM,
    }
    fila1 = st.columns(3)
    fila2 = st.columns(2)
    columnas_grid = fila1 + fila2
    for col, nombre in zip(columnas_grid, nombres_kpi):
        with col:
            grafico_evolucion_trimestral(hojas_filtradas[nombre], nombre, color=colores_gestion[nombre])

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
    c1.metric("Total declaraciones", f"{total:,}".replace(",", "."), delta=delta_ultimo_trimestre(df))
    c2.metric("Obra ingresada", f"{pct_ingresado:.1f}%")
    c3.metric("Solicita transcripción", f"{pct_transcripcion:.1f}%")
    c4.metric("Declara uso de IA", f"{pct_ia:.1f}%")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        grafico_evolucion_trimestral(df, "Evolución trimestral de declaraciones", color=COLOR_PRIMARY)
    with col2:
        grafico_barras_categoria(df, "Categoría", "Declaraciones por categoría de socio")

    col3, col4 = st.columns(2)
    with col3:
        grafico_barras_categoria(df, "Estado Declaración", "Estado de la declaración", horizontal=True,
                                  mostrar_porcentaje=True, color_unico=COLOR_SECONDARY)
    with col4:
        grafico_barras_categoria(df, "Uso IA", "Declaración de uso de IA", horizontal=True,
                                  mostrar_porcentaje=True, color_unico=COLOR_GOLD)

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
    c1.metric("Total declaraciones", f"{total:,}".replace(",", "."), delta=delta_ultimo_trimestre(df))
    c2.metric("Conexo ingresado", f"{pct_ingresado:.1f}%")
    c3.metric("Son agrupación", f"{pct_agrupacion:.1f}%")
    c4.metric("Tipo de uso más frecuente", tipo_uso_top)

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        grafico_evolucion_trimestral(df, "Evolución trimestral de declaraciones", color=COLOR_SECONDARY)
    with col2:
        grafico_barras_categoria(df, "Categoría", "Declaraciones por categoría de socio")

    col3, col4 = st.columns(2)
    with col3:
        grafico_barras_categoria(df, "Tipo", "Solista vs. Agrupación", horizontal=True,
                                  mostrar_porcentaje=True, color_unico=COLOR_ACCENT)
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
    c1.metric("Total solicitudes", f"{total:,}".replace(",", "."), delta=delta_ultimo_trimestre(df))
    c2.metric("Solicitudes aprobadas", f"{pct_aprobado:.1f}%")
    c3.metric("Firma contrato de cesión", f"{pct_firma:.1f}%")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        grafico_evolucion_trimestral(df, "Evolución trimestral de solicitudes", color=COLOR_ACCENT)
    with col2:
        grafico_barras_categoria(df, "Categoría", "Solicitudes por categoría de socio")

    col3, col4 = st.columns(2)
    with col3:
        grafico_barras_categoria(df, "Estado solicitud", "Estado de la solicitud", horizontal=True,
                                  mostrar_porcentaje=True, color_unico=COLOR_PRIMARY)
    with col4:
        grafico_barras_categoria(df, "Firma Contrato de Cesión fiduciaria", "Firma de contrato de cesión fiduciaria",
                                  horizontal=True, mostrar_porcentaje=True, color_unico=COLOR_SAGE)

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
    c1.metric("Total solicitudes", f"{total:,}".replace(",", "."), delta=delta_ultimo_trimestre(df))
    c2.metric("Monto líquido total", formatear_moneda_compacta(monto_total),
              help=f"Valor exacto: {formatear_moneda_exacta(monto_total)}")
    c3.metric("Monto líquido promedio", formatear_moneda_exacta(monto_prom))

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        grafico_evolucion_trimestral(df, "Evolución trimestral de solicitudes", color=COLOR_GOLD)
    with col2:
        grafico_barras_categoria(df, "País(es) Destino", "Top 10 países de destino", top_n=10, horizontal=True)

    col3, col4 = st.columns(2)
    with col3:
        if not df.empty and "Monto Liquido" in df.columns and "Categoría" in df.columns:
            prom_cat = df.groupby("Categoría", observed=True)["Monto Liquido"].mean().reset_index()
            prom_cat = prom_cat.sort_values("Monto Liquido", ascending=False)
            prom_cat["Etiqueta"] = prom_cat["Monto Liquido"].apply(formatear_moneda_exacta)
            max_val = prom_cat["Monto Liquido"].max() if not prom_cat.empty else 0
            fig = px.bar(prom_cat, x="Categoría", y="Monto Liquido", template=PLOTLY_TEMPLATE,
                         color="Categoría", color_discrete_sequence=PALETTE, text="Etiqueta")
            fig.update_traces(textposition="outside", cliponaxis=False, showlegend=False)
            fig.update_layout(title="Monto líquido promedio por categoría de socio", height=380,
                              yaxis_title="Monto líquido promedio ($)",
                              yaxis=dict(range=[0, max_val * 1.25]))
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
    c1, c2, c3, c4 = st.columns(4)
    total = len(df)
    pagos_con_monto = int(df["Monto"].notna().sum()) if total else 0
    pagos_sin_monto = total - pagos_con_monto
    monto_total = df["Monto"].sum() if total else 0
    monto_prom = (monto_total / pagos_con_monto) if pagos_con_monto else 0
    c1.metric("Total de pagos", f"{total:,}".replace(",", "."), delta=delta_ultimo_trimestre(df))
    c2.metric("Pagos con monto registrado", f"{pagos_con_monto:,}".replace(",", "."))
    c3.metric("Monto total pagado", formatear_moneda_compacta(monto_total),
              help=f"Valor exacto: {formatear_moneda_exacta(monto_total)}")
    c4.metric("Monto promedio por pago", formatear_moneda_exacta(monto_prom),
              help="Se calcula como Monto total pagado ÷ Pagos con monto registrado.")
    if pagos_sin_monto > 0:
        st.caption(f"⚠️ {pagos_sin_monto} pago(s) no tienen un monto registrado en la base y se excluyen "
                   f"del cálculo de Monto total y Monto promedio.")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        grafico_evolucion_trimestral(df, "Evolución trimestral de pagos (cantidad)", color=COLOR_PLUM)
    with col2:
        if not df.empty and "Año" in df.columns and "Mes" in df.columns and "Monto" in df.columns:
            temp = _agregar_periodo_trimestral(df)
            monto_trim = temp.groupby(["Año", "Trimestre", "Periodo"], observed=True)["Monto"].sum().reset_index()
            monto_trim = monto_trim.sort_values(["Año", "Trimestre"])
            monto_trim["Monto (MM$)"] = (monto_trim["Monto"] / 1_000_000).round(1)
            max_val = monto_trim["Monto (MM$)"].max() if not monto_trim.empty else 0
            fig = px.bar(monto_trim, x="Periodo", y="Monto (MM$)", template=PLOTLY_TEMPLATE,
                         color_discrete_sequence=[COLOR_ACCENT], text="Monto (MM$)")
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(title="Evolución trimestral del monto pagado", xaxis_title="Trimestre",
                              yaxis_title="Monto pagado (millones de $)", height=380,
                              yaxis=dict(range=[0, max_val * 1.2]))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")

    col3, col4 = st.columns(2)
    with col3:
        if not df.empty and "Tipo de Pago" in df.columns:
            conteo = df["Tipo de Pago"].value_counts().reset_index()
            conteo.columns = ["Tipo de Pago", "Cantidad"]
            max_val = conteo["Cantidad"].max() if not conteo.empty else 0
            fig = px.bar(conteo, x="Cantidad", y="Tipo de Pago", orientation="h", template=PLOTLY_TEMPLATE,
                         color_discrete_sequence=[COLOR_PRIMARY], text="Cantidad")
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(title="Cantidad de pagos por tipo", height=460,
                              yaxis={"categoryorder": "total ascending"},
                              xaxis=dict(range=[0, max_val * 1.22]),
                              uniformtext_minsize=10, uniformtext_mode="hide")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")
    with col4:
        if not df.empty and "Tipo de Pago" in df.columns and "Monto" in df.columns:
            monto_tipo = df.groupby("Tipo de Pago", observed=True)["Monto"].sum().reset_index()
            monto_tipo["Monto (MM$)"] = (monto_tipo["Monto"] / 1_000_000).round(1)
            monto_tipo = monto_tipo.sort_values("Monto (MM$)", ascending=True)
            max_val = monto_tipo["Monto (MM$)"].max() if not monto_tipo.empty else 0
            fig = px.bar(monto_tipo, x="Monto (MM$)", y="Tipo de Pago", orientation="h", template=PLOTLY_TEMPLATE,
                         color_discrete_sequence=[COLOR_SECONDARY], text="Monto (MM$)")
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(title="Monto total pagado por tipo (millones de $)", height=460,
                              xaxis_title="Millones de $", xaxis=dict(range=[0, max_val * 1.22]),
                              uniformtext_minsize=10, uniformtext_mode="hide")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")

    st.divider()
    grafico_barras_categoria(df, "Categoría", "Pagos por categoría de socio")

    st.divider()
    seccion_demografia(df, "pagos")

# ------------------------------------------------------------
# TAB: FONDO EMERGENCIA
# ------------------------------------------------------------
with tab_fondo:
    df = hojas_filtradas["Fondo Emergencia"]

    titulo_seccion("Indicadores clave")
    c1, c2, c3 = st.columns(3)
    total = len(df)
    monto_total = df["Monto otorgado"].sum() if total else 0
    monto_prom = df["Monto otorgado"].mean() if total else 0
    c1.metric("Total solicitudes", f"{total:,}".replace(",", "."), delta=delta_ultimo_trimestre(df))
    c2.metric("Monto total otorgado", formatear_moneda_compacta(monto_total),
              help=f"Valor exacto: {formatear_moneda_exacta(monto_total)}")
    c3.metric("Monto promedio otorgado", formatear_moneda_exacta(monto_prom))

    st.divider()
    titulo_seccion("Tipos de solicitud recibidas")
    if not df.empty and "MOTIVO" in df.columns:
        conteo_motivo = df["MOTIVO"].value_counts().reset_index()
        conteo_motivo.columns = ["MOTIVO", "Cantidad"]
        fig = px.treemap(conteo_motivo, path=["MOTIVO"], values="Cantidad",
                         color="MOTIVO", color_discrete_sequence=PALETTE)
        fig.update_traces(textinfo="label+value+percent root", textfont_size=16)
        fig.update_layout(title="Distribución de solicitudes por motivo", height=420,
                          margin=dict(t=50, l=10, r=10, b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        grafico_evolucion_trimestral(df, "Evolución trimestral de solicitudes", color=COLOR_SAGE)
    with col2:
        if not df.empty and "Año" in df.columns and "Mes" in df.columns and "Monto otorgado" in df.columns:
            temp = _agregar_periodo_trimestral(df)
            monto_trim = temp.groupby(["Año", "Trimestre", "Periodo"], observed=True)["Monto otorgado"] \
                .sum().reset_index()
            monto_trim = monto_trim.sort_values(["Año", "Trimestre"])
            monto_trim["Monto (MM$)"] = (monto_trim["Monto otorgado"] / 1_000_000).round(2)
            max_val = monto_trim["Monto (MM$)"].max() if not monto_trim.empty else 0
            fig = px.bar(monto_trim, x="Periodo", y="Monto (MM$)", template=PLOTLY_TEMPLATE,
                         color_discrete_sequence=[COLOR_GOLD], text="Monto (MM$)")
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(title="Evolución trimestral del monto otorgado", xaxis_title="Trimestre",
                              yaxis_title="Monto (millones de $)", height=380,
                              yaxis=dict(range=[0, max_val * 1.2]))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")

    col3, col4 = st.columns(2)
    with col3:
        grafico_barras_categoria(df, "Categoría", "Solicitudes por categoría de socio")
    with col4:
        if not df.empty and "MOTIVO" in df.columns and "Monto otorgado" in df.columns:
            monto_motivo = df.groupby("MOTIVO", observed=True)["Monto otorgado"].mean().reset_index()
            monto_motivo = monto_motivo.sort_values("Monto otorgado", ascending=False)
            monto_motivo["Etiqueta"] = monto_motivo["Monto otorgado"].apply(formatear_moneda_exacta)
            max_val = monto_motivo["Monto otorgado"].max() if not monto_motivo.empty else 0
            fig = px.bar(monto_motivo, x="MOTIVO", y="Monto otorgado", template=PLOTLY_TEMPLATE,
                         color="MOTIVO", color_discrete_sequence=PALETTE, text="Etiqueta")
            fig.update_traces(textposition="outside", cliponaxis=False, showlegend=False)
            fig.update_layout(title="Monto promedio otorgado por motivo", height=380,
                              yaxis_title="Monto promedio ($)", yaxis=dict(range=[0, max_val * 1.25]))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos suficientes para este gráfico con los filtros seleccionados.")

    st.divider()
    seccion_demografia(df, "fondo")

# ------------------------------------------------------------
# TAB: CONVENIOS
# ------------------------------------------------------------
with tab_convenios:
    df = hojas_filtradas["Convenios"]
    st.caption("Los convenios corresponden a acuerdos vigentes con empresas para acceso a descuentos y "
               "beneficios; no son solicitudes de socios, por lo que no se ven afectados por los filtros "
               "de sexo, región, edad, año o trimestre del panel lateral.")

    titulo_seccion("Indicadores clave")
    c1, c2, c3, c4 = st.columns(4)
    total_convenios = int(df["Cantidad"].sum()) if not df.empty else 0
    tipos_distintos = df["Tipo"].nunique() if not df.empty else 0
    nacional = int(df.loc[df["Presencia"] == "Nacional", "Cantidad"].sum()) if not df.empty else 0
    metropolitana = int(df.loc[df["Presencia"] == "Metropolitana", "Cantidad"].sum()) if not df.empty else 0
    c1.metric("Total convenios vigentes", f"{total_convenios:,}".replace(",", "."))
    c2.metric("Tipos de convenio", f"{tipos_distintos:,}".replace(",", "."))
    c3.metric("Alcance nacional", f"{nacional:,}".replace(",", "."))
    c4.metric("Alcance Región Metropolitana", f"{metropolitana:,}".replace(",", "."))

    st.divider()
    titulo_seccion("Tipos de convenio y alcance geográfico")
    if not df.empty:
        fig = px.treemap(df, path=["Presencia", "Tipo"], values="Cantidad",
                         color="Tipo", color_discrete_sequence=PALETTE)
        fig.update_traces(textinfo="label+value", textfont_size=15)
        fig.update_layout(title="Convenios por tipo y alcance geográfico", height=460,
                          margin=dict(t=50, l=10, r=10, b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No hay datos disponibles.")

    st.divider()
    if not df.empty:
        orden = df.sort_values("Cantidad", ascending=True)
        max_val = orden["Cantidad"].max()
        fig = px.bar(orden, x="Cantidad", y="Tipo", orientation="h", template=PLOTLY_TEMPLATE,
                     color="Tipo", color_discrete_sequence=PALETTE, text="Cantidad")
        fig.update_traces(textposition="outside", showlegend=False, cliponaxis=False)
        fig.update_layout(title="Cantidad de convenios por tipo", height=420,
                          yaxis={"categoryorder": "total ascending"},
                          xaxis=dict(range=[0, max_val * 1.25]))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No hay datos disponibles.")
