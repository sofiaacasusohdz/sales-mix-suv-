# -*- coding: utf-8 -*-
"""
Sales Mix Analyzer

Sube un archivo (Excel o CSV) con ventas por mes de coches (marca, modelo, versión, segmento) y
esta app calculará automáticamente:

• Ventas por mes y por segmento
• Crecimiento % mes a mes (MoM) y Δ unidades
• Participación (mix) por segmento, marca, modelo y versión
• Análisis de competidores vs Volvo
• Gráficas interactivas y descarga de resultados en Excel

Cómo ejecutar:
1) pip install streamlit pandas numpy plotly openpyxl xlsxwriter
2) streamlit run sales_mix_app.py

Estructura típica del archivo de entrada (columnas sugeridas; pueden variar):
- Marca (o Brand/Make) — opcional
- Modelo (obligatorio)
- Version (o Versión) — opcional
- Segmento (ej. "SUV B", "SUV-C", "B-SUV", "SUV C"), se normaliza a B/C/D/E
- Meses como columnas: puede ser "Ene", "Feb", ..., "Dic" o "Jan"..\"Dec\" o "2025-01" etc.

La app incluye un mapeador de columnas para adaptarse a tus encabezados.
"""

import io
import re
from typing import List, Dict, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# =========================
# CONFIGURACIÓN DE PÁGINA
# =========================
st.set_page_config(page_title="Sales Mix SUVs — Analyzer", layout="wide")
st.title("Sales Mix Analyzer")
st.caption("Carga tu archivo y obtén análisis, tablas, gráficas y un Excel descargable.")

# =========================
# UTILIDADES
# =========================
MONTH_ALIASES = {
    # Español
    "ene": 1, "enero": 1,
    "feb": 2, "febrero": 2,
    "mar": 3, "marzo": 3,
    "abr": 4, "abril": 4,
    "may": 5, "mayo": 5,
    "jun": 6, "junio": 6,
    "jul": 7, "julio": 7,
    "ago": 8, "agosto": 8,
    "sep": 9, "sept": 9, "septiembre": 9,
    "oct": 10, "octubre": 10,
    "nov": 11, "noviembre": 11,
    "dic": 12, "diciembre": 12,
    # Inglés
    "jan": 1, "january": 1,
    "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "september": 9,
    "october": 10,
    "november": 11,
    "dec": 12, "december": 12,
}

SEGMENT_PAT = re.compile(r"(?i)s\s*uv\s*[-_ ]*([bcde])|^([bcde])$")
DATE_PAT = re.compile(r"^(?P<y>20\d{2})[-_\.\/ ]?(?P<m>0?[1-9]|1[0-2])$")


@st.cache_data(show_spinner=False)
def detect_month_columns(df: pd.DataFrame) -> List[str]:
    """Detecta columnas que parecen meses. Ordena de enero..diciembre (o por año-mes)."""
    month_cols = []
    for col in df.columns:
        s = str(col).strip()
        s_l = s.lower()
        # a) Año-Mes tipo 2024-01, 2024_1, 2024/01, etc.
        m = DATE_PAT.match(s_l)
        if m:
            month_cols.append(col)
            continue
        # b) Alias de mes: ene, enero, jan, etc. opcionalmente con año
        tokens = re.split(r"[^a-zA-Záéíóúñ0-9]+", s_l)
        tokens = [t for t in tokens if t]
        if not tokens:
            continue
        found = False
        for t in tokens:
            if t in MONTH_ALIASES:
                found = True
                break
        if found:
            month_cols.append(col)
            continue
        # c) Números 1..12 simples
        if s_l.isdigit():
            try:
                v = int(s_l)
                if 1 <= v <= 12:
                    month_cols.append(col)
                    continue
            except Exception:
                pass
    # Ordenar: intentar por año-mes si aplica; si no, por mes 1..12 según alias
    def month_key(c):
        s = str(c).strip().lower()
        m = DATE_PAT.match(s)
        if m:
            y = int(m.group("y"))
            mo = int(m.group("m"))
            return (y, mo)
        # buscar alias
        for k, v in MONTH_ALIASES.items():
            if k in s:
                return (9999, v)  # sin año → al final pero ordenado por mes
        # num puro
        if s.isdigit():
            return (9999, int(s))
        return (9999, 99)

    month_cols_sorted = sorted(month_cols, key=month_key)
    return month_cols_sorted


def normalize_segment(val) -> str | None:
    if pd.isna(val):
        return None
    s = str(val).strip()
    m = SEGMENT_PAT.search(s)
    if m:
        letter = m.group(1) or m.group(2)
        return letter.upper()
    return None


def pick_first_present(candidates: List[str], cols: List[str]) -> str | None:
    lookup = {c.lower(): c for c in cols}
    for cand in candidates:
        if cand.lower() in lookup:
            return lookup[cand.lower()]
    return None


def build_share(df: pd.DataFrame, group_keys: List[str], value_col: str, share_col: str, by_keys: List[str]):
    """Agrega columna de share (participación) dentro de by_keys.
    group_keys: columnas de agrupación del numerador.
    by_keys: columnas que definen el grupo de denominador (ej. segmento + mes).
    """
    totals = df.groupby(by_keys, dropna=False, as_index=False)[value_col].sum().rename(columns={value_col: "__denom__"})
    out = df.merge(totals, on=by_keys, how="left")
    out[share_col] = np.where(out["__denom__"].eq(0), np.nan, out[value_col] / out["__denom__"])
    out.drop(columns=["__denom__"], inplace=True)
    return out


@st.cache_data(show_spinner=False)
def tidy_data(df: pd.DataFrame, config: Dict[str, str]) -> Tuple[pd.DataFrame, List[pd.Timestamp]]:
    """Convierte de formato ancho (meses como columnas) a largo.
    config = {"brand": col/None, "model": col, "version": col/None, "segment": col}
    Devuelve (df_long, months_sorted)
    """
    month_cols = detect_month_columns(df)
    if not month_cols:
        raise ValueError("No se detectaron columnas de meses. Revisa encabezados (Ene..Dic, Jan..Dec, 2025-01, etc.).")

    id_cols = [c for c in [config.get("brand"), config.get("model"), config.get("version"), config.get("segment")] if c]
    value_vars = month_cols

    melted = df[id_cols + value_vars].copy()
    # a numérico
    for c in value_vars:
        melted[c] = pd.to_numeric(melted[c], errors="coerce")
    melted = melted.melt(id_vars=id_cols, value_vars=value_vars, var_name="MesCol", value_name="Unidades").fillna({"Unidades": 0})

    # Parsear MesCol → datetime del primer día del mes
    def parse_month(s: str) -> pd.Timestamp:
        s_l = str(s).strip().lower()
        m = DATE_PAT.match(s_l)
        if m:
            y = int(m.group("y")); mo = int(m.group("m"))
            return pd.Timestamp(year=y, month=mo, day=1)
        # alias
        # buscar un token que sea mes, y un token que sea año
        tokens = re.split(r"[^a-zA-Záéíóúñ0-9]+", s_l)
        year = None
        mo = None
        for t in tokens:
            if t in MONTH_ALIASES and mo is None:
                mo = MONTH_ALIASES[t]
            elif t.isdigit() and len(t) == 4 and t.startswith("20"):
                year = int(t)
        if mo is not None and year is not None:
            return pd.Timestamp(year=year, month=mo, day=1)
        if mo is not None and year is None:
            # si no hay año, usar 1900 provisoriamente para ordenar; luego normalizamos al rango
            return pd.Timestamp(year=1900, month=mo, day=1)
        # número 1..12
        if s_l.isdigit():
            v = int(s_l)
            if 1 <= v <= 12:
                return pd.Timestamp(year=1900, month=v, day=1)
        # fallback
        return pd.NaT

    melted["Mes"] = melted["MesCol"].map(parse_month)
    if melted["Mes"].isna().all():
        raise ValueError("No se pudieron interpretar las columnas de mes. Renombra a formatos reconocidos (Ene, Feb, 2025-01, etc.).")

    # Normalizar año si se usó 1900
    # Si hay años válidos, sustituir 1900 por el año más frecuente o por el máximo año presente
    real_years = melted.loc[melted["Mes"].dt.year.ne(1900) & melted["Mes"].notna(), "Mes"].dt.year
    if not real_years.empty:
        target_year = int(real_years.mode().iloc[0])
        mask1900 = melted["Mes"].dt.year.eq(1900)
        melted.loc[mask1900, "Mes"] = melted.loc[mask1900, "Mes"].apply(lambda dt: pd.Timestamp(year=target_year, month=dt.month, day=1))

    # Segmento normalizado
    seg_col = config.get("segment")
    melted["SegmentoSUV"] = melted[seg_col].apply(normalize_segment)

    # Limpiar columnas básicas
    if config.get("brand"):
        melted.rename(columns={config["brand"]: "Marca"}, inplace=True)
    else:
        melted["Marca"] = np.nan
    if config.get("model"):
        melted.rename(columns={config["model"]: "Modelo"}, inplace=True)
    if config.get("version"):
        melted.rename(columns={config["version"]: "Version"}, inplace=True)
    else:
        melted["Version"] = np.nan

    melted = melted[["Marca", "Modelo", "Version", seg_col, "SegmentoSUV", "Mes", "Unidades"]].rename(columns={seg_col: "SegmentoRaw"})
    melted.dropna(subset=["Modelo", "Mes"], inplace=True)

    # Orden meses
    months_sorted = sorted(melted["Mes"].dropna().unique())
    return melted, months_sorted


@st.cache_data(show_spinner=False)
def compute_metrics(df_long: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Calcula métricas clave y devuelve varios DataFrames listos para mostrar/descargar."""
    d = df_long.copy()
    d["Unidades"] = pd.to_numeric(d["Unidades"], errors="coerce").fillna(0).astype(float)

    # Totales por Mes y SegmentoSUV (sólo B/C/D/E)
    d_suv = d[d["SegmentoSUV"].isin(list("BCDE"))].copy()

    # MoM por grupo (SegmentoSUV, Marca, Modelo, Version)
    group_cols = ["SegmentoSUV", "Marca", "Modelo", "Version"]
    d_suv.sort_values(["Modelo", "Version", "Mes"], inplace=True)

    def add_mom(g: pd.DataFrame) -> pd.DataFrame:
        g = g.sort_values("Mes").copy()
        g["Unidades_prev"] = g["Unidades"].shift(1)
        g["Delta"] = g["Unidades"] - g["Unidades_prev"]
        g["MoM_%"] = np.where(g["Unidades_prev"].fillna(0) == 0, np.nan, (g["Unidades"] / g["Unidades_prev"]) - 1.0)
        return g

    by = d_suv.groupby(group_cols, dropna=False, as_index=False, group_keys=False)
    d_mom = by.apply(add_mom)

    # Participación por segmento-mes (mix dentro del segmento)
    by_keys = ["SegmentoSUV", "Mes"]
    d_share_seg = build_share(d_mom, group_cols, "Unidades", "Share_Segmento", by_keys)

    # Participación dentro de cada modelo (para ver mix de versiones)
    by_keys_model = ["SegmentoSUV", "Modelo", "Mes"]
    d_share_model = build_share(d_share_seg, group_cols, "Unidades", "Share_Modelo", by_keys_model)

    # Resúmenes agregados
    summary_segment = d_share_model.groupby(["SegmentoSUV", "Mes"], dropna=False, as_index=False)["Unidades"].sum()

    summary_brand = d_share_model.groupby(["SegmentoSUV", "Marca", "Mes"], dropna=False, as_index=False)["Unidades"].sum()

    summary_model = d_share_model.groupby(["SegmentoSUV", "Marca", "Modelo", "Mes"], dropna=False, as_index=False)["Unidades"].sum()

    # Último mes disponible
    if d_share_model["Mes"].notna().any():
        last_month = d_share_model["Mes"].max()
        prev_month = (last_month - pd.offsets.MonthBegin(1))
    else:
        last_month = None
        prev_month = None

    # Top movers (modelo) del último mes (por % MoM)
    top_movers = pd.DataFrame()
    if last_month is not None:
        last_rows = d_share_model[d_share_model["Mes"].eq(last_month) & d_share_model["Version"].isna()]
        # Si no hay versiones, nos quedamos con nivel modelo agregando versiones
        if last_rows.empty:
            tmp = d_share_model[d_share_model["Mes"].eq(last_month)].groupby(["SegmentoSUV", "Marca", "Modelo"], as_index=False)["Unidades"].sum()
            prev = d_share_model[d_share_model["Mes"].eq(prev_month)].groupby(["SegmentoSUV", "Marca", "Modelo"], as_index=False)["Unidades"].sum().rename(columns={"Unidades": "Unidades_prev"})
            tmp = tmp.merge(prev, on=["SegmentoSUV", "Marca", "Modelo"], how="left")
            tmp["MoM_%"] = np.where(tmp["Unidades_prev"].fillna(0) == 0, np.nan, (tmp["Unidades"]/tmp["Unidades_prev"]) - 1)
            top_movers = tmp.sort_values("MoM_%", ascending=False).head(15)
        else:
            top_movers = last_rows.sort_values("MoM_%", ascending=False).head(15)

    out = {
        "detail": d_share_model,
        "summary_segment": summary_segment,
        "summary_brand": summary_brand,
        "summary_model": summary_model,
        "top_movers": top_movers,
        "last_month": pd.to_datetime(last_month) if last_month is not None else None,
        "prev_month": pd.to_datetime(prev_month) if prev_month is not None else None,
    }
    return out


def competitor_vs_volvo(df_long: pd.DataFrame, brands_selected: List[str] | None = None) -> pd.DataFrame:
    """Comparativa Volvo vs competidores por segmento y mes."""
    d = df_long.copy()
    d = d[d["SegmentoSUV"].isin(list("BCDE"))]
    d["Marca"] = d["Marca"].fillna("(Sin marca)")
    # Si no se pasan marcas, tomar top 6 por volumen (excl. Volvo)
    vol_by_brand = d.groupby("Marca", as_index=False)["Unidades"].sum().sort_values("Unidades", ascending=False)
    if brands_selected:
        comp = [b for b in brands_selected if b.lower() != "volvo"]
    else:
        comp = [b for b in vol_by_brand["Marca"].tolist() if b.lower() != "volvo"][:6]

    d_sub = d[d["Marca"].str.lower().isin(["volvo"] + [b.lower() for b in comp])]

    g = d_sub.groupby(["SegmentoSUV", "Marca", "Mes"], as_index=False)["Unidades"].sum()

    # share por segmento-mes
    g = build_share(g, ["SegmentoSUV", "Marca"], "Unidades", "Share_Segmento", ["SegmentoSUV", "Mes"])
    return g


# =========================
# CARGA DE ARCHIVO
# =========================
with st.sidebar:
    st.header("⚙️ Configuración")
    up = st.file_uploader("Sube tu archivo (.xlsx, .xls o .csv)", type=["xlsx", "xls", "csv"])

    sheet_name = None
    if up is not None and up.name.lower().endswith((".xlsx", ".xls")):
        try:
            xl = pd.ExcelFile(up)
            sheet_name = st.selectbox("Hoja de Excel", options=xl.sheet_names)
            df_raw = xl.parse(sheet_name)
        except Exception:
            st.warning("No se pudo leer el Excel, intentando lectura directa...")
            df_raw = pd.read_excel(up)
    elif up is not None and up.name.lower().endswith(".csv"):
        sep = st.selectbox("Separador CSV", [",", ";", "\t"], index=0)
        df_raw = pd.read_csv(up, sep=sep)
    else:
        df_raw = None

    if df_raw is not None:
        st.success(f"Archivo cargado: {up.name}")
        with st.expander("Ver primeras filas"):
            st.dataframe(df_raw.head(20))

        cols = df_raw.columns.tolist()
        # Sugerencias de mapeo
        brand_guess = pick_first_present(["Marca", "Brand", "Make"], cols)
        model_guess = pick_first_present(["Modelo", "Model"], cols)
        version_guess = pick_first_present(["Version", "Versión", "Trim", "Variante"], cols)
        segment_guess = pick_first_present(["Segmento", "Segment", "Clase", "Categoria", "Categoría"], cols)

        st.subheader("Mapeo de columnas")
        brand_col = st.selectbox("Columna de Marca (opcional)", [None] + cols, index=(0 if brand_guess is None else ([None] + cols).index(brand_guess)))
        model_col = st.selectbox("Columna de Modelo (obligatoria)", cols, index=(0 if model_guess is None else cols.index(model_guess)))
        version_col = st.selectbox("Columna de Versión (opcional)", [None] + cols, index=(0 if version_guess is None else ([None] + cols).index(version_guess)))
        segment_col = st.selectbox("Columna de Segmento (SUV)", cols, index=(0 if segment_guess is None else cols.index(segment_guess)))

        config = {"brand": brand_col, "model": model_col, "version": version_col, "segment": segment_col}

        run_btn = st.button("Procesar análisis", type="primary")
    else:
        run_btn = False

# =========================
# ANÁLISIS
# =========================
# 1) Cuando presionas "Procesar análisis", guardamos los datos en session_state
if run_btn and df_raw is not None:
    try:
        df_long, months_sorted = tidy_data(df_raw, config)
        st.session_state["df_long"] = df_long
        st.session_state["months"] = months_sorted
    except Exception as e:
        st.error(f"Error preparando datos: {e}")

# 2) Si ya tenemos datos preparados, mostramos filtros, KPIs, tabs, descargas e insights
if "df_long" in st.session_state:
    df_long = st.session_state["df_long"]
    months_sorted = st.session_state["months"]

    # ---- Filtros ----
    with st.sidebar:
        st.markdown("---")
        seg_filter = st.multiselect("Segmentos SUV a incluir", options=list("BCDE"), default=list("BCDE"))
        marcas = sorted([m for m in df_long["Marca"].dropna().unique()])
        marca_filter = st.multiselect("Filtrar marcas (opcional)", options=marcas, default=marcas)
        modelos_all = sorted(df_long["Modelo"].dropna().unique())
        modelo_filter = st.multiselect("Filtrar modelos (opcional)", options=modelos_all, default=modelos_all[:min(25, len(modelos_all))])

    dflt = df_long[df_long["SegmentoSUV"].isin(seg_filter) & df_long["Modelo"].isin(modelo_filter)]
    if marca_filter:
        dflt = dflt[dflt["Marca"].isin(marca_filter)]

    # ---- KPIs ----
    total_units = dflt.groupby("Mes", as_index=False)["Unidades"].sum()
    last_m = total_units["Mes"].max() if not total_units.empty else None

    last_units: float = 0.0
    prev_units: float = np.nan
    mom_pct_all: float = np.nan
    delta_all: float = np.nan
    prev_m = None

    if last_m is not None:
        prev_m = (pd.Timestamp(last_m) - pd.offsets.MonthBegin(1))
        last_units = float(total_units.loc[total_units["Mes"].eq(last_m), "Unidades"].sum())
        if (total_units["Mes"].eq(prev_m)).any():
            prev_units = float(total_units.loc[total_units["Mes"].eq(prev_m), "Unidades"].sum())
        delta_all = last_units - (0 if np.isnan(prev_units) else prev_units)
        if not np.isnan(prev_units) and prev_units != 0:
            mom_pct_all = (last_units / prev_units) - 1.0

    kpi1, kpi2, kpi3 = st.columns(3)
    with kpi1:
        st.metric("Unidades último mes", f"{0 if last_m is None else int(last_units):,}")
    with kpi2:
        st.metric("Δ unidades vs mes previo", f"{0 if np.isnan(delta_all) else int(delta_all):,}")
    with kpi3:
        st.metric("Crecimiento MoM %", "N/A" if np.isnan(mom_pct_all) else f"{mom_pct_all:.1%}")

    # ---- Cálculo de resultados para tabs ----
    results = compute_metrics(dflt)
    detail = results["detail"].copy()
    summary_segment = results["summary_segment"].copy()
    summary_brand = results["summary_brand"].copy()
    summary_model = results["summary_model"].copy()
    top_movers = results["top_movers"].copy()
    last_month = results["last_month"]
    prev_month = results.get("prev_month", None)

    # =========================
    # PESTAÑAS DE ANÁLISIS
    # =========================
    t1, t2, t3, t4, t5 = st.tabs([
        "📊 Ventas por mes",
        "📈 Crecimiento MoM",
        "🧩 Mix por segmento y modelo",
        "🔎 Detalle por modelo/versión",
        "🆚 Volvo vs competidores",
    ])

    # --- T1: Ventas por mes ---
    with t1:
        st.subheader("Ventas totales por mes (todos los segmentos seleccionados)")
        if not summary_segment.empty:
            fig = px.bar(summary_segment, x="Mes", y="Unidades", color="SegmentoSUV", barmode="stack")
            fig.update_layout(legend_title_text="Segmento")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("---")
        st.subheader("Top modelos por volumen — último mes")
        if last_month is not None:
            top_models_last = summary_model[summary_model["Mes"].eq(last_month)].sort_values("Unidades", ascending=False).head(20)
            if not top_models_last.empty:
                fig2 = px.bar(top_models_last, x="Modelo", y="Unidades", color="SegmentoSUV")
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("No hay datos del último mes para modelos.")
        else:
            st.info("Aún no es posible identificar el último mes.")

    # --- T2: Crecimiento MoM ---
    with t2:
        st.subheader("Crecimiento mes a mes por modelo")
        growth_model = detail.groupby(["SegmentoSUV", "Marca", "Modelo", "Mes"], as_index=False)[["Unidades", "Delta"]].sum()
        growth_model.sort_values(["Modelo", "Mes"], inplace=True)
        growth_model["Unidades_prev"] = growth_model.groupby(["Modelo"])['Unidades'].shift(1)
        growth_model["MoM_%"] = np.where(growth_model["Unidades_prev"].fillna(0) == 0, np.nan, (growth_model["Unidades"] / growth_model["Unidades_prev"]) - 1)
        pivot = growth_model.pivot_table(index=["Marca", "Modelo"], columns="Mes", values="MoM_%", aggfunc="mean")
        if not pivot.empty:
            fig = px.imshow(pivot, aspect="auto", origin="lower", color_continuous_scale="RdBu", zmin=-1, zmax=1)
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("---")
        st.subheader("Top movers (mayor MoM % último mes)")
        if not top_movers.empty:
            st.dataframe(top_movers)
        else:
            st.info("No es posible calcular Top movers (faltan dos meses consecutivos).")

    # --- T3: Mix ---
    with t3:
        st.subheader("Participación por segmento → marca → modelo (último mes)")
        if last_month is not None:
            seg_month = detail[detail["Mes"].eq(last_month)].copy()
            by = seg_month.groupby(["SegmentoSUV", "Marca", "Modelo"], as_index=False)["Unidades"].sum()
            # share dentro de cada segmento (último mes)
            totals = by.groupby(["SegmentoSUV"], as_index=False)["Unidades"].sum().rename(columns={"Unidades": "__den__"})
            by = by.merge(totals, on="SegmentoSUV", how="left")
            by["ShareSeg"] = np.where(by["__den__"].eq(0), np.nan, by["Unidades"]/by["__den__"])
            by.drop(columns="__den__", inplace=True)
            fig = px.bar(by, x="SegmentoSUV", y="ShareSeg", color="Marca", hover_data=["Modelo"], barmode="stack")
            fig.update_yaxes(tickformat=",.0%")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("---")
        st.subheader("Mix de versiones dentro de un modelo (último mes)")
        modelos_disp = sorted(detail["Modelo"].dropna().unique())
        modelo_sel = st.selectbox("Elige un modelo", options=modelos_disp)
        if last_month is not None and modelo_sel:
            ver = detail[(detail["Modelo"].eq(modelo_sel)) & (detail["Mes"].eq(last_month))].groupby(["Version"], as_index=False)["Unidades"].sum()
            ver = ver[ver["Version"].notna()]
            if not ver.empty:
                total_ver = float(ver["Unidades"].sum())
                ver["ShareVersion"] = np.where(total_ver == 0, np.nan, ver["Unidades"]/total_ver)
                ver = ver.sort_values("ShareVersion", ascending=False)
                figv = px.pie(ver, names="Version", values="Unidades", hole=0.4)
                st.plotly_chart(figv, use_container_width=True)
            else:
                st.info("No hay columna de Versión o no hay datos por versión para el último mes.")

    # --- T4: Detalle ---
    with t4:
        st.subheader("Evolución de ventas por modelo")
        if not summary_model.empty:
            model_sel = st.selectbox("Selecciona un modelo para ver su evolución", options=sorted(summary_model["Modelo"].unique()))
            dfm = summary_model[summary_model["Modelo"].eq(model_sel)]
            fig = px.line(dfm, x="Mes", y="Unidades", color="SegmentoSUV")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("---")
            st.subheader("Evolución por versión (si aplica)")
            versions = detail[detail["Modelo"].eq(model_sel)]["Version"].dropna().unique()
            if versions.size > 0:
                ver_sel = st.multiselect("Elige versiones", options=sorted(versions), default=list(versions)[:5])
                dfv = detail[(detail["Modelo"].eq(model_sel)) & (detail["Version"].isin(ver_sel))]
                figv = px.line(dfv, x="Mes", y="Unidades", color="Version")
                st.plotly_chart(figv, use_container_width=True)
            else:
                st.info("No hay versiones disponibles o no fueron mapeadas.")
        else:
            st.info("No hay datos para el detalle por modelo.")

    # --- T5: Volvo vs competidores ---
    with t5:
        st.subheader("Participación por segmento: Volvo vs competidores")
        marcas_all = sorted([m for m in df_long["Marca"].dropna().unique() if str(m).lower() != "volvo"])
        comp_pick = st.multiselect("Elige competidores (si no eliges, se toman los 6 con mayor volumen)", options=marcas_all)
        comp_df = competitor_vs_volvo(dflt, comp_pick)
        if not comp_df.empty:
            fig = px.bar(comp_df, x="Mes", y="Unidades", color="Marca", facet_row="SegmentoSUV", barmode="stack")
            st.plotly_chart(fig, use_container_width=True)
            comp_df2 = comp_df.copy()
            comp_df2["Grupo"] = np.where(comp_df2["Marca"].str.lower().eq("volvo"), "Volvo", "Competidores")
            comp2 = comp_df2.groupby(["SegmentoSUV", "Mes", "Grupo"], as_index=False)["Unidades"].sum()
            totals = comp2.groupby(["SegmentoSUV", "Mes"], as_index=False)["Unidades"].sum().rename(columns={"Unidades": "__den__"})
            comp2 = comp2.merge(totals, on=["SegmentoSUV", "Mes"], how="left")
            comp2["Share"] = np.where(comp2["__den__"].eq(0), np.nan, comp2["Unidades"]/comp2["__den__"])
            comp2.drop(columns="__den__", inplace=True)
            fig2 = px.area(comp2, x="Mes", y="Share", color="Grupo", facet_row="SegmentoSUV", groupnorm="fraction")
            fig2.update_yaxes(tickformat=",.0%")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No hay datos suficientes para la comparativa.")

    # =========================
    # DESCARGAS + INSIGHTS
    # =========================
    st.markdown("---")
    st.subheader("Descargar resultados")

    import io
    def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            for name, df in sheets.items():
                safe_name = name[:31]
                df.to_excel(writer, sheet_name=safe_name, index=False)
        return output.getvalue()

    excel_bytes = to_excel_bytes({
        "detail_long": detail,
        "summary_segment": summary_segment,
        "summary_brand": summary_brand,
        "summary_model": summary_model,
        "top_movers": top_movers,
    })

    st.download_button(
        label="💾 Descargar Excel (múltiples hojas)",
        data=excel_bytes,
        file_name="sales_mix_resultados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.download_button(
        label="⬇️ Descargar CSV consolidado (detail_long)",
        data=detail.to_csv(index=False).encode("utf-8"),
        file_name="sales_mix_detail_long.csv",
        mime="text/csv",
    )

    st.markdown("---")
    st.subheader("🧠 Insights automáticos")
    if last_month is not None and not summary_model.empty:
        last_tot = summary_segment.loc[summary_segment["Mes"].eq(last_month), "Unidades"].sum()
        best_model = summary_model[summary_model["Mes"].eq(last_month)].sort_values("Unidades", ascending=False).head(1)
        if not best_model.empty:
            bm = best_model.iloc[0]
            st.write(
                f"**{int(last_tot):,}** unidades totales en **{last_month.strftime('%b %Y')}**; "
                f"modelo líder: **{bm['Modelo']}** de **{bm['Marca'] or '(Marca no definida)'}** en **{bm['SegmentoSUV']}** con **{int(bm['Unidades']):,}**."
            )
        tmp = competitor_vs_volvo(dflt)
        if not tmp.empty:
            tmp_last = tmp[tmp["Mes"].eq(last_month)]
            volvo_last = tmp_last[tmp_last["Marca"].str.lower().eq("volvo")]
            if not volvo_last.empty:
                volvo_share = volvo_last.groupby("SegmentoSUV")["Share_Segmento"].mean().mean()
                st.write(f"Participación promedio de **Volvo** (promedio simple entre segmentos) en el último mes: **{volvo_share*100:.1f}%**.")
else:
    st.info("🗂️ Sube tu archivo y mapea columnas en la barra lateral; luego presiona **Procesar análisis**.")
