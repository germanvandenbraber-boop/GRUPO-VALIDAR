"""
Motor de Análisis — Lotería de Río Negro
Red de Subagencias · Dashboard de Rendimiento v3.0

Cambios v3.0 (corrección de exactitud de datos):
  - CRÍTICO: Importe_Ajustado ahora es estrictamente solo sobre Concepto='Validos'
    Los Premios Pagados (negativos) NO se incluyen en recaudación ni en Importe_Ajustado.
  - Fuente de datos: hoja 'Ventas' del Excel (ya pre-filtrada y limpia).
    No se usa la tabla dinámica ni Datos_Crudos directamente.
  - Telebingo Patagonia: factor 0.5 aplicado sobre Validos únicamente.
  - Sub=0 (agencia madre) y subagencias de baja excluidas en la fuente.

Uso:  streamlit run dashboard_lrn.py
Req:  pip install streamlit pandas openpyxl plotly
Arch: BASE_DATOS_LRN.xlsx en la misma carpeta
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings

warnings.filterwarnings("ignore")

# ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Motor de Análisis — LRN",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

AG13_C = "#7EC924"
AG33_C = "#F5C800"
LRN_C  = "#2E8B3A"

JUEGO_C = {
    "Quiniela":            "#2A78D6",
    "Quini Express":       "#1BAF7A",
    "Quini6":              "#EDA100",
    "Loto Plus":           "#6A52C7",
    "Patagonia Telebingo": "#E34948",
    "Patagonia MiniBingo": "#E87BA4",
    "TeleKino TJ":         "#2E8B3A",
    "Pozo de la Quiniela": "#EB6834",
    "Lotería":             "#888780",
    "Brinco":              "#BA7517",
}

MESES_NOM = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

st.markdown("""
<style>
  [data-testid="stSidebar"]       { background:#1A1A18; }
  [data-testid="stSidebar"] *     { color:#fff !important; }
  [data-testid="stSidebar"] label { color:#7EC924 !important; font-weight:600; }
  [data-testid="stMetric"]        { background:#f7f7f5; border-radius:10px;
                                    padding:12px 16px; border:0.5px solid #e0e0d8; }
  [data-testid="stMetricLabel"]   { font-size:11px !important; color:#52514e !important;
                                    text-transform:uppercase; letter-spacing:.05em; }
  [data-testid="stMetricValue"]   { font-size:22px !important; font-weight:500 !important; }
  h1 { color:#7EC924 !important; }
  h2 { color:#1A1A18 !important; border-bottom:2px solid #7EC924; padding-bottom:4px; }
  h3 { color:#4A7A14 !important; }
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = "BASE_DATOS_LRN.xlsx"


# ══════════════════════════════════════════════════════════════════════════════
# CARGA DE DATOS
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=300)
def cargar_datos(filepath: str):
    """
    Lee la hoja 'Ventas' del Excel.
    Esa hoja ya contiene SOLO registros Concepto='Validos' + Cierre M5
    + sin Sub=0 + sin bajas. Telebingo ya tiene factor 0.5 aplicado.
    No se mezclan Premios Pagados — recaudación siempre positiva y exacta.
    """
    # ── Ventas ────────────────────────────────────────────────────────────────
    ventas = pd.read_excel(filepath, sheet_name="Ventas", dtype=str)
    ventas.columns = [str(c).strip() for c in ventas.columns]

    # Descartar filas donde Año no sea numérico
    ventas = ventas[pd.to_numeric(ventas.get("Año", pd.Series()), errors="coerce").notna()].copy()

    for col in ["Año", "Mes", "Tickets", "Importe_Bruto", "Importe_Ajustado"]:
        ventas[col] = pd.to_numeric(ventas[col], errors="coerce").fillna(0)

    ventas["Año"]             = ventas["Año"].astype(int)
    ventas["Mes"]             = ventas["Mes"].astype(int)
    ventas["Tickets"]         = ventas["Tickets"].astype(int)
    ventas["Agencia_Madre"]   = ventas["Agencia_Madre"].astype(str).str.strip().str.zfill(2)
    ventas["Num_Subagente"]   = ventas["Num_Subagente"].astype(str).str.strip().str.zfill(2)
    ventas["Codigo_Completo"] = ventas["Codigo_Completo"].astype(str).str.strip()
    ventas["Denominacion"]    = ventas["Denominacion"].astype(str).str.strip()
    ventas["Juego"]           = ventas["Juego"].astype(str).str.strip()

    ventas["Periodo"]       = ventas["Año"].astype(str) + "-" + ventas["Mes"].astype(str).str.zfill(2)
    ventas["Mes_Label"]     = ventas["Mes"].apply(lambda m: MESES_NOM[m] if 1 <= m <= 12 else str(m))
    ventas["Periodo_Label"] = ventas["Mes_Label"] + " " + ventas["Año"].astype(str)

    # ── Objetivos ─────────────────────────────────────────────────────────────
    objvs = pd.read_excel(filepath, sheet_name="Objetivos", dtype=str)
    objvs.columns = [str(c).strip() for c in objvs.columns]
    objvs = objvs[pd.to_numeric(objvs.get("Año", pd.Series()), errors="coerce").notna()].copy()

    for col in ["Año", "Mes", "Objetivo_Pesos"]:
        objvs[col] = pd.to_numeric(objvs[col], errors="coerce").fillna(0)

    objvs["Año"]             = objvs["Año"].astype(int)
    objvs["Mes"]             = objvs["Mes"].astype(int)
    objvs["Agencia_Madre"]   = objvs["Agencia_Madre"].astype(str).str.strip().str.zfill(2)
    objvs["Codigo_Completo"] = objvs["Codigo_Completo"].astype(str).str.strip()
    objvs["Juego"]           = objvs["Juego"].astype(str).str.strip()
    # Solo columnas necesarias para evitar conflictos en merge
    objvs = objvs[["Año", "Mes", "Agencia_Madre", "Codigo_Completo",
                    "Juego", "Objetivo_Pesos"]].copy()
    objvs["Periodo"] = objvs["Año"].astype(str) + "-" + objvs["Mes"].astype(str).str.zfill(2)

    # ── Diccionario maestro de nombres (sin merge, sin conflictos) ─────────────
    maestro_nombres = (
        ventas[["Codigo_Completo", "Denominacion"]]
        .drop_duplicates(subset=["Codigo_Completo"])
        .set_index("Codigo_Completo")["Denominacion"]
        .to_dict()
    )

    return ventas, objvs, maestro_nombres


def fmt_pesos(val: float) -> str:
    if val >= 1_000_000:
        return f"${val / 1_000_000:.1f}M"
    if val >= 1_000:
        return f"${val / 1_000:.0f}K"
    return f"${val:,.0f}".replace(",", ".")


def sub_label(codigo: str, nombres: dict, max_len: int = 22) -> str:
    num = codigo[2:].lstrip("0") or "0"
    den = str(nombres.get(codigo, codigo))[:max_len]
    return f"{num} — {den}"


# ── Cargar ────────────────────────────────────────────────────────────────────
try:
    ventas, objvs, maestro_nombres = cargar_datos(EXCEL_FILE)
except FileNotFoundError:
    st.error(f"❌ No se encontró **{EXCEL_FILE}**.")
    st.info("Colocá `BASE_DATOS_LRN.xlsx` en la misma carpeta que este script y recargá.")
    st.stop()
except Exception as exc:
    st.error(f"❌ Error al cargar datos: {exc}")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — FILTROS
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"""
    <div style='text-align:center;padding:8px 0 16px'>
      <div style='width:42px;height:42px;border-radius:50%;background:{LRN_C};
           display:inline-flex;align-items:center;justify-content:center;font-size:20px'>🍀</div>
      <div style='font-size:13px;font-weight:600;margin-top:6px'>Lotería de Río Negro</div>
      <div style='font-size:10px;color:#aaa'>Motor de Análisis · Subagencias</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("**📅 PERÍODO**")
    anios = sorted(ventas["Año"].unique(), reverse=True)
    anio_sel = st.selectbox("Año", anios)

    meses_disp = sorted(ventas[ventas["Año"] == anio_sel]["Mes"].unique())
    mes_sel = st.selectbox(
        "Mes", meses_disp,
        format_func=lambda m: f"{MESES_NOM[m] if m <= 12 else m} {anio_sel}",
        index=len(meses_disp) - 1,
    )
    st.markdown("---")

    st.markdown("**🏢 AGENCIA MADRE**")
    agencias = ["Todas"] + sorted(ventas["Agencia_Madre"].unique().tolist())
    ag_sel = st.selectbox("Agencia", agencias)

    st.markdown("**👤 SUBAGENTE**")
    if ag_sel == "Todas":
        subs_df = ventas[["Codigo_Completo"]].drop_duplicates()
    else:
        subs_df = ventas[ventas["Agencia_Madre"] == ag_sel][["Codigo_Completo"]].drop_duplicates()
    subs_df = subs_df.sort_values("Codigo_Completo")

    opc_sub = {"Todos": "Todos"}
    for _, r in subs_df.iterrows():
        opc_sub[r["Codigo_Completo"]] = sub_label(r["Codigo_Completo"], maestro_nombres)

    sub_sel = st.selectbox("Subagente", list(opc_sub.keys()), format_func=lambda k: opc_sub[k])

    st.markdown("---")
    st.markdown("**🎮 JUEGO**")
    juegos_disp = ["Todos"] + sorted(ventas["Juego"].unique().tolist())
    juego_sel = st.selectbox("Juego", juegos_disp)

    st.markdown("---")
    st.caption(
        f"Fuente: BASE_DATOS_LRN.xlsx\n"
        f"Solo Concepto=Validos + M5\n"
        f"Telebingo ajustado al 50%\n"
        f"Período: {ventas['Periodo'].min()} → {ventas['Periodo'].max()}"
    )


# ── FILTRADO ──────────────────────────────────────────────────────────────────
def filt(df, anio, mes, ag, sub, juego):
    d = df[(df["Año"] == anio) & (df["Mes"] == mes)].copy()
    if ag    != "Todas": d = d[d["Agencia_Madre"]   == ag]
    if sub   != "Todos": d = d[d["Codigo_Completo"] == sub]
    if juego != "Todos": d = d[d["Juego"]           == juego]
    return d


def filt_hist(df, ag, sub, juego):
    d = df.copy()
    if ag    != "Todas": d = d[d["Agencia_Madre"]   == ag]
    if sub   != "Todos": d = d[d["Codigo_Completo"] == sub]
    if juego != "Todos": d = d[d["Juego"]           == juego]
    return d


mes_ant = (anio_sel, mes_sel - 1) if mes_sel > 1 else (anio_sel - 1, 12)

d_act  = filt(ventas, anio_sel, mes_sel, ag_sel, sub_sel, juego_sel)
d_ant  = filt(ventas, mes_ant[0], mes_ant[1], ag_sel, sub_sel, juego_sel)
d_hist = filt_hist(ventas, ag_sel, sub_sel, juego_sel)

d_obj = objvs[(objvs["Año"] == anio_sel) & (objvs["Mes"] == mes_sel)].copy()
if ag_sel    != "Todas": d_obj = d_obj[d_obj["Agencia_Madre"]   == ag_sel]
if sub_sel   != "Todos": d_obj = d_obj[d_obj["Codigo_Completo"] == sub_sel]
if juego_sel != "Todos": d_obj = d_obj[d_obj["Juego"]           == juego_sel]


# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
col_t, col_b = st.columns([3, 1])
with col_t:
    mes_str = MESES_NOM[mes_sel] if mes_sel <= 12 else str(mes_sel)
    st.markdown("# 📊 Motor de Análisis — Subagencias LRN")
    st.markdown(
        f"**{mes_str} {anio_sel}** · "
        f"{'Agencia ' + ag_sel if ag_sel != 'Todas' else 'Todas las agencias'} · "
        f"{opc_sub.get(sub_sel, 'Todos')} · {juego_sel}"
    )
with col_b:
    st.markdown(f"""
    <div style='display:flex;gap:8px;justify-content:flex-end;margin-top:16px'>
      <span style='background:{AG13_C};color:#1a3a05;padding:4px 12px;
            border-radius:20px;font-weight:700;font-size:12px'>AG. 13</span>
      <span style='background:{AG33_C};color:#3a2e00;padding:4px 12px;
            border-radius:20px;font-weight:700;font-size:12px'>AG. 33</span>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")


# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
t1, t2, t3, t4, t5 = st.tabs([
    "📋 Resumen",
    "📈 Evolución",
    "🎮 Por Juego",
    "🏆 Ranking",
    "🎯 Objetivos",
])


# ──────────────────────────────────────────────────────────────────────────────
# TAB 1 — RESUMEN
# ──────────────────────────────────────────────────────────────────────────────
with t1:
    rec_a  = d_act["Importe_Ajustado"].sum()
    rec_p  = d_ant["Importe_Ajustado"].sum()
    tkt_a  = d_act["Tickets"].sum()
    tkt_p  = d_ant["Tickets"].sum()
    tkt_pr = rec_a / tkt_a if tkt_a > 0 else 0
    var_r  = ((rec_a - rec_p) / rec_p * 100) if rec_p > 0 else 0.0
    var_t  = ((tkt_a - tkt_p) / tkt_p * 100) if tkt_p > 0 else 0.0
    n_subs = d_act["Codigo_Completo"].nunique()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Recaudación Ajustada",     fmt_pesos(rec_a),
              f"{var_r:+.1f}% vs mes ant.")
    k2.metric("Tickets Emitidos",          f"{int(tkt_a):,}".replace(",", "."),
              f"{var_t:+.1f}% vs mes ant.")
    k3.metric("Ticket Promedio",           fmt_pesos(tkt_pr))
    k4.metric("Subagentes con actividad",  str(n_subs))

    st.markdown("---")
    cl, cr = st.columns(2)

    with cl:
        st.markdown("### Por Agencia Madre")
        ag_d = d_act.groupby("Agencia_Madre")["Importe_Ajustado"].sum().reset_index()
        if not ag_d.empty:
            fig = go.Figure(go.Pie(
                labels=[f"Agencia {a}" for a in ag_d["Agencia_Madre"]],
                values=ag_d["Importe_Ajustado"], hole=0.45,
                marker_colors=[AG13_C if a == "13" else AG33_C for a in ag_d["Agencia_Madre"]],
                textinfo="label+percent", textfont_size=12,
            ))
            fig.update_layout(height=300, margin=dict(t=10, b=10, l=10, r=10),
                              showlegend=False, paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sin datos para el período.")

    with cr:
        st.markdown("### Por Juego")
        j_d = (d_act.groupby("Juego")["Importe_Ajustado"].sum()
               .reset_index().sort_values("Importe_Ajustado"))
        if not j_d.empty:
            fig = go.Figure(go.Bar(
                x=j_d["Importe_Ajustado"], y=j_d["Juego"], orientation="h",
                marker_color=[JUEGO_C.get(j, "#888") for j in j_d["Juego"]],
                text=[fmt_pesos(v) for v in j_d["Importe_Ajustado"]],
                textposition="outside",
            ))
            fig.update_layout(height=300, margin=dict(t=10, b=10, l=10, r=90),
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            fig.update_xaxes(visible=False)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sin datos para el período.")

    st.markdown("### Tabla de Subagentes")
    t_sub = (d_act.groupby(["Codigo_Completo", "Agencia_Madre"])
             .agg(Recaudacion=("Importe_Ajustado", "sum"), Tickets=("Tickets", "sum"))
             .reset_index().sort_values("Recaudacion", ascending=False))
    if not t_sub.empty:
        t_sub["Tkt_Prom"]  = (t_sub["Recaudacion"] / t_sub["Tickets"].replace(0, 1)).round(0)
        t_sub["Subagente"] = t_sub["Codigo_Completo"].apply(
            lambda c: sub_label(c, maestro_nombres, 30))
        ts = t_sub[["Subagente", "Agencia_Madre", "Recaudacion", "Tickets", "Tkt_Prom"]].copy()
        ts.columns = ["Subagente", "Ag.", "Recaudación $", "Tickets", "Tkt Prom $"]
        ts["Recaudación $"] = ts["Recaudación $"].apply(lambda x: f"${x:,.0f}".replace(",", "."))
        ts["Tickets"]       = ts["Tickets"].apply(lambda x: f"{int(x):,}".replace(",", "."))
        ts["Tkt Prom $"]    = ts["Tkt Prom $"].apply(lambda x: f"${x:,.0f}".replace(",", "."))
        st.dataframe(ts, use_container_width=True, hide_index=True)
    else:
        st.info("Sin datos.")


# ──────────────────────────────────────────────────────────────────────────────
# TAB 2 — EVOLUCIÓN
# ──────────────────────────────────────────────────────────────────────────────
with t2:
    evol = (d_hist.groupby(["Periodo", "Periodo_Label", "Año", "Mes"])
            .agg(Recaudacion=("Importe_Ajustado", "sum"), Tickets=("Tickets", "sum"))
            .reset_index().sort_values("Periodo"))
    evol["Tkt_Prom"] = (evol["Recaudacion"] / evol["Tickets"].replace(0, 1)).fillna(0)

    if evol.empty:
        st.info("Sin datos históricos con los filtros seleccionados.")
    else:
        st.markdown("### Evolución — Recaudación y Tickets")
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(
            x=evol["Periodo_Label"], y=evol["Recaudacion"],
            name="Recaudación $", marker_color=AG13_C,
            text=[fmt_pesos(v) for v in evol["Recaudacion"]],
            textposition="outside", textfont_size=9,
        ), secondary_y=False)
        fig.add_trace(go.Scatter(
            x=evol["Periodo_Label"], y=evol["Tickets"],
            name="Tickets", mode="lines+markers",
            line=dict(color=AG33_C, width=2.5),
            marker=dict(size=7, color=AG33_C),
        ), secondary_y=True)
        fig.update_layout(
            height=380, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(t=30, b=10, l=10, r=10),
        )
        fig.update_yaxes(title_text="Recaudación ($)", secondary_y=False,
                         showgrid=True, gridcolor="#e8e8e4")
        fig.update_yaxes(title_text="Tickets", secondary_y=True, showgrid=False)
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("### Ticket Promedio por Período")
        fig2 = go.Figure(go.Scatter(
            x=evol["Periodo_Label"], y=evol["Tkt_Prom"],
            mode="lines+markers+text",
            line=dict(color=LRN_C, width=2.5),
            marker=dict(size=8, color=LRN_C),
            text=[fmt_pesos(v) for v in evol["Tkt_Prom"]],
            textposition="top center", textfont_size=9,
            fill="tozeroy", fillcolor="rgba(46,139,58,0.08)",
        ))
        fig2.update_layout(height=260, plot_bgcolor="rgba(0,0,0,0)",
                           paper_bgcolor="rgba(0,0,0,0)",
                           yaxis=dict(showgrid=True, gridcolor="#e8e8e4"),
                           margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(fig2, use_container_width=True)

        if sub_sel == "Todos" and ag_sel != "Todas":
            st.markdown("### Evolución por Subagente")
            ev_sub = d_hist.copy()
            ev_sub["Lbl"] = ev_sub["Codigo_Completo"].apply(
                lambda c: sub_label(c, maestro_nombres, 14))
            ev_sub = (ev_sub.groupby(["Periodo", "Periodo_Label", "Lbl"])
                      .agg(Recaudacion=("Importe_Ajustado", "sum"))
                      .reset_index().sort_values("Periodo"))
            if not ev_sub.empty:
                fig3 = px.line(ev_sub, x="Periodo_Label", y="Recaudacion",
                               color="Lbl", markers=True,
                               color_discrete_sequence=px.colors.qualitative.Set2)
                fig3.update_layout(height=350, plot_bgcolor="rgba(0,0,0,0)",
                                   paper_bgcolor="rgba(0,0,0,0)",
                                   legend_title="Subagente",
                                   margin=dict(t=10, b=10, l=10, r=10))
                st.plotly_chart(fig3, use_container_width=True)


# ──────────────────────────────────────────────────────────────────────────────
# TAB 3 — POR JUEGO
# ──────────────────────────────────────────────────────────────────────────────
with t3:
    j_act = (d_act.groupby("Juego")
             .agg(Recaudacion=("Importe_Ajustado", "sum"), Tickets=("Tickets", "sum"))
             .reset_index())
    j_act["Tkt_Prom"] = (j_act["Recaudacion"] / j_act["Tickets"].replace(0, 1)).fillna(0)
    total_rec         = j_act["Recaudacion"].sum()
    j_act["Pct"]      = ((j_act["Recaudacion"] / total_rec * 100)
                         if total_rec > 0 else 0).round(1)

    if j_act.empty:
        st.info("Sin datos.")
    else:
        st.markdown("### Distribución por Juego")
        cl2, cr2 = st.columns(2)
        with cl2:
            fig = go.Figure(go.Pie(
                labels=j_act["Juego"], values=j_act["Recaudacion"], hole=0.4,
                marker_colors=[JUEGO_C.get(j, "#888") for j in j_act["Juego"]],
                textinfo="label+percent", textfont_size=10, sort=True,
            ))
            fig.update_layout(height=320, margin=dict(t=20, b=0, l=0, r=0),
                              showlegend=False, paper_bgcolor="rgba(0,0,0,0)",
                              title=dict(text="Recaudación", font_size=12))
            st.plotly_chart(fig, use_container_width=True)
        with cr2:
            fig2 = go.Figure(go.Pie(
                labels=j_act["Juego"], values=j_act["Tickets"], hole=0.4,
                marker_colors=[JUEGO_C.get(j, "#888") for j in j_act["Juego"]],
                textinfo="label+percent", textfont_size=10, sort=True,
            ))
            fig2.update_layout(height=320, margin=dict(t=20, b=0, l=0, r=0),
                               showlegend=False, paper_bgcolor="rgba(0,0,0,0)",
                               title=dict(text="Tickets", font_size=12))
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("### Tabla por Juego")
        js = j_act.sort_values("Recaudacion", ascending=False)[
            ["Juego", "Recaudacion", "Tickets", "Tkt_Prom", "Pct"]].copy()
        js.columns = ["Juego", "Recaudación $", "Tickets", "Tkt Prom $", "% Total"]
        js["Recaudación $"] = js["Recaudación $"].apply(lambda x: f"${x:,.0f}".replace(",", "."))
        js["Tickets"]       = js["Tickets"].apply(lambda x: f"{int(x):,}".replace(",", "."))
        js["Tkt Prom $"]    = js["Tkt Prom $"].apply(lambda x: f"${x:,.0f}".replace(",", "."))
        js["% Total"]       = js["% Total"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(js, use_container_width=True, hide_index=True)

        st.markdown("### Evolución por Juego — Histórico")
        j_hist = (d_hist.groupby(["Periodo", "Periodo_Label", "Juego"])
                  .agg(Recaudacion=("Importe_Ajustado", "sum"))
                  .reset_index().sort_values("Periodo"))
        if not j_hist.empty:
            fig3 = px.area(j_hist, x="Periodo_Label", y="Recaudacion",
                           color="Juego", color_discrete_map=JUEGO_C)
            fig3.update_layout(height=350, plot_bgcolor="rgba(0,0,0,0)",
                               paper_bgcolor="rgba(0,0,0,0)",
                               yaxis=dict(showgrid=True, gridcolor="#e8e8e4"),
                               legend_title="Juego",
                               margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig3, use_container_width=True)


# ──────────────────────────────────────────────────────────────────────────────
# TAB 4 — RANKING
# ──────────────────────────────────────────────────────────────────────────────
with t4:
    rank = (d_act.groupby(["Codigo_Completo", "Agencia_Madre"])
            .agg(Recaudacion=("Importe_Ajustado", "sum"), Tickets=("Tickets", "sum"))
            .reset_index()
            .sort_values("Recaudacion", ascending=False)
            .reset_index(drop=True))
    rank["Tkt_Prom"] = (rank["Recaudacion"] / rank["Tickets"].replace(0, 1)).fillna(0)
    rank["Lbl"]      = rank["Codigo_Completo"].apply(
        lambda c: sub_label(c, maestro_nombres, 22))

    if rank.empty:
        st.info("Sin datos.")
    else:
        m1, m2, m3 = st.columns(3)
        top_r = rank.iloc[0]
        top_t = rank.sort_values("Tickets", ascending=False).iloc[0]
        top_p = rank.sort_values("Tkt_Prom", ascending=False).iloc[0]
        m1.metric("🥇 Líder Recaudación",  top_r["Lbl"][:26], fmt_pesos(top_r["Recaudacion"]))
        m2.metric("🎟️ Líder Tickets",      top_t["Lbl"][:26],
                  f"{int(top_t['Tickets']):,}".replace(",", "."))
        m3.metric("💰 Mayor Tkt Promedio", top_p["Lbl"][:26], fmt_pesos(top_p["Tkt_Prom"]))

        st.markdown("---")
        st.markdown("### Ranking por Recaudación")
        fig = go.Figure(go.Bar(
            y=rank["Lbl"], x=rank["Recaudacion"], orientation="h",
            marker_color=[AG13_C if ag == "13" else AG33_C for ag in rank["Agencia_Madre"]],
            text=[fmt_pesos(v) for v in rank["Recaudacion"]],
            textposition="outside",
        ))
        fig.update_layout(
            height=max(280, len(rank) * 28), yaxis=dict(autorange="reversed"),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            xaxis=dict(showgrid=True, gridcolor="#e8e8e4"),
            margin=dict(t=10, b=10, l=10, r=100),
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("### Mapa de Posicionamiento — Recaudación vs. Tickets")
        fig2 = px.scatter(
            rank, x="Tickets", y="Recaudacion",
            size="Tkt_Prom", color="Agencia_Madre",
            color_discrete_map={"13": AG13_C, "33": AG33_C},
            hover_name="Lbl",
            hover_data={"Tkt_Prom": True, "Recaudacion": True, "Tickets": True},
            text="Lbl", size_max=60,
        )
        fig2.update_traces(textposition="top center", textfont_size=8)
        fig2.update_layout(
            height=440, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            xaxis=dict(showgrid=True, gridcolor="#e8e8e4"),
            yaxis=dict(showgrid=True, gridcolor="#e8e8e4"),
            margin=dict(t=10, b=10, l=10, r=10), legend_title="Agencia",
        )
        st.plotly_chart(fig2, use_container_width=True)

        if sub_sel == "Todos":
            st.markdown("### Heatmap — Recaudación por Subagente y Juego")
            d_lbl = d_act.copy()
            d_lbl["Lbl2"] = d_lbl["Codigo_Completo"].apply(
                lambda c: sub_label(c, maestro_nombres, 18))
            heat = (d_lbl.groupby(["Lbl2", "Juego"])["Importe_Ajustado"]
                    .sum().unstack(fill_value=0))
            if not heat.empty:
                fig_h = px.imshow(heat, color_continuous_scale="YlGn",
                                  aspect="auto", text_auto=".2s")
                fig_h.update_layout(
                    height=max(280, len(heat) * 24),
                    paper_bgcolor="rgba(0,0,0,0)",
                    margin=dict(t=10, b=10, l=10, r=10),
                )
                st.plotly_chart(fig_h, use_container_width=True)


# ──────────────────────────────────────────────────────────────────────────────
# TAB 5 — OBJETIVOS
# ──────────────────────────────────────────────────────────────────────────────
with t5:
    mes_str_t5 = MESES_NOM[mes_sel] if mes_sel <= 12 else str(mes_sel)
    st.markdown(f"### Cumplimiento de Objetivos — {mes_str_t5} {anio_sel}")

    if d_obj.empty or d_obj["Objetivo_Pesos"].sum() == 0:
        st.warning("⚠️ No hay objetivos cargados para el período seleccionado.")
        st.info(
            "Para cargar objetivos: abrir `BASE_DATOS_LRN.xlsx` → "
            "hoja **Objetivos** → completar las celdas azules (Objetivo_Pesos) "
            "→ guardar el archivo y recargar el dashboard."
        )
    else:
        # Real solo de Codigo_Completo + Juego (sin Denominacion para evitar conflicto en merge)
        real_agg = (d_act.groupby(["Codigo_Completo", "Juego"])["Importe_Ajustado"]
                    .sum().reset_index()
                    .rename(columns={"Importe_Ajustado": "Real"}))

        # Merge limpio (d_obj solo tiene las columnas mínimas necesarias)
        mg = d_obj.merge(real_agg, on=["Codigo_Completo", "Juego"], how="left")
        mg["Real"]   = mg["Real"].fillna(0)
        mg["Cumpl"]  = ((mg["Real"] / mg["Objetivo_Pesos"].replace(0, 1)) * 100).round(1)
        mg["Brecha"] = mg["Real"] - mg["Objetivo_Pesos"]
        mg["Estado"] = mg["Cumpl"].apply(
            lambda x: "✅ Cumplido" if x >= 100 else ("⚠️ En riesgo" if x >= 70 else "❌ Bajo"))
        # Enriquecer con nombre desde diccionario (sin merge adicional)
        mg["Lbl"] = mg["Codigo_Completo"].apply(
            lambda c: sub_label(c, maestro_nombres, 20))

        n_tot  = len(mg[mg["Objetivo_Pesos"] > 0])
        n_ok   = len(mg[mg["Estado"] == "✅ Cumplido"])
        n_rsg  = len(mg[mg["Estado"] == "⚠️ En riesgo"])
        n_baj  = len(mg[mg["Estado"] == "❌ Bajo"])
        pct_gl = mg[mg["Objetivo_Pesos"] > 0]["Cumpl"].mean() if n_tot > 0 else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Cumplimiento Global",    f"{pct_gl:.1f}%")
        k2.metric("✅ Cumplidos",            str(n_ok))
        k3.metric("⚠️ En Riesgo (70-99%)", str(n_rsg))
        k4.metric("❌ Bajo (<70%)",          str(n_baj))
        st.markdown("---")

        sub_cum = (mg[mg["Objetivo_Pesos"] > 0]
                   .groupby(["Codigo_Completo", "Agencia_Madre"])
                   .agg(Real=("Real", "sum"), Objetivo=("Objetivo_Pesos", "sum"))
                   .reset_index())
        sub_cum["Pct"] = ((sub_cum["Real"] / sub_cum["Objetivo"].replace(0, 1)) * 100).round(1)
        sub_cum["Lbl"] = sub_cum["Codigo_Completo"].apply(
            lambda c: sub_label(c, maestro_nombres, 20))
        sub_cum = sub_cum.sort_values("Pct", ascending=False)

        st.markdown("### % Cumplimiento por Subagente")
        bar_c = [AG13_C if p >= 100 else (AG33_C if p >= 70 else "#E34948")
                 for p in sub_cum["Pct"]]
        fig = go.Figure(go.Bar(
            y=sub_cum["Lbl"], x=sub_cum["Pct"], orientation="h",
            marker_color=bar_c,
            text=[f"{p:.1f}%" for p in sub_cum["Pct"]], textposition="outside",
        ))
        fig.add_vline(x=100, line_dash="dash", line_color="#276221", line_width=2,
                      annotation_text="Meta 100%", annotation_position="top right")
        fig.add_vline(x=70,  line_dash="dot",  line_color="#9C6500", line_width=1,
                      annotation_text="Alerta 70%", annotation_position="bottom right")
        max_x = sub_cum["Pct"].max() * 1.18 if not sub_cum.empty else 120
        fig.update_layout(
            height=max(280, len(sub_cum) * 30),
            yaxis=dict(autorange="reversed"),
            xaxis=dict(range=[0, max(max_x, 120)], showgrid=True, gridcolor="#e8e8e4"),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=10, b=10, l=10, r=90),
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("### Detalle — Real vs. Objetivo por Subagente y Juego")
        tab_o = (mg[mg["Objetivo_Pesos"] > 0][
            ["Lbl", "Juego", "Real", "Objetivo_Pesos", "Cumpl", "Brecha", "Estado"]]
            .sort_values("Cumpl", ascending=False).copy())
        tab_o.columns = ["Subagente", "Juego", "Real $", "Objetivo $", "Cumpl. %",
                         "Brecha $", "Estado"]
        tab_o["Real $"]     = tab_o["Real $"].apply(lambda x: f"${x:,.0f}".replace(",", "."))
        tab_o["Objetivo $"] = tab_o["Objetivo $"].apply(lambda x: f"${x:,.0f}".replace(",", "."))
        tab_o["Brecha $"]   = tab_o["Brecha $"].apply(
            lambda x: f"${x:,.0f}".replace(",", ".")
            if x >= 0 else f"(${abs(x):,.0f})".replace(",", "."))
        tab_o["Cumpl. %"]   = tab_o["Cumpl. %"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(tab_o, use_container_width=True, hide_index=True)


# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<div style='text-align:center;font-size:10px;color:#aaa'>"
    "Motor de Análisis v3.0 · Lotería de Río Negro · Agencias 13 y 33 · "
    "Solo Concepto=Validos + Cierre M5 · Telebingo 50% · Bajas excluidas"
    "</div>",
    unsafe_allow_html=True,
)
