"""
Frontend Streamlit — Dashboard ISI Ultra Évolué + Pointage
Lancement : streamlit run frontend/app.py
Backend   : uvicorn backend.main:app --reload --port 8000
"""
from __future__ import annotations

import os, sys
from datetime import date
import datetime as dt

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import frontend.api_client as api

# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ISI — Dashboard Suivi",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEPT_OPTIONS = ["IAID", "KM", "DRS", "GI"]
MOIS_COLS    = ["Oct", "Nov", "Déc", "Jan", "Fév", "Mars", "Avril", "Mai", "Juin", "Juil", "Août"]
MOIS_API     = ["oct","nov","dec","jan","fev","mars","avril","mai","juin","juil","aout"]
MOIS_MAP_API = dict(zip(MOIS_API, MOIS_COLS))

DEFAULT_THRESHOLDS = {"taux_vert": 0.90, "taux_orange": 0.60, "ecart_critique": -6}

# ══════════════════════════════════════════════════════════════════════════════
# CSS — THÈME DARK EXÉCUTIF IDENTIQUE À L'ORIGINAL
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700;800&display=swap');

:root, html, body, .stApp { color-scheme: dark !important; }
html, body, .stApp { zoom:1!important; font-size:16px!important; font-family:'Space Grotesk',sans-serif!important; }

:root{
  --bg:#080E1A; --bg2:#0D1526; --card:#111B2E; --card2:#162035;
  --text:#E8EDF5; --muted:#7A90B8; --line:rgba(255,255,255,0.07);
  --blue:#1F6FEB; --blue2:#2A80FF; --blue3:#5AA2FF; --blue-glow:rgba(31,111,235,0.25);
  --ok:#00C97A; --ok-glow:rgba(0,201,122,0.20);
  --warn:#FF9500; --warn-glow:rgba(255,149,0,0.20);
  --bad:#FF3B3B; --bad-glow:rgba(255,59,59,0.20);
}
html,body,.stApp{ background:radial-gradient(ellipse at 10% 0%,#0D1E3A 0%,var(--bg) 60%)!important; }
body,.stApp,p,span,div,label{ color:var(--text)!important; -webkit-font-smoothing:antialiased; }
h1,h2,h3,h4,h5{ color:var(--blue3)!important; font-weight:850!important; }
a,a:visited{ color:var(--blue3)!important; text-decoration:none; }
.stCaption,small{ color:var(--muted)!important; font-weight:650; }
.block-container{ padding-top:.5rem!important; padding-bottom:4rem!important; }
header[data-testid="stHeader"]{ background:transparent!important; border-bottom:0!important; }
::-webkit-scrollbar{ width:6px; height:6px; }
::-webkit-scrollbar-track{ background:var(--bg2); }
::-webkit-scrollbar-thumb{ background:rgba(90,162,255,0.25); border-radius:4px; }

/* SIDEBAR */
section[data-testid="stSidebar"]{ background:var(--bg2)!important; border-right:1px solid var(--line); }
.sidebar-card{
  background:var(--card); border:1px solid var(--line); border-radius:16px;
  padding:12px; margin-bottom:10px; box-shadow:0 4px 20px rgba(0,0,0,.40);
}

/* HEADER */
.iaid-header{
  background:linear-gradient(135deg,#0A1E44 0%,#0F2860 40%,#1A3A80 100%);
  border:1px solid rgba(90,162,255,.20); color:#FFF!important; padding:20px 26px;
  border-radius:20px; box-shadow:0 0 40px rgba(31,111,235,.15),0 20px 48px rgba(0,0,0,.50),inset 0 1px 0 rgba(255,255,255,.08);
  margin-bottom:16px; position:relative; overflow:hidden;
}
.iaid-header::before{
  content:""; position:absolute; top:-60px; right:-60px; width:200px; height:200px;
  background:radial-gradient(circle,rgba(90,162,255,.15) 0%,transparent 70%); pointer-events:none;
}
.iaid-header *{ color:#FFF!important; text-shadow:0 1px 3px rgba(0,0,0,.40); }
.iaid-htitle{ font-size:20px; font-weight:950; letter-spacing:-.3px; }
.iaid-hsub{ font-size:13px; opacity:.80; margin-top:4px; }
.iaid-badges{ margin-top:12px; display:flex; gap:8px; flex-wrap:wrap; }
.iaid-badge{
  background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.18);
  padding:5px 12px; border-radius:999px; font-size:12px; font-weight:850;
}

/* KPI CARDS */
.kpi-grid{ display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:14px; margin:14px 0; }
.kpi{
  background:linear-gradient(135deg,var(--card) 0%,var(--card2) 100%);
  border:1px solid var(--line); border-radius:20px; padding:18px 18px 14px 18px;
  box-shadow:0 8px 32px rgba(0,0,0,.40); position:relative; overflow:hidden;
  transition:transform .18s ease,box-shadow .18s ease;
}
.kpi:hover{ transform:translateY(-3px); box-shadow:0 16px 48px rgba(0,0,0,.55); }
.kpi:before{
  content:""; position:absolute; top:0; left:0; width:100%; height:3px;
  background:var(--blue); border-radius:20px 20px 0 0;
}
.kpi:after{
  content:""; position:absolute; bottom:-30px; right:-30px; width:90px; height:90px;
  background:radial-gradient(circle,var(--blue-glow) 0%,transparent 70%); pointer-events:none;
}
.kpi-title{ font-size:11px; font-weight:800; text-transform:uppercase; letter-spacing:.6px; color:var(--muted)!important; margin-bottom:8px; }
.kpi-value{ font-size:28px; font-weight:950; letter-spacing:-.5px; margin-top:2px; line-height:1; }
.kpi-good:before{ background:linear-gradient(90deg,var(--ok),#00FFA3); }
.kpi-good:after { background:radial-gradient(circle,var(--ok-glow) 0%,transparent 70%); }
.kpi-good .kpi-value{ color:var(--ok)!important; }
.kpi-warn:before{ background:linear-gradient(90deg,var(--warn),#FFD060); }
.kpi-warn:after { background:radial-gradient(circle,var(--warn-glow) 0%,transparent 70%); }
.kpi-warn .kpi-value{ color:var(--warn)!important; }
.kpi-bad:before{ background:linear-gradient(90deg,var(--bad),#FF7070); }
.kpi-bad:after { background:radial-gradient(circle,var(--bad-glow) 0%,transparent 70%); }
.kpi-bad .kpi-value{ color:var(--bad)!important; }

/* TABS */
button[data-baseweb="tab"]{
  background:var(--card)!important; color:var(--muted)!important; border-radius:999px!important;
  padding:10px 16px!important; font-weight:800!important; border:1px solid var(--line)!important;
  transition:all .15s ease!important;
}
button[data-baseweb="tab"]:hover{ color:var(--text)!important; border-color:rgba(90,162,255,.30)!important; }
button[data-baseweb="tab"][aria-selected="true"]{
  background:rgba(31,111,235,.18)!important; color:var(--blue3)!important;
  border:1px solid rgba(90,162,255,.40)!important; box-shadow:0 0 14px var(--blue-glow)!important;
}

/* INPUTS */
div[data-baseweb="input"]>div, div[data-baseweb="select"]>div{
  background:var(--card2)!important; border:1px solid var(--line)!important; border-radius:14px!important;
}
div[data-baseweb="input"] input, div[data-baseweb="select"] *{ color:var(--text)!important; font-weight:700!important; }

/* BUTTONS */
.stButton>button{
  background:linear-gradient(135deg,var(--blue),var(--blue2))!important;
  color:#FFF!important; border:none!important; border-radius:14px!important;
  font-weight:800!important; padding:10px 20px!important;
  box-shadow:0 4px 16px rgba(31,111,235,.35)!important;
  transition:all .15s ease!important;
}
.stButton>button:hover{
  transform:translateY(-2px)!important;
  box-shadow:0 8px 24px rgba(31,111,235,.50)!important;
}

/* POINTAGE CARD */
.pointage-card{
  background:linear-gradient(135deg,var(--card) 0%,var(--card2) 100%);
  border:1px solid rgba(90,162,255,.20); border-radius:20px; padding:20px;
  margin:8px 0; box-shadow:0 8px 24px rgba(0,0,0,.35);
}
.pointage-badge-ok{ background:#16a34a; color:#fff; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:700; }
.pointage-badge-warn{ background:#d97706; color:#fff; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:700; }
.pointage-badge-bad{ background:#dc2626; color:#fff; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:700; }
.pointage-badge-nd{ background:#4b5563; color:#fff; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:700; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    # Logo
    logo_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "logo_iaid.jpg")
    if os.path.exists(logo_path):
        st.markdown('<div class="sidebar-logo-wrap">', unsafe_allow_html=True)
        st.image(logo_path, width=150)
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown("<h2 style='text-align:center'>📊 ISI</h2>", unsafe_allow_html=True)

    # Connexion API
    try:
        api.health()
        st.success("✅ API connectée")
    except Exception:
        st.error("❌ API non joignable\n\n`uvicorn backend.main:app --reload`")
        st.stop()

    st.markdown('<div class="sidebar-card">', unsafe_allow_html=True)
    dept = st.selectbox("🏛️ Département", DEPT_OPTIONS)
    st.markdown('</div>', unsafe_allow_html=True)

    # Filtres
    st.markdown('<div class="sidebar-card">', unsafe_allow_html=True)
    st.markdown("**🔎 Filtres**")
    classes_data = api.get_classes(dept)
    classe_noms = ["Toutes"] + [c["nom"] for c in classes_data]
    filtre_classe = st.selectbox("Classe", classe_noms)
    filtre_statut = st.multiselect("Statut", ["Non démarré", "En cours", "Terminé"],
                                    default=["Non démarré", "En cours", "Terminé"])
    st.markdown('</div>', unsafe_allow_html=True)

    # Seuils
    st.markdown('<div class="sidebar-card">', unsafe_allow_html=True)
    st.markdown("**⚙️ Seuils d'alerte**")
    taux_vert   = st.slider("Seuil Vert (%)",   50, 100, 90, 5) / 100
    taux_orange = st.slider("Seuil Orange (%)", 10, 95,  60, 5) / 100
    ecart_crit  = st.slider("Écart critique (h)", -40, 0, -6, 1)
    st.markdown('</div>', unsafe_allow_html=True)

    if st.button("🔄 Rafraîchir"):
        st.cache_data.clear()
        st.rerun()

    st.caption(f"API : {os.environ.get('API_BASE_URL','http://localhost:8000')}")

thresholds = {"taux_vert": taux_vert, "taux_orange": taux_orange, "ecart_critique": ecart_crit}


# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT DONNÉES
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=30, show_spinner=False)
def load_resume(dept: str) -> pd.DataFrame:
    data = api.get_resume(dept)
    if not data:
        return pd.DataFrame()
    df = pd.DataFrame(data)
    # Renommer pour correspondre à l'original
    rename = {
        "matiere": "Matière", "classe": "Classe", "responsable": "Responsable",
        "semestre": "Semestre", "email": "Email",
        "vhp": "VHP", "vhr": "VHR", "taux": "Taux", "ecart": "Écart",
        "statut": "Statut_auto", "enseignement_id": "_rowid",
        "oct":"Oct","nov":"Nov","dec":"Déc","jan":"Jan","fev":"Fév",
        "mars":"Mars","avril":"Avril","mai":"Mai","juin":"Juin","juil":"Juil","aout":"Août",
    }
    df = df.rename(columns=rename)
    # Colonnes manquantes
    for col in MOIS_COLS:
        if col not in df.columns:
            df[col] = 0.0
    for col in ["Semestre","Email","Observations","Début prévu","Fin prévue"]:
        if col not in df.columns:
            df[col] = ""
    if "Statut" not in df.columns:
        df["Statut"] = df["Statut_auto"]
    df["Écart"] = df["Écart"].fillna(0)
    df["Taux"]  = df["Taux"].fillna(0)
    df["VHR"]   = df["VHR"].fillna(0)
    df["Écart_num"] = df["Écart"]
    return df


def statut_badge(s: str) -> str:
    m = {"Terminé": "🟢 Terminé", "En cours": "🟡 En cours", "Non démarré": "⚫ Non démarré"}
    return m.get(s, s)


with st.spinner("Chargement des données…"):
    df_all = load_resume(dept)

# ── Header ────────────────────────────────────────────────────────────────────
dept_labels = {
    "IAID": "IA & Ingénierie des Données (IAID)",
    "KM":   "Directions des Études — KM",
    "DRS":  "Réseaux et Systèmes (DRS)",
    "GI":   "Génie Informatique (GI)",
}
nb_classes = len(classes_data)
nb_ens     = len(df_all) if not df_all.empty else 0

st.markdown(f"""
<div class="iaid-header">
  <div class="iaid-htitle">📊 Dashboard Suivi des Enseignements</div>
  <div class="iaid-hsub">{dept_labels.get(dept, dept)} — Pilotage mensuel</div>
  <div class="iaid-badges">
    <span class="iaid-badge">🏛️ {dept}</span>
    <span class="iaid-badge">📚 {nb_classes} classe(s)</span>
    <span class="iaid-badge">📖 {nb_ens} matière(s)</span>
    <span class="iaid-badge">🗓️ {dt.date.today().strftime('%d/%m/%Y')}</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Filtre actif ──────────────────────────────────────────────────────────────
if df_all.empty:
    st.info("Aucune donnée. Configure des classes et enseignements dans l'onglet **Administration**.")
    filtered = pd.DataFrame()
else:
    filtered = df_all.copy()
    if filtre_classe != "Toutes":
        filtered = filtered[filtered["Classe"] == filtre_classe]
    if filtre_statut:
        filtered = filtered[filtered["Statut_auto"].isin(filtre_statut)]

# ══════════════════════════════════════════════════════════════════════════════
# ONGLETS
# ══════════════════════════════════════════════════════════════════════════════
(tab_overview, tab_classes, tab_matieres, tab_enseignants,
 tab_mensuel, tab_alertes, tab_pointage, tab_admin) = st.tabs([
    "🌐 Vue globale", "🏫 Par classe", "📖 Par matière",
    "👤 Par enseignant", "📅 Analyse mensuelle", "🚨 Alertes",
    "✏️ Pointage", "⚙️ Administration",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — VUE GLOBALE
# ══════════════════════════════════════════════════════════════════════════════
with tab_overview:
    if filtered.empty:
        st.info("Aucune donnée avec les filtres actuels.")
    else:
        total       = len(filtered)
        taux_moy    = float(filtered["Taux"].mean() * 100) if total else 0.0
        nb_term     = int((filtered["Statut_auto"] == "Terminé").sum())
        nb_enc      = int((filtered["Statut_auto"] == "En cours").sum())
        nb_nd       = int((filtered["Statut_auto"] == "Non démarré").sum())
        retard_tot  = float(filtered.loc[filtered["Écart"] < 0, "Écart"].sum())
        total_vhp   = float(filtered["VHP"].sum())
        total_vhr   = float(filtered["VHR"].sum())

        retard_class = "kpi-bad" if retard_tot < 0 else "kpi-warn"

        st.markdown(f"""
        <div class="kpi-grid">
          <div class="kpi kpi-good">
            <div class="kpi-title">Matières</div>
            <div class="kpi-value">{total}</div>
          </div>
          <div class="kpi kpi-warn">
            <div class="kpi-title">Taux moyen</div>
            <div class="kpi-value">{taux_moy:.1f}%</div>
          </div>
          <div class="kpi kpi-good">
            <div class="kpi-title">Terminées</div>
            <div class="kpi-value">{nb_term}</div>
          </div>
          <div class="kpi kpi-warn">
            <div class="kpi-title">En cours</div>
            <div class="kpi-value">{nb_enc}</div>
          </div>
          <div class="kpi {retard_class}">
            <div class="kpi-title">Retard cumulé (h)</div>
            <div class="kpi-value">{retard_tot:.0f}</div>
          </div>
        </div>
        <div class="kpi-grid" style="grid-template-columns:repeat(3,minmax(0,1fr));margin-top:0;">
          <div class="kpi kpi-good">
            <div class="kpi-title">VHP total</div>
            <div class="kpi-value">{total_vhp:.0f}h</div>
          </div>
          <div class="kpi kpi-warn">
            <div class="kpi-title">VHR total</div>
            <div class="kpi-value">{total_vhr:.0f}h</div>
          </div>
          <div class="kpi kpi-warn">
            <div class="kpi-title">Non démarrées</div>
            <div class="kpi-value">{nb_nd}</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.divider()
        col1, col2 = st.columns(2)

        with col1:
            st.write("### Avancement par classe")
            g = filtered.groupby("Classe")["Taux"].mean().sort_values(ascending=False).reset_index()
            g["Taux (%)"] = (g["Taux"] * 100).round(1)
            st.dataframe(g[["Classe","Taux (%)"]], use_container_width=True,
                column_config={"Taux (%)": st.column_config.ProgressColumn(
                    "Taux (%)", min_value=0, max_value=100, format="%.1f%%")})

        with col2:
            st.write("### Répartition des statuts")
            stat = filtered["Statut_auto"].value_counts().reset_index()
            stat.columns = ["Statut","Nombre"]
            fig = px.pie(stat, names="Statut", values="Nombre",
                         color="Statut", color_discrete_map={
                            "Terminé":"#00C97A","En cours":"#FF9500","Non démarré":"#6b7280"})
            fig.update_layout(height=320, margin=dict(l=0,r=0,t=30,b=0),
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                              font_color="#E8EDF5")
            st.plotly_chart(fig, use_container_width=True)

        st.write("### VHP vs VHR par classe")
        bar_data = filtered.groupby("Classe")[["VHP","VHR"]].sum().reset_index()
        fig2 = px.bar(bar_data, x="Classe", y=["VHP","VHR"], barmode="group",
                      color_discrete_map={"VHP":"#1F6FEB","VHR":"#00C97A"})
        fig2.update_layout(height=350, margin=dict(l=0,r=0,t=30,b=0),
                           paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                           font_color="#E8EDF5", legend_title_text="")
        st.plotly_chart(fig2, use_container_width=True)

        st.write("### Top retards (Écart le plus négatif)")
        top_ret = filtered[["Classe","Matière","VHP","VHR","Écart","Taux","Statut_auto"]].copy()
        top_ret = top_ret.sort_values("Écart").head(20)
        top_ret["Taux (%)"] = (top_ret["Taux"]*100).round(1)
        top_ret["Statut"] = top_ret["Statut_auto"].apply(statut_badge)
        st.dataframe(top_ret[["Classe","Matière","VHP","VHR","Écart","Taux (%)","Statut"]],
            use_container_width=True, height=420,
            column_config={
                "Taux (%)": st.column_config.ProgressColumn("Taux (%)", min_value=0, max_value=100, format="%.1f%%"),
                "Écart": st.column_config.NumberColumn("Écart (h)", format="%.0f"),
                "VHP": st.column_config.NumberColumn("VHP", format="%.0f"),
                "VHR": st.column_config.NumberColumn("VHR", format="%.0f"),
            })


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — PAR CLASSE
# ══════════════════════════════════════════════════════════════════════════════
with tab_classes:
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        classes_filtered = sorted(filtered["Classe"].unique().tolist())
        colA, colB = st.columns([2,1])

        with colB:
            cls1 = st.selectbox("Classe A", classes_filtered, index=0, key="c1")
            cls2 = st.selectbox("Classe B", classes_filtered,
                                index=min(1, len(classes_filtered)-1), key="c2")

        with colA:
            st.write("### Synthèse par classe")
            synth = filtered.groupby("Classe").agg(
                Matieres=("Matière","count"),
                Taux_moy=("Taux","mean"),
                VHP_total=("VHP","sum"),
                VHR_total=("VHR","sum"),
                Retard_h=("Écart", lambda s: float(s[s<0].sum())),
                Terminees=("Statut_auto", lambda s: int((s=="Terminé").sum())),
                Non_demarre=("Statut_auto", lambda s: int((s=="Non démarré").sum())),
            ).reset_index()
            synth["Taux (%)"] = (synth["Taux_moy"]*100).round(1)
            st.dataframe(synth[["Classe","Matieres","Taux (%)","VHP_total","VHR_total","Retard_h","Terminees","Non_demarre"]],
                use_container_width=True,
                column_config={
                    "Taux (%)": st.column_config.ProgressColumn("Taux (%)", min_value=0, max_value=100, format="%.1f%%"),
                    "Retard_h":  st.column_config.NumberColumn("Retard (h)", format="%.0f"),
                    "VHP_total": st.column_config.NumberColumn("VHP total", format="%.0f"),
                    "VHR_total": st.column_config.NumberColumn("VHR total", format="%.0f"),
                })

        st.divider()
        st.write(f"### Comparaison — {cls1} vs {cls2}")
        A = filtered[filtered["Classe"]==cls1]
        B = filtered[filtered["Classe"]==cls2]

        def kpis(one):
            return {"Matières": len(one),
                    "Taux moyen": f"{float(one['Taux'].mean()*100):.1f}%" if len(one) else "0%",
                    "Retard (h)": float(one.loc[one["Écart"]<0,"Écart"].sum()),
                    "Non démarré": int((one["Statut_auto"]=="Non démarré").sum())}

        kA, kB = kpis(A), kpis(B)
        comp = pd.DataFrame({"Indicateur": list(kA.keys()), cls1: list(kA.values()), cls2: list(kB.values())})
        st.dataframe(comp, use_container_width=True, hide_index=True)

        for cls, df_cls in [(cls1, A), (cls2, B)]:
            st.write(f"### Top retards — {cls}")
            t = df_cls.sort_values("Écart").head(15)[["Matière","VHP","VHR","Écart","Taux","Statut_auto"]].copy()
            t["Taux (%)"] = (t["Taux"]*100).round(1)
            t["Statut"] = t["Statut_auto"].apply(statut_badge)
            st.dataframe(t[["Matière","VHP","VHR","Écart","Taux (%)","Statut"]],
                use_container_width=True,
                column_config={
                    "Taux (%)": st.column_config.ProgressColumn("Taux (%)", min_value=0, max_value=100, format="%.1f%%"),
                    "Écart": st.column_config.NumberColumn("Écart (h)", format="%.0f"),
                    "VHP": st.column_config.NumberColumn("VHP", format="%.0f"),
                    "VHR": st.column_config.NumberColumn("VHR", format="%.0f"),
                })


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — PAR MATIÈRE
# ══════════════════════════════════════════════════════════════════════════════
with tab_matieres:
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        st.subheader("Analyse par matière (toutes classes)")
        mat = filtered.groupby("Matière").agg(
            Classes=("Classe","nunique"), VHP=("VHP","sum"), VHR=("VHR","sum"),
            Taux=("Taux","mean"), Retard=("Écart", lambda s: float(s[s<0].sum())),
            Non_demarre=("Statut_auto", lambda s: int((s=="Non démarré").sum())),
        ).reset_index()
        mat["Taux (%)"] = (mat["Taux"]*100).round(1)
        st.dataframe(mat.sort_values(["Taux (%)","Retard"], ascending=[True,True]),
                     use_container_width=True)

        st.write("### Matières en alerte")
        al = mat[(mat["Taux"] < thresholds["taux_orange"]) | (mat["Retard"] <= thresholds["ecart_critique"])].copy()
        if al.empty:
            st.success("✅ Aucune matière en alerte selon les seuils.")
        else:
            st.dataframe(al.sort_values("Taux (%)").head(30), use_container_width=True)

        st.write("### Top matières — taux de réalisation")
        fig = px.bar(mat.sort_values("Taux (%)").head(20),
                     x="Matière", y="Taux (%)", color="Taux (%)",
                     color_continuous_scale=["#FF3B3B","#FF9500","#00C97A"])
        fig.update_layout(height=400, margin=dict(l=0,r=0,t=30,b=80),
                          paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                          font_color="#E8EDF5", xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — PAR ENSEIGNANT
# ══════════════════════════════════════════════════════════════════════════════
with tab_enseignants:
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        st.subheader("Suivi par enseignant — retards & charge")
        tmp = filtered.copy()
        tmp["Responsable"] = tmp["Responsable"].astype(str).replace({"nan":"","None":""}).fillna("").str.strip()
        tmp["Responsable"] = tmp["Responsable"].replace({"": "⚠️ Non affecté"})

        synth_r = tmp.groupby("Responsable").agg(
            Matieres=("Matière","count"), Classes=("Classe","nunique"),
            VHP_total=("VHP","sum"), VHR_total=("VHR","sum"),
            Taux_moy=("Taux","mean"),
            Retard_h=("Écart", lambda s: float(s[s<0].sum())),
            Non_demarre=("Statut_auto", lambda s: int((s=="Non démarré").sum())),
            En_cours=("Statut_auto", lambda s: int((s=="En cours").sum())),
            Termine=("Statut_auto", lambda s: int((s=="Terminé").sum())),
        ).reset_index()
        synth_r["Taux (%)"] = (synth_r["Taux_moy"]*100).round(1)
        synth_r = synth_r.sort_values(["Retard_h","Taux (%)"], ascending=[True,True])

        st.write("### Synthèse par enseignant")
        st.dataframe(synth_r[["Responsable","Matieres","Classes","Taux (%)","VHP_total","VHR_total","Retard_h","Non_demarre","En_cours","Termine"]],
            use_container_width=True,
            column_config={
                "Taux (%)":  st.column_config.ProgressColumn("Taux (%)", min_value=0, max_value=100, format="%.1f%%"),
                "Retard_h":  st.column_config.NumberColumn("Retard (h)", format="%.0f"),
                "VHP_total": st.column_config.NumberColumn("VHP total", format="%.0f"),
                "VHR_total": st.column_config.NumberColumn("VHR total", format="%.0f"),
            })

        st.divider()
        st.write("### Top retards — par enseignant")
        top_n = st.slider("Nb lignes", 10, 200, 50, 10, key="top_ens")
        top_e = tmp[tmp["Écart"]<0].sort_values("Écart").head(top_n)[
            ["Responsable","Classe","Matière","Semestre","VHP","VHR","Écart","Taux","Statut_auto"]].copy()
        st.dataframe(top_e, use_container_width=True,
            column_config={
                "Taux": st.column_config.ProgressColumn("Taux", min_value=0, max_value=1, format="%.0f%%"),
                "Écart": st.column_config.NumberColumn("Écart (h)", format="%.0f"),
            })

        st.divider()
        st.write("### Non démarrés — par enseignant")
        nd = tmp[tmp["Statut_auto"]=="Non démarré"].groupby("Responsable").size().sort_values(ascending=False)
        if nd.empty:
            st.success("✅ Aucun 'Non démarré'")
        else:
            fig = px.bar(nd.reset_index(), x="Responsable", y=0,
                         color_discrete_sequence=["#FF9500"])
            fig.update_layout(height=350, yaxis_title="Nb matières",
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                              font_color="#E8EDF5")
            st.plotly_chart(fig, use_container_width=True)

        st.divider()
        st.write("### Charge par enseignant — VHP vs VHR")
        charge = tmp.groupby("Responsable").agg(VHP_total=("VHP","sum"), VHR_total=("VHR","sum")).reset_index()
        charge["Écart_total"] = charge["VHR_total"] - charge["VHP_total"]
        st.dataframe(charge.sort_values("Écart_total"),
            use_container_width=True,
            column_config={
                "VHP_total": st.column_config.NumberColumn("VHP total", format="%.0f"),
                "VHR_total": st.column_config.NumberColumn("VHR total", format="%.0f"),
                "Écart_total": st.column_config.NumberColumn("Écart (h)", format="%.0f"),
            })


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — ANALYSE MENSUELLE
# ══════════════════════════════════════════════════════════════════════════════
with tab_mensuel:
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        st.subheader("Analyse mensuelle — heures réalisées & tendances")

        mois_dispo = [m for m in MOIS_COLS if m in filtered.columns]
        monthly = filtered[mois_dispo].sum()

        st.write("### Heures totales par mois")
        fig = px.line(x=mois_dispo, y=monthly.values, markers=True,
                      labels={"x":"Mois","y":"Heures"},
                      color_discrete_sequence=["#1F6FEB"])
        fig.update_traces(line_width=3, marker_size=8)
        fig.update_layout(height=350, paper_bgcolor="rgba(0,0,0,0)",
                          plot_bgcolor="rgba(0,0,0,0)", font_color="#E8EDF5")
        st.plotly_chart(fig, use_container_width=True)

        st.write("### Matrice Classe × Mois (heures)")
        pivot = filtered.groupby("Classe")[mois_dispo].sum()
        st.dataframe(pivot, use_container_width=True)

        if pivot.shape[0] * pivot.shape[1] <= 250:
            st.write("### Heatmap — heures par classe et par mois")
            fig2 = px.imshow(pivot.values, x=mois_dispo, y=pivot.index.tolist(),
                             aspect="auto", color_continuous_scale="Blues",
                             labels={"color":"Heures"})
            fig2.update_layout(height=max(300, len(pivot)*40),
                               paper_bgcolor="rgba(0,0,0,0)", font_color="#E8EDF5")
            st.plotly_chart(fig2, use_container_width=True)

        st.write("### Classe la plus active par mois")
        if not pivot.empty:
            top_by_month = pivot.idxmax(axis=0).to_frame(name="Classe top").T
            st.dataframe(top_by_month, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — ALERTES
# ══════════════════════════════════════════════════════════════════════════════
with tab_alertes:
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        st.subheader("Alertes intelligentes")
        tmp = filtered.copy()
        tmp["Alerte_retard_critique"] = tmp["Écart"] <= thresholds["ecart_critique"]
        tmp["Alerte_non_demarre"]     = tmp["Statut_auto"] == "Non démarré"

        _ret = np.where(tmp["Alerte_retard_critique"], "🔻 Retard critique", "")
        _nd  = np.where(tmp["Alerte_non_demarre"], "🛑 Non démarré", "")
        _sep = np.where((_ret != "") & (_nd != ""), " • ", "")
        tmp["Raison_alerte"] = pd.Series(
            np.char.add(np.char.add(_ret, _sep), _nd), index=tmp.index)
        tmp["En_alerte"] = tmp["Raison_alerte"].ne("")
        tmp["_prio"] = (tmp["Alerte_retard_critique"].astype(int)*2
                        + tmp["Alerte_non_demarre"].astype(int)*1)
        tmp = tmp.sort_values(["_prio","Écart"], ascending=[False,True])

        nb_alertes = int(tmp["En_alerte"].sum())
        nb_ret = int(tmp["Alerte_retard_critique"].sum())
        nb_nd  = int(tmp["Alerte_non_demarre"].sum())

        st.markdown(f"""
        <div style="display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin:10px 0;">
          <div class="kpi kpi-bad"><div class="kpi-title">Total alertes</div><div class="kpi-value">{nb_alertes}</div></div>
          <div class="kpi kpi-bad"><div class="kpi-title">Retards critiques</div><div class="kpi-value">{nb_ret}</div></div>
          <div class="kpi kpi-warn"><div class="kpi-title">Non démarrés</div><div class="kpi-value">{nb_nd}</div></div>
        </div>
        """, unsafe_allow_html=True)

        t1, t2 = st.tabs(["📌 Vue priorisée", "📊 Graphiques"])

        with t1:
            alerts = tmp.loc[tmp["En_alerte"],
                ["Classe","Matière","VHP","VHR","Écart","Taux","Statut_auto","Raison_alerte"]].copy()
            alerts["Taux (%)"] = (alerts["Taux"]*100).round(1)
            alerts["Statut"]   = alerts["Statut_auto"].apply(statut_badge)
            st.dataframe(alerts[["Classe","Matière","VHP","VHR","Écart","Taux (%)","Statut","Raison_alerte"]],
                use_container_width=True, height=520,
                column_config={
                    "Taux (%)": st.column_config.ProgressColumn("Taux (%)", min_value=0, max_value=100, format="%.1f%%"),
                    "Écart": st.column_config.NumberColumn("Écart (h)", format="%.0f"),
                })

        with t2:
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                al_types = pd.Series({"Retard critique": nb_ret, "Non démarré": nb_nd})
                fig = px.pie(values=al_types.values, names=al_types.index,
                             color_discrete_map={"Retard critique":"#FF3B3B","Non démarré":"#FF9500"})
                fig.update_layout(height=320, paper_bgcolor="rgba(0,0,0,0)", font_color="#E8EDF5")
                st.plotly_chart(fig, use_container_width=True)
            with col_g2:
                al_classe = tmp[tmp["En_alerte"]].groupby("Classe").size().sort_values(ascending=False)
                fig2 = px.bar(al_classe.reset_index(), x="Classe", y=0,
                              color_discrete_sequence=["#FF3B3B"])
                fig2.update_layout(height=320, yaxis_title="Nb alertes",
                                   paper_bgcolor="rgba(0,0,0,0)", font_color="#E8EDF5")
                st.plotly_chart(fig2, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — POINTAGE
# ══════════════════════════════════════════════════════════════════════════════
with tab_pointage:
    st.subheader("✏️ Système de Pointage des Séances")

    if df_all.empty:
        st.info("Aucun enseignement trouvé. Configure d'abord les classes dans Administration.")
    else:
        pt1, pt2 = st.tabs(["➕ Nouvelle saisie", "📋 Historique"])

        with pt1:
            st.markdown('<div class="pointage-card">', unsafe_allow_html=True)
            classes_list = sorted(df_all["Classe"].unique().tolist())
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                sel_classe = st.selectbox("Classe", classes_list, key="pt_cl")
            ens_df = df_all[df_all["Classe"]==sel_classe][["_rowid","Matière","Responsable"]].copy()
            ens_options = {
                f"{r['Matière']} — {r['Responsable'] or 'N/A'}": r["_rowid"]
                for _, r in ens_df.iterrows()
            }
            with col_s2:
                sel_label = st.selectbox("Matière — Responsable", list(ens_options.keys()), key="pt_mat")
            sel_id = ens_options.get(sel_label)

            col_d1, col_d2, col_d3 = st.columns(3)
            with col_d1: pt_date   = st.date_input("Date séance", value=date.today(), key="pt_dt")
            with col_d2: pt_heures = st.number_input("Heures", min_value=0.5, max_value=10.0, value=2.0, step=0.5, key="pt_h")
            with col_d3: pt_saisi  = st.text_input("Saisi par", placeholder="Votre nom", key="pt_s")
            pt_note = st.text_area("Remarque", key="pt_note", height=70)

            if st.button("✅ Enregistrer le pointage", type="primary"):
                try:
                    api.create_pointage(sel_id, str(pt_date), pt_heures, pt_note, pt_saisi)
                    st.success(f"✅ {pt_heures}h enregistrée(s) le {pt_date}")
                    st.cache_data.clear(); st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")
            st.markdown('</div>', unsafe_allow_html=True)

            # Récapitulatif de la matière sélectionnée
            if sel_id:
                row = df_all[df_all["_rowid"]==sel_id].iloc[0] if not df_all[df_all["_rowid"]==sel_id].empty else None
                if row is not None:
                    vhp = row["VHP"]; vhr = row["VHR"]; taux = row["Taux"]
                    statut = row["Statut_auto"]
                    badge_cls = {"Terminé":"ok","En cours":"warn","Non démarré":"nd"}.get(statut,"nd")
                    st.markdown(f"""
                    <div class="pointage-card" style="margin-top:12px;">
                      <strong>{row['Matière']}</strong> — <span class="pointage-badge-{badge_cls}">{statut}</span><br/>
                      <span style="color:var(--muted)">VHP: {vhp:.0f}h &nbsp;|&nbsp; VHR: {vhr:.1f}h &nbsp;|&nbsp; Taux: {taux*100:.1f}%</span>
                    </div>
                    """, unsafe_allow_html=True)

        with pt2:
            classes_list2 = sorted(df_all["Classe"].unique().tolist())
            col_h1, col_h2 = st.columns(2)
            with col_h1:
                hist_cl = st.selectbox("Classe", classes_list2, key="hcl")
            ens_df2 = df_all[df_all["Classe"]==hist_cl][["_rowid","Matière","Responsable"]].copy()
            ens_h2 = {f"{r['Matière']} — {r['Responsable'] or 'N/A'}": r["_rowid"]
                      for _, r in ens_df2.iterrows()}
            with col_h2:
                hist_label = st.selectbox("Matière", list(ens_h2.keys()), key="hmat")
            hist_id = ens_h2.get(hist_label)

            if hist_id:
                try:
                    hist = api.get_pointages(hist_id)
                    if hist:
                        df_hist = pd.DataFrame(hist)[["id","date","heures","note","saisi_par"]]
                        df_hist = df_hist.rename(columns={
                            "id":"ID","date":"Date","heures":"Heures","note":"Remarque","saisi_par":"Saisi par"})
                        st.dataframe(df_hist, use_container_width=True, hide_index=True)
                        total_h = sum(p["heures"] for p in hist)
                        vhp_v = float(df_all[df_all["_rowid"]==hist_id]["VHP"].values[0])
                        col_m1, col_m2 = st.columns(2)
                        col_m1.metric("Total réalisé", f"{total_h:.1f}h", delta=f"{total_h-vhp_v:+.1f}h vs VHP")
                        col_m2.metric("Nb séances", len(hist))

                        with st.expander("🗑️ Supprimer un pointage"):
                            del_id = st.selectbox("Pointage à supprimer", [p["id"] for p in hist],
                                format_func=lambda i: f"#{i} — {next(p['date'] for p in hist if p['id']==i)} — {next(p['heures'] for p in hist if p['id']==i)}h")
                            if st.button("Supprimer", type="secondary"):
                                api.delete_pointage(del_id)
                                st.cache_data.clear(); st.rerun()
                    else:
                        st.info("Aucun pointage enregistré pour cette matière.")
                except Exception as e:
                    st.error(f"Erreur : {e}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 8 — ADMINISTRATION
# ══════════════════════════════════════════════════════════════════════════════
with tab_admin:
    st.subheader(f"⚙️ Administration — {dept}")
    adm1, adm2 = st.tabs(["🏫 Classes", "📖 Enseignements"])

    with adm1:
        with st.form("form_cl"):
            n = st.text_input("Nouvelle classe", placeholder="ex: L3 IA")
            if st.form_submit_button("➕ Ajouter"):
                if n.strip():
                    try: api.create_classe(dept, n.strip()); st.success(f"'{n}' créée."); st.rerun()
                    except Exception as e: st.error(f"Erreur : {e}")
                else: st.warning("Saisis un nom.")

        cls_data = api.get_classes(dept)
        if cls_data:
            st.dataframe(pd.DataFrame(cls_data)[["id","nom"]].rename(columns={"id":"ID","nom":"Classe"}),
                         use_container_width=True, hide_index=True)
            with st.expander("🗑️ Supprimer une classe"):
                del_cl = st.selectbox("Classe", [c["id"] for c in cls_data],
                    format_func=lambda i: next(c["nom"] for c in cls_data if c["id"]==i))
                if st.button("Supprimer la classe", type="secondary"):
                    api.delete_classe(del_cl); st.cache_data.clear(); st.rerun()
        else:
            st.info("Aucune classe.")

    with adm2:
        cls_data2 = api.get_classes(dept)
        if not cls_data2:
            st.warning("Crée d'abord une classe.")
        else:
            cm = {c["nom"]: c["id"] for c in cls_data2}
            with st.form("form_ens"):
                c1, c2 = st.columns(2)
                with c1:
                    ec = st.selectbox("Classe", list(cm.keys()))
                    em = st.text_input("Matière *")
                    ev = st.number_input("VHP *", min_value=0.0, value=30.0, step=1.0)
                    es = st.selectbox("Semestre", ["","S1","S2","S3","S4"])
                with c2:
                    er = st.text_input("Responsable")
                    ee = st.text_input("Email")
                    ed = st.text_input("Début prévu", placeholder="ex: Oct 2024")
                    ef = st.text_input("Fin prévue",  placeholder="ex: Jan 2025")
                eo = st.text_area("Observations", height=60)
                if st.form_submit_button("➕ Ajouter l'enseignement"):
                    if em.strip() and ev > 0:
                        try:
                            api.create_enseignement({"classe_id":cm[ec],"matiere":em,"vhp":ev,
                                "responsable":er,"email":ee,"semestre":es,
                                "debut_prevu":ed,"fin_prevue":ef,"observations":eo})
                            st.success(f"'{em}' ajouté."); st.cache_data.clear(); st.rerun()
                        except Exception as e: st.error(f"Erreur : {e}")
                    else: st.warning("Matière et VHP obligatoires.")

            st.subheader("Liste des enseignements")
            sel_cl = st.selectbox("Filtrer", ["Toutes"]+list(cm.keys()), key="adm_cl")
            cid = cm.get(sel_cl) if sel_cl != "Toutes" else None
            try:
                ens_l = api.get_enseignements(dept, cid)
                if ens_l:
                    de = pd.DataFrame(ens_l)[["id","matiere","vhp","vhr","taux","responsable","semestre","statut"]]
                    de["taux"] = de["taux"].apply(lambda x: f"{x*100:.1f}%")
                    de = de.rename(columns={"id":"ID","matiere":"Matière","vhp":"VHP","vhr":"VHR",
                                            "taux":"Taux","responsable":"Responsable",
                                            "semestre":"Sem.","statut":"Statut"})
                    st.dataframe(de, use_container_width=True, hide_index=True)
                    with st.expander("🗑️ Supprimer un enseignement"):
                        del_e = st.selectbox("Enseignement", [e["id"] for e in ens_l],
                            format_func=lambda i: f"#{i} — {next(e['matiere'] for e in ens_l if e['id']==i)}")
                        if st.button("Supprimer l'enseignement", type="secondary"):
                            api.delete_enseignement(del_e); st.cache_data.clear(); st.rerun()
                else:
                    st.info("Aucun enseignement.")
            except Exception as e:
                st.error(f"Erreur : {e}")
