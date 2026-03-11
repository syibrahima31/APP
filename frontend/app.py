"""
Frontend Streamlit — Dashboard ISI v3
Lancement : streamlit run frontend/app.py
"""
from __future__ import annotations

import os, sys, datetime as dt
from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import frontend.api_client as api
from frontend.components.kpi_cards import kpi_band, kpi_card
from frontend.components.charts import (
    bar_avancement_classes, bar_vhp_vhr_classes, heatmap_classes_mois,
    line_mensuel, scatter_vhp_vhr, pie_statuts, gantt_planning,
    bar_alertes_par_classe, COLORS, PLOTLY_LAYOUT,
)

st.set_page_config(
    page_title="ISI — Dashboard Suivi",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEPT_OPTIONS = ["IAID", "KM", "DRS", "GI"]
DEPT_LABELS  = {
    "IAID": "IA & Ingénierie des Données",
    "KM":   "Directions des Études",
    "DRS":  "Réseaux et Systèmes",
    "GI":   "Génie Informatique",
}
MOIS_COLS = ["Oct", "Nov", "Déc", "Jan", "Fév", "Mars", "Avril", "Mai", "Juin", "Juil", "Août"]
MOIS_API  = ["oct","nov","dec","jan","fev","mars","avril","mai","juin","juil","aout"]

# ─────────────────────────────────────────────────────────────────────────────
# CSS THÈME DARK EXÉCUTIF
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700;800&display=swap');
:root{
  --bg:#080E1A;--bg2:#0D1526;--card:#111B2E;--card2:#162035;
  --text:#E8EDF5;--muted:#7A90B8;--line:rgba(255,255,255,0.07);
  --blue:#1F6FEB;--blue2:#2A80FF;--blue3:#5AA2FF;
  --ok:#00C97A;--warn:#FF9500;--bad:#FF3B3B;
}
html,body,.stApp{
  background:radial-gradient(ellipse at 10% 0%,#0D1E3A 0%,var(--bg) 60%)!important;
  font-family:'Space Grotesk',sans-serif!important; color:var(--text)!important;
}
h1,h2,h3,h4,h5{color:var(--blue3)!important;font-weight:850!important;}
.block-container{padding-top:.5rem!important;padding-bottom:4rem!important;}
header[data-testid="stHeader"]{background:transparent!important;border-bottom:0!important;}
::-webkit-scrollbar{width:6px;height:6px;}
::-webkit-scrollbar-thumb{background:rgba(90,162,255,0.25);border-radius:4px;}
section[data-testid="stSidebar"]{background:var(--bg2)!important;border-right:1px solid var(--line);}
.sidebar-card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:12px;margin-bottom:10px;}
.iaid-header{background:linear-gradient(135deg,#0A1E44 0%,#0F2860 40%,#1A3A80 100%);
  border:1px solid rgba(90,162,255,.20);padding:20px 26px;border-radius:20px;margin-bottom:16px;}
.iaid-header *{color:#FFF!important;}
.iaid-htitle{font-size:20px;font-weight:950;}
.iaid-hsub{font-size:13px;opacity:.80;margin-top:4px;}
.iaid-badges{margin-top:12px;display:flex;gap:8px;flex-wrap:wrap;}
.iaid-badge{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.18);
  padding:5px 12px;border-radius:999px;font-size:12px;font-weight:850;}
button[data-baseweb="tab"]{background:var(--card)!important;color:var(--muted)!important;
  border-radius:999px!important;padding:10px 16px!important;font-weight:800!important;border:1px solid var(--line)!important;}
button[data-baseweb="tab"][aria-selected="true"]{background:rgba(31,111,235,.18)!important;
  color:var(--blue3)!important;border:1px solid rgba(90,162,255,.40)!important;}
.stButton>button{background:linear-gradient(135deg,var(--blue),var(--blue2))!important;
  color:#FFF!important;border:none!important;border-radius:14px!important;font-weight:800!important;}
.pointage-card{background:linear-gradient(135deg,var(--card) 0%,var(--card2) 100%);
  border:1px solid rgba(90,162,255,.20);border-radius:20px;padding:20px;margin:8px 0;}
.alert-critical{background:rgba(255,59,59,.08);border-left:4px solid #FF3B3B;
  border-radius:0 12px 12px 0;padding:12px 16px;margin:6px 0;}
.alert-warning{background:rgba(255,149,0,.08);border-left:4px solid #FF9500;
  border-radius:0 12px 12px 0;padding:12px 16px;margin:6px 0;}
.login-card{background:var(--card);border:1px solid var(--line);border-radius:20px;
  padding:40px;max-width:440px;margin:60px auto;}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# AUTH
# ─────────────────────────────────────────────────────────────────────────────
def _login_wall():
    st.markdown('<div class="login-card"><div style="text-align:center;margin-bottom:24px">'
                '<div style="font-size:40px">📊</div>'
                '<div style="font-size:22px;font-weight:900;color:#5AA2FF">Dashboard ISI</div>'
                '<div style="color:#7A90B8;font-size:13px">Suivi des Enseignements</div>'
                '</div></div>', unsafe_allow_html=True)
    with st.form("login_form"):
        username = st.text_input("Identifiant", placeholder="admin")
        password = st.text_input("Mot de passe", type="password")
        if st.form_submit_button("Se connecter", type="primary", use_container_width=True):
            try:
                r = api.login(username, password)
                st.session_state.update({
                    "access_token": r["access_token"], "refresh_token": r["refresh_token"],
                    "role": r["role"], "dept_login": r["dept"], "username": r["username"],
                })
                st.rerun()
            except Exception:
                st.error("Identifiants incorrects.")
    st.caption("Compte par défaut : admin / changeme")


try:
    api.health()
except Exception:
    st.error("❌ API non joignable — `uvicorn backend.main:app --reload`")
    st.stop()

if "access_token" not in st.session_state:
    _login_wall()
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# FONCTIONS CACHE (définies avant la sidebar)
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load_classes(d: str) -> list:
    try: return api.get_classes(d)
    except Exception: return []


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    logo_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "logo_iaid.jpg")
    if os.path.exists(logo_path):
        st.image(logo_path, width=140)
    else:
        st.markdown("<h2 style='text-align:center'>📊 ISI</h2>", unsafe_allow_html=True)

    user = st.session_state.get("username", "")
    role = st.session_state.get("role", "viewer")
    rl   = {"admin":"🔴 Admin","chef_dept":"🟡 Chef","viewer":"🟢 Lecteur"}.get(role, role)
    st.markdown(f'<div style="color:#5AA2FF;font-size:12px;text-align:center;margin-bottom:8px">👤 {user} — {rl}</div>', unsafe_allow_html=True)
    if st.button("🚪 Déconnexion", use_container_width=True):
        for k in ["access_token","refresh_token","role","dept_login","username"]:
            st.session_state.pop(k, None)
        st.rerun()
    st.divider()

    st.markdown('<div class="sidebar-card">', unsafe_allow_html=True)
    dept_default = 0
    dept_login = st.session_state.get("dept_login", "IAID")
    if role != "admin" and dept_login in DEPT_OPTIONS:
        dept_default = DEPT_OPTIONS.index(dept_login)
    dept = st.selectbox("🏛️ Département", DEPT_OPTIONS, index=dept_default)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="sidebar-card">', unsafe_allow_html=True)
    st.markdown("**🔎 Filtres**")
    classes_data = load_classes(dept)
    filtre_classe = st.selectbox("Classe", ["Toutes"] + [c["nom"] for c in classes_data])
    filtre_statut = st.multiselect("Statut",
        ["Non démarré","En cours","Terminé"], default=["Non démarré","En cours","Terminé"])
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="sidebar-card">', unsafe_allow_html=True)
    st.markdown("**⚙️ Seuils**")
    taux_vert   = st.slider("Seuil Vert (%)",   50, 100, 90, 5) / 100
    taux_orange = st.slider("Seuil Orange (%)", 10, 95,  60, 5) / 100
    ecart_crit  = st.slider("Écart critique (h)", -40, 0, -6, 1)
    st.markdown('</div>', unsafe_allow_html=True)

    if st.button("🔄 Rafraîchir", use_container_width=True):
        st.cache_data.clear(); st.rerun()
    st.caption(f"v3.0 | {os.environ.get('API_BASE_URL','localhost:8000')}")

thresholds = {"taux_vert": taux_vert, "taux_orange": taux_orange, "ecart_critique": ecart_crit}


# ─────────────────────────────────────────────────────────────────────────────
# DATA
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load_dashboard(d: str) -> dict:
    try:
        return api.get_dashboard(d)
    except Exception:
        data = api.get_resume(d)
        return {"matieres": data, "total_matieres": len(data), "taux_global": 0,
                "total_vhp": 0, "total_vhr": 0, "nb_terminees": 0,
                "nb_en_cours": 0, "nb_non_demarrees": 0, "nb_critiques": 0, "retard_cumule": 0}


@st.cache_data(ttl=300, show_spinner=False)
def load_alertes(d: str) -> list:
    try: return api.get_alertes(d)
    except Exception: return []


def _to_df(matieres: list) -> pd.DataFrame:
    if not matieres: return pd.DataFrame()
    df = pd.DataFrame(matieres)
    rn = {
        "matiere":"Matière","classe":"Classe","responsable":"Responsable",
        "semestre":"Semestre","email":"Email","vhp":"VHP","vhr":"VHR","taux":"Taux",
        "ecart":"Écart","statut":"Statut_auto","enseignement_id":"_rowid",
        "debut_prevu":"Début prévu","fin_prevue":"Fin prévue","observations":"Observations",
        "oct":"Oct","nov":"Nov","dec":"Déc","jan":"Jan","fev":"Fév","mars":"Mars",
        "avril":"Avril","mai":"Mai","juin":"Juin","juil":"Juil","aout":"Août",
    }
    df = df.rename(columns={k:v for k,v in rn.items() if k in df.columns})
    for c in MOIS_COLS:
        if c not in df.columns: df[c] = 0.0
    for c in ["Semestre","Email","Observations","Début prévu","Fin prévue","Statut"]:
        if c not in df.columns: df[c] = ""
    if "Statut" not in df.columns or df["Statut"].eq("").all():
        df["Statut"] = df.get("Statut_auto", "")
    for c in ["Écart","Taux","VHR","VHP"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df


def sbadge(s): return {"Terminé":"🟢 Terminé","En cours":"🟡 En cours","Non démarré":"⚫ Non démarré"}.get(str(s),str(s))


with st.spinner("Chargement…"):
    dashboard_data = load_dashboard(dept)
    alertes_data   = load_alertes(dept)

df_all = _to_df(dashboard_data.get("matieres", []))
nb_alertes_crit = sum(1 for a in alertes_data if a.get("niveau") == "critical")

if nb_alertes_crit > 0:
    st.markdown(
        f'<div class="alert-critical">🚨 <strong>{nb_alertes_crit} alerte(s) critique(s)</strong> '
        f'pour {dept} — voir l\'onglet <strong>🚨 Alertes</strong></div>',
        unsafe_allow_html=True)

st.markdown(f"""
<div class="iaid-header">
  <div class="iaid-htitle">📊 Dashboard Suivi des Enseignements</div>
  <div class="iaid-hsub">{DEPT_LABELS.get(dept,dept)} — Pilotage mensuel</div>
  <div class="iaid-badges">
    <span class="iaid-badge">🏛️ {dept}</span>
    <span class="iaid-badge">📚 {len(classes_data)} classe(s)</span>
    <span class="iaid-badge">📖 {len(df_all)} matière(s)</span>
    <span class="iaid-badge">🗓️ {dt.date.today().strftime('%d/%m/%Y')}</span>
    {"<span class='iaid-badge' style='background:rgba(255,59,59,.25)'>🚨 "+str(nb_alertes_crit)+"</span>" if nb_alertes_crit else ""}
  </div>
</div>""", unsafe_allow_html=True)

if df_all.empty:
    st.info("Aucune donnée. Importez un fichier Excel dans **Administration → Import Excel**.")
    filtered = pd.DataFrame()
else:
    filtered = df_all.copy()
    if filtre_classe != "Toutes":
        filtered = filtered[filtered["Classe"] == filtre_classe]
    if filtre_statut:
        filtered = filtered[filtered["Statut_auto"].isin(filtre_statut)]


# ─────────────────────────────────────────────────────────────────────────────
# ONGLETS
# ─────────────────────────────────────────────────────────────────────────────
tabs = st.tabs(["🌐 Vue globale","🏫 Par classe","📖 Par matière","👤 Par enseignant",
                "📅 Mensuel","🚨 Alertes","📆 Planning","🤖 IA","📋 Programmation","✏️ Pointage","⚙️ Admin"])
(tab_ov, tab_cl, tab_mat, tab_ens,
 tab_men, tab_al, tab_gtt, tab_ai,
 tab_prog, tab_pt, tab_adm) = tabs


# ═══ TAB 1 — VUE GLOBALE ════════════════════════════════════════════════════
with tab_ov:
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        kpi_band({
            "taux_global": dashboard_data.get("taux_global", float(filtered["Taux"].mean())),
            "nb_critiques": nb_alertes_crit,
            "nb_non_demarrees": int((filtered["Statut_auto"]=="Non démarré").sum()),
            "retard_cumule": float(filtered.loc[filtered["Écart"]<0,"Écart"].sum()),
            "total_matieres": len(filtered),
            "nb_terminees": int((filtered["Statut_auto"]=="Terminé").sum()),
            "nb_en_cours": int((filtered["Statut_auto"]=="En cours").sum()),
            "total_vhp": float(filtered["VHP"].sum()),
            "total_vhr": float(filtered["VHR"].sum()),
        })
        st.divider()
        c1, c2 = st.columns([3,2])
        with c1:
            g = filtered.groupby("Classe")["Taux"].mean().reset_index()
            g.columns = ["classe","taux_moy"]
            st.plotly_chart(bar_avancement_classes(g), use_container_width=True)
        with c2:
            st.plotly_chart(pie_statuts(filtered["Statut_auto"].value_counts().to_dict()), use_container_width=True)
        g2 = filtered.groupby("Classe").agg(vhp_total=("VHP","sum"),vhr_total=("VHR","sum")).reset_index()
        g2.columns = ["classe","vhp_total","vhr_total"]
        st.plotly_chart(bar_vhp_vhr_classes(g2), use_container_width=True)
        st.write("### Top 20 retards")
        tr = filtered[["Classe","Matière","VHP","VHR","Écart","Taux","Statut_auto"]].copy().sort_values("Écart").head(20)
        tr["Taux (%)"] = (tr["Taux"]*100).round(1)
        tr["Statut"] = tr["Statut_auto"].apply(sbadge)
        st.dataframe(tr[["Classe","Matière","VHP","VHR","Écart","Taux (%)","Statut"]], use_container_width=True, height=400,
            column_config={"Taux (%)": st.column_config.ProgressColumn("Taux (%)", min_value=0, max_value=100, format="%.1f%%"),
                           "Écart": st.column_config.NumberColumn("Écart (h)", format="%.0f"),
                           "VHP": st.column_config.NumberColumn("VHP", format="%.0f"),
                           "VHR": st.column_config.NumberColumn("VHR", format="%.0f")})


# ═══ TAB 2 — PAR CLASSE ═════════════════════════════════════════════════════
with tab_cl:
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        cf = sorted(filtered["Classe"].unique().tolist())
        ca, cb = st.columns([2,1])
        with cb:
            c1 = st.selectbox("Classe A", cf, index=0, key="cls1")
            c2 = st.selectbox("Classe B", cf, index=min(1,len(cf)-1), key="cls2")
        with ca:
            st.write("### Synthèse par classe")
            sy = filtered.groupby("Classe").agg(
                Matieres=("Matière","count"), Taux_moy=("Taux","mean"),
                VHP_total=("VHP","sum"), VHR_total=("VHR","sum"),
                Retard_h=("Écart", lambda s: float(s[s<0].sum())),
                Terminees=("Statut_auto", lambda s: int((s=="Terminé").sum())),
                Non_demarre=("Statut_auto", lambda s: int((s=="Non démarré").sum())),
            ).reset_index()
            sy["Taux (%)"] = (sy["Taux_moy"]*100).round(1)
            st.dataframe(sy[["Classe","Matieres","Taux (%)","VHP_total","VHR_total","Retard_h","Terminees","Non_demarre"]],
                use_container_width=True,
                column_config={"Taux (%)": st.column_config.ProgressColumn("Taux (%)", min_value=0, max_value=100, format="%.1f%%"),
                               "Retard_h": st.column_config.NumberColumn("Retard (h)", format="%.0f")})
        st.divider()
        A = filtered[filtered["Classe"]==c1]; B = filtered[filtered["Classe"]==c2]
        def kpis(o): return {"Matières":len(o),"Taux moyen":f"{float(o['Taux'].mean()*100):.1f}%" if len(o) else "0%",
                              "Retard (h)":float(o.loc[o["Écart"]<0,"Écart"].sum()),"Non démarré":int((o["Statut_auto"]=="Non démarré").sum())}
        kA,kB = kpis(A),kpis(B)
        st.write(f"### Comparaison — {c1} vs {c2}")
        st.dataframe(pd.DataFrame({"Indicateur":list(kA.keys()),c1:list(kA.values()),c2:list(kB.values())}),
                     use_container_width=True, hide_index=True)
        if not A.empty:
            st.write(f"### Scatter VHP/VHR — {c1}")
            st.plotly_chart(scatter_vhp_vhr(A), use_container_width=True)


# ═══ TAB 3 — PAR MATIÈRE ════════════════════════════════════════════════════
with tab_mat:
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        st.subheader("Analyse par matière")
        mat = filtered.groupby("Matière").agg(
            Classes=("Classe","nunique"), VHP=("VHP","sum"), VHR=("VHR","sum"),
            Taux=("Taux","mean"), Retard=("Écart", lambda s: float(s[s<0].sum())),
            Non_demarre=("Statut_auto", lambda s: int((s=="Non démarré").sum())),
        ).reset_index()
        mat["Taux (%)"] = (mat["Taux"]*100).round(1)
        st.dataframe(mat.sort_values(["Taux (%)","Retard"], ascending=[True,True]), use_container_width=True, height=420)
        al = mat[(mat["Taux"] < thresholds["taux_orange"]) | (mat["Retard"] <= thresholds["ecart_critique"])].copy()
        st.write("### Matières en alerte")
        if al.empty: st.success("✅ Aucune alerte.")
        else: st.dataframe(al.sort_values("Taux (%)").head(30), use_container_width=True)
        fig = px.bar(mat.nsmallest(20,"Taux (%)"), x="Matière", y="Taux (%)",
                     color="Taux (%)", color_continuous_scale=["#FF3B3B","#FF9500","#00C97A"])
        fig.update_layout(**PLOTLY_LAYOUT, height=400, xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)


# ═══ TAB 4 — PAR ENSEIGNANT ═════════════════════════════════════════════════
with tab_ens:
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        st.subheader("Suivi par enseignant")
        tmp = filtered.copy()
        tmp["Responsable"] = tmp["Responsable"].astype(str).replace({"nan":"","None":""}).fillna("").str.strip()
        tmp["Responsable"] = tmp["Responsable"].replace({"":"⚠️ Non affecté"})
        sr = tmp.groupby("Responsable").agg(
            Matieres=("Matière","count"), Classes=("Classe","nunique"),
            VHP_total=("VHP","sum"), VHR_total=("VHR","sum"), Taux_moy=("Taux","mean"),
            Retard_h=("Écart", lambda s: float(s[s<0].sum())),
            Non_demarre=("Statut_auto", lambda s: int((s=="Non démarré").sum())),
        ).reset_index()
        sr["Taux (%)"] = (sr["Taux_moy"]*100).round(1)
        SEUIL = 300
        sr["Charge"] = sr["VHP_total"].apply(lambda v: "🔴 Surcharge" if v>SEUIL*1.2 else ("🟡 Élevée" if v>SEUIL else "🟢 Normale"))
        st.dataframe(sr[["Responsable","Matieres","Classes","Taux (%)","VHP_total","VHR_total","Retard_h","Charge","Non_demarre"]].sort_values("Taux (%)"),
            use_container_width=True,
            column_config={"Taux (%)": st.column_config.ProgressColumn("Taux (%)", min_value=0, max_value=100, format="%.1f%%"),
                           "Retard_h": st.column_config.NumberColumn("Retard (h)", format="%.0f")})
        st.divider()
        st.write("### Simulation de charge")
        profs = sorted([p for p in tmp["Responsable"].unique() if "Non affecté" not in str(p)])
        sc1,sc2,sc3 = st.columns(3)
        with sc1: sp = st.selectbox("Enseignant", profs, key="smp")
        with sc2: sn = st.number_input("Nouvelles matières", 0, 20, 2, 1, key="smn")
        with sc3: sv = st.slider("VHP moyen/matière (h)", 10, 80, 30, 5, key="smv")
        pd_row = sr[sr["Responsable"]==sp]
        ca_v = float(pd_row["VHP_total"].values[0]) if not pd_row.empty else 0
        cs_v = ca_v + sn * sv
        sm1,sm2,sm3 = st.columns(3)
        sm1.metric("Charge actuelle", f"{ca_v:.0f}h")
        sm2.metric("Charge simulée", f"{cs_v:.0f}h", delta=f"+{sn*sv:.0f}h")
        sm3.metric("Statut", "⚠️ Surcharge" if cs_v>SEUIL else "✅ OK", delta=f"seuil={SEUIL}h")
        fs = go.Figure()
        fs.add_trace(go.Bar(name="Actuelle", x=["Charge"], y=[ca_v], marker_color="#1F6FEB"))
        fs.add_trace(go.Bar(name="Simulée", x=["Charge"], y=[cs_v], marker_color="#FF3B3B" if cs_v>SEUIL else "#00C97A"))
        fs.add_hline(y=SEUIL, line_dash="dash", line_color="#FF9500", annotation_text=f"Seuil ({SEUIL}h)", annotation_font_color="#FF9500")
        fs.update_layout(**PLOTLY_LAYOUT, height=280, barmode="group")
        st.plotly_chart(fs, use_container_width=True)


# ═══ TAB 5 — MENSUEL ════════════════════════════════════════════════════════
with tab_men:
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        st.subheader("Analyse mensuelle")
        md = [m for m in MOIS_COLS if m in filtered.columns]
        my = filtered[md].sum()
        totaux = {MOIS_API[MOIS_COLS.index(m)]: float(my[m]) for m in md}
        st.plotly_chart(line_mensuel(totaux, MOIS_API), use_container_width=True)
        pivot = filtered.groupby("Classe")[md].sum()
        if pivot.shape[0]*pivot.shape[1] <= 300:
            st.plotly_chart(heatmap_classes_mois(pivot, md), use_container_width=True)
        else:
            st.dataframe(pivot, use_container_width=True)
        if not pivot.empty:
            st.write("### Classe la plus active par mois")
            st.dataframe(pivot.idxmax(axis=0).to_frame(name="Classe top").T, use_container_width=True)


# ═══ TAB 6 — ALERTES ════════════════════════════════════════════════════════
with tab_al:
    st.subheader("🚨 Alertes intelligentes")
    if not alertes_data:
        st.success("✅ Aucune alerte.")
    else:
        dfa = pd.DataFrame(alertes_data)
        nc,nw,ni = int((dfa["niveau"]=="critical").sum()), int((dfa["niveau"]=="warning").sum()), int((dfa["niveau"]=="info").sum())
        aa1,aa2,aa3,aa4 = st.columns(4)
        with aa1: st.markdown(kpi_card("Total", str(len(dfa)), "danger", icon="🚨"), unsafe_allow_html=True)
        with aa2: st.markdown(kpi_card("Critiques", str(nc), "danger", icon="🔴"), unsafe_allow_html=True)
        with aa3: st.markdown(kpi_card("Avertissements", str(nw), "warning", icon="🟡"), unsafe_allow_html=True)
        with aa4: st.markdown(kpi_card("Informations", str(ni), "info", icon="🔵"), unsafe_allow_html=True)
        st.markdown("---")
        al1,al2 = st.tabs(["📋 Liste","📊 Graphiques"])
        with al1:
            for _,row in dfa.iterrows():
                niv = row.get("niveau","info")
                css = "alert-critical" if niv=="critical" else "alert-warning"
                icon = "🔴" if niv=="critical" else ("🟡" if niv=="warning" else "🔵")
                st.markdown(f'<div class="{css}"><strong>{icon} {str(row.get("type_alerte","")).replace("_"," ").title()}</strong>'
                            f' — {row.get("classe","")} | {str(row.get("matiere",""))[:60]}<br>'
                            f'<span style="color:#7A90B8;font-size:12px">{row.get("message","")} | '
                            f'VHP={row.get("vhp",0):.0f}h VHR={row.get("vhr",0):.0f}h Écart={row.get("ecart",0):.0f}h | '
                            f'👤 {row.get("responsable","")}</span></div>', unsafe_allow_html=True)
        with al2:
            ag1,ag2 = st.columns(2)
            with ag1:
                cv = dfa["niveau"].value_counts().to_dict()
                fp = px.pie(values=list(cv.values()), names=list(cv.keys()),
                            color_discrete_map={"critical":"#FF3B3B","warning":"#FF9500","info":"#1F6FEB"})
                fp.update_layout(**PLOTLY_LAYOUT, height=300)
                st.plotly_chart(fp, use_container_width=True)
            with ag2:
                if "classe" in dfa.columns:
                    st.plotly_chart(bar_alertes_par_classe(dfa), use_container_width=True)


# ═══ TAB 7 — GANTT ══════════════════════════════════════════════════════════
with tab_gtt:
    st.subheader("📆 Planning prévisionnel — Gantt")
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        gc = st.selectbox("Filtrer", ["Toutes"]+sorted(filtered["Classe"].unique().tolist()), key="gc")
        df_g = filtered if gc=="Toutes" else filtered[filtered["Classe"]==gc]
        df_gr = df_g.rename(columns={"Matière":"matiere","Statut_auto":"statut","Taux":"taux",
                                      "Début prévu":"debut_prevu","Fin prévue":"fin_prevue"})
        fg = gantt_planning(df_gr)
        if fg:
            st.plotly_chart(fg, use_container_width=True)
        else:
            st.info("Renseignez **Début prévu** et **Fin prévue** dans votre Excel pour afficher le Gantt.")
        plan_c = [c for c in ["Classe","Matière","Responsable","Début prévu","Fin prévue","VHP","VHR","Statut_auto"] if c in filtered.columns]
        st.dataframe(filtered[plan_c].head(50), use_container_width=True)


# ═══ TAB 9 — PROGRAMMATION DES SÉANCES ══════════════════════════════════════
with tab_prog:
    st.subheader("📋 Programmation des séances")
    cls_list = load_classes(dept)
    cls_map = {c["nom"]: c["id"] for c in cls_list}

    pg1, pg2 = st.tabs(["➕ Programmer une séance", "📅 Calendrier & Pointage"])

    with pg1:
        st.markdown('<div class="pointage-card">', unsafe_allow_html=True)
        pg_c1, pg_c2 = st.columns(2)
        with pg_c1:
            pg_cls = st.selectbox("Classe *", list(cls_map.keys()), key="pg_cls")
        pg_cid = cls_map.get(pg_cls)
        if pg_cid:
            ens_list = []
            try: ens_list = api.get_enseignements(dept, pg_cid)
            except Exception: pass
            ens_map = {f"{e['matiere']} — {e.get('responsable','') or 'N/A'}": e["id"] for e in ens_list}
            with pg_c2:
                pg_ens = st.selectbox("Matière / Enseignant *", list(ens_map.keys()) if ens_map else ["Aucun enseignement"], key="pg_ens")
            pg_eid = ens_map.get(pg_ens)

            pg_d1, pg_d2, pg_d3 = st.columns(3)
            with pg_d1: pg_date = st.date_input("Date *", value=date.today(), key="pg_date")
            with pg_d2: pg_hdeb = st.text_input("Heure début", value="08:00", placeholder="HH:MM", key="pg_hdeb")
            with pg_d3: pg_hfin = st.text_input("Heure fin", value="10:00", placeholder="HH:MM", key="pg_hfin")
            pg_s1, pg_s2 = st.columns(2)
            with pg_s1: pg_salle = st.text_input("Salle", placeholder="Amphi A / Salle 12", key="pg_salle")
            with pg_s2: pg_saisi = st.text_input("Programmé par", value=st.session_state.get("username",""), key="pg_saisi")
            pg_note = st.text_area("Note", height=60, key="pg_note")

            if st.button("📅 Programmer la séance", type="primary", key="pg_btn"):
                if pg_eid and pg_date:
                    try:
                        # calcul heures auto
                        try:
                            h1 = int(pg_hdeb.split(":")[0]) + int(pg_hdeb.split(":")[1])/60
                            h2 = int(pg_hfin.split(":")[0]) + int(pg_hfin.split(":")[1])/60
                            nb_h = round(h2 - h1, 2) if h2 > h1 else 2.0
                        except Exception: nb_h = 2.0
                        api.create_seance({
                            "enseignement_id": pg_eid,
                            "date_prevue": str(pg_date),
                            "heure_debut": pg_hdeb,
                            "heure_fin": pg_hfin,
                            "nb_heures": nb_h,
                            "salle": pg_salle,
                            "note": pg_note,
                            "created_by": pg_saisi,
                        })
                        st.success(f"✅ Séance programmée le {pg_date} de {pg_hdeb} à {pg_hfin}")
                        st.cache_data.clear(); st.rerun()
                    except Exception as e: st.error(f"Erreur : {e}")
                else:
                    st.warning("Sélectionnez une matière et une date.")
        st.markdown('</div>', unsafe_allow_html=True)

    with pg2:
        # Filtres semaine
        pv1, pv2, pv3 = st.columns([2,2,1])
        with pv1:
            today = date.today()
            lundi_def = today - dt.timedelta(days=today.weekday())
            pg_lundi = st.date_input("Semaine du lundi", value=lundi_def, key="pg_lundi")
        with pv2:
            pg_statut_f = st.selectbox("Statut", ["Tous","planifie","effectue","annule"], key="pg_statut_f")
        with pv3:
            if st.button("🔄", key="pg_refresh"): st.cache_data.clear(); st.rerun()

        try:
            seances = api.get_seances_semaine(dept, str(pg_lundi))
            if pg_statut_f != "Tous":
                seances = [s for s in seances if s["statut"] == pg_statut_f]
        except Exception: seances = []

        if not seances:
            st.info("Aucune séance programmée pour cette semaine.")
        else:
            JOURS = {0:"Lun",1:"Mar",2:"Mer",3:"Jeu",4:"Ven",5:"Sam",6:"Dim"}
            STATUT_CSS = {"planifie":"🔵 Planifié","effectue":"🟢 Effectué","annule":"🔴 Annulé"}
            STATUT_COLOR = {"planifie":"rgba(31,111,235,.12)","effectue":"rgba(0,201,122,.12)","annule":"rgba(255,59,59,.10)"}

            for s in seances:
                d = dt.date.fromisoformat(str(s["date_prevue"]))
                jour = JOURS.get(d.weekday(), "")
                color = STATUT_COLOR.get(s["statut"], "var(--card)")
                label = STATUT_CSS.get(s["statut"], s["statut"])
                with st.container():
                    sc1, sc2 = st.columns([4, 1])
                    with sc1:
                        st.markdown(
                            f'<div style="background:{color};border:1px solid var(--line);border-radius:14px;'
                            f'padding:12px 16px;margin:4px 0">'
                            f'<strong>{jour} {d.strftime("%d/%m")}</strong> '
                            f'<span style="color:#7A90B8">{s["heure_debut"]}→{s["heure_fin"]}</span> '
                            f'<span style="margin-left:8px">{label}</span><br/>'
                            f'<strong>{s.get("matiere","")}</strong> — {s.get("classe_nom","")} '
                            f'| 👤 {s.get("responsable","N/A")} '
                            f'{"| 🏫 " + s["salle"] if s.get("salle") else ""} '
                            f'| ⏱️ {s["nb_heures"]}h'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                    with sc2:
                        if s["statut"] == "planifie":
                            if st.button("✅ Pointer", key=f"pt_{s['id']}"):
                                try:
                                    api.pointer_seance(s["id"], saisi_par=st.session_state.get("username",""))
                                    st.success("Pointé !")
                                    st.cache_data.clear(); st.rerun()
                                except Exception as e: st.error(str(e))
                            if st.button("❌ Annuler", key=f"ann_{s['id']}"):
                                try:
                                    api.annuler_seance(s["id"])
                                    st.cache_data.clear(); st.rerun()
                                except Exception as e: st.error(str(e))
                        elif s["statut"] == "effectue":
                            st.markdown('<div style="color:#00C97A;font-size:11px;padding:4px">✅ Effectué</div>', unsafe_allow_html=True)


# ═══ TAB 8 — IA INSIGHTS ════════════════════════════════════════════════════
with tab_ai:
    st.subheader("🤖 Analyse IA — Résumé exécutif")
    if filtered.empty:
        st.info("Aucune donnée.")
    else:
        ia1,ia2 = st.columns([2,1])
        with ia2:
            mc = st.selectbox("Mois courant", MOIS_COLS, index=min(dt.date.today().month-1,len(MOIS_COLS)-1), key="amc")
            mr = st.slider("Top retards", 5, 20, 10)
        with ia1:
            if st.button("🤖 Générer le résumé IA", type="primary"):
                with st.spinner("Analyse…"):
                    crit = filtered[filtered["Taux"]<thresholds["taux_orange"]].sort_values("Écart").head(mr)
                    pcls = (filtered.groupby("Classe").agg(nb_matieres=("Matière","count"),taux_moy=("Taux","mean"),
                                                           retard_h=("Écart",lambda s:float(s[s<0].sum()))).reset_index())
                    obs_l = [{"classe":r["Classe"],"matiere":r["Matière"],"texte":r["Observations"]}
                             for _,r in filtered[filtered["Observations"].str.len()>5].head(30).iterrows()]
                    sd = {
                        "dept":dept,"mois_courant":mc,"total_matieres":len(filtered),
                        "total_vhp":float(filtered["VHP"].sum()),"total_vhr":float(filtered["VHR"].sum()),
                        "taux_global":dashboard_data.get("taux_global",0),
                        "nb_terminees":int((filtered["Statut_auto"]=="Terminé").sum()),
                        "nb_en_cours":int((filtered["Statut_auto"]=="En cours").sum()),
                        "nb_non_demarrees":int((filtered["Statut_auto"]=="Non démarré").sum()),
                        "nb_critiques":nb_alertes_crit,
                        "retard_cumule":float(filtered.loc[filtered["Écart"]<0,"Écart"].sum()),
                        "matieres_critiques":crit[["Classe","Matière","VHP","VHR","Taux","Écart","Responsable"]].rename(
                            columns={"Classe":"classe","Matière":"matiere","Responsable":"responsable","Taux":"taux","Écart":"ecart"}).to_dict("records"),
                        "par_classe":pcls.rename(columns={"Classe":"classe"}).to_dict("records"),
                        "observations":obs_l,
                    }
                    try:
                        import asyncio
                        from backend.services.ai_summary import generate_summary
                        st.session_state["ai_sum"] = asyncio.run(generate_summary(sd))
                    except Exception as e:
                        from backend.services.ai_summary import _fallback_summary
                        st.session_state["ai_sum"] = _fallback_summary(sd, str(e))
            if "ai_sum" in st.session_state:
                st.markdown("---")
                st.markdown(st.session_state["ai_sum"])
                st.download_button("📥 Exporter", st.session_state["ai_sum"],
                                   file_name=f"resume_ia_{dept}_{date.today()}.md", mime="text/markdown")
            else:
                st.info("Cliquez sur **Générer** pour lancer l'analyse IA.")


# ═══ TAB 9 — POINTAGE ═══════════════════════════════════════════════════════
with tab_pt:
    st.subheader("✏️ Pointage des séances")
    if df_all.empty:
        st.info("Aucun enseignement. Configurez d'abord les classes.")
    else:
        pp1,pp2 = st.tabs(["➕ Nouvelle saisie","📋 Historique"])
        with pp1:
            st.markdown('<div class="pointage-card">', unsafe_allow_html=True)
            cls_l = sorted(df_all["Classe"].unique().tolist())
            ps1,ps2 = st.columns(2)
            with ps1: scl = st.selectbox("Classe", cls_l, key="ptcl")
            ed = df_all[df_all["Classe"]==scl][["_rowid","Matière","Responsable"]].copy()
            eo = {f"{r['Matière']} — {r['Responsable'] or 'N/A'}": r["_rowid"] for _,r in ed.iterrows()}
            with ps2: sl = st.selectbox("Matière — Responsable", list(eo.keys()), key="ptmat")
            sid = eo.get(sl)
            pd1,pd2,pd3 = st.columns(3)
            with pd1: ptd = st.date_input("Date", value=date.today(), key="ptdt")
            with pd2: pth = st.number_input("Heures", min_value=0.5, max_value=10.0, value=2.0, step=0.5, key="pth")
            with pd3: pts = st.text_input("Saisi par", value=st.session_state.get("username",""), key="pts")
            ptn = st.text_area("Remarque", key="ptn", height=70)
            if st.button("✅ Enregistrer", type="primary"):
                try:
                    api.create_pointage(sid, str(ptd), pth, ptn, pts)
                    st.success(f"✅ {pth}h enregistrée(s) le {ptd}")
                    st.cache_data.clear(); st.rerun()
                except Exception as e: st.error(f"Erreur : {e}")
            st.markdown('</div>', unsafe_allow_html=True)
            if sid and not df_all[df_all["_rowid"]==sid].empty:
                rw = df_all[df_all["_rowid"]==sid].iloc[0]
                bc = {"Terminé":"#00C97A","En cours":"#FF9500","Non démarré":"#4B5563"}.get(rw["Statut_auto"],"#4B5563")
                st.markdown(f'<div class="pointage-card" style="margin-top:12px"><strong>{rw["Matière"]}</strong>'
                            f'<span style="background:{bc};color:#fff;padding:3px 10px;border-radius:20px;font-size:12px;margin-left:8px">{rw["Statut_auto"]}</span><br/>'
                            f'<span style="color:#7A90B8">VHP:{rw["VHP"]:.0f}h | VHR:{rw["VHR"]:.1f}h | Taux:{rw["Taux"]*100:.1f}%</span></div>',
                            unsafe_allow_html=True)
        with pp2:
            ph1,ph2 = st.columns(2)
            with ph1: hcl = st.selectbox("Classe", cls_l, key="hcl")
            eh = df_all[df_all["Classe"]==hcl][["_rowid","Matière","Responsable"]].copy()
            eh2 = {f"{r['Matière']} — {r['Responsable'] or 'N/A'}": r["_rowid"] for _,r in eh.iterrows()}
            with ph2: hlb = st.selectbox("Matière", list(eh2.keys()), key="hmat")
            hid = eh2.get(hlb)
            if hid:
                try:
                    hist = api.get_pointages(hid)
                    if hist:
                        dh = pd.DataFrame(hist)[["id","date","heures","note","saisi_par"]].rename(
                            columns={"id":"ID","date":"Date","heures":"Heures","note":"Remarque","saisi_par":"Saisi par"})
                        st.dataframe(dh, use_container_width=True, hide_index=True)
                        th = sum(p["heures"] for p in hist)
                        vv = float(df_all[df_all["_rowid"]==hid]["VHP"].values[0])
                        hm1,hm2 = st.columns(2)
                        hm1.metric("Total réalisé", f"{th:.1f}h", delta=f"{th-vv:+.1f}h vs VHP")
                        hm2.metric("Nb séances", len(hist))
                        with st.expander("🗑️ Supprimer"):
                            di = st.selectbox("Pointage", [p["id"] for p in hist],
                                format_func=lambda i: f"#{i} — {next(p['date'] for p in hist if p['id']==i)} — {next(p['heures'] for p in hist if p['id']==i)}h")
                            if st.button("Supprimer", type="secondary"):
                                api.delete_pointage(di); st.cache_data.clear(); st.rerun()
                    else: st.info("Aucun pointage.")
                except Exception as e: st.error(f"Erreur : {e}")


# ═══ TAB 10 — ADMINISTRATION ════════════════════════════════════════════════
with tab_adm:
    st.subheader(f"⚙️ Administration — {dept}")
    adm1,adm2,adm3,adm4,adm5 = st.tabs(["🏫 Classes","📖 Enseignements","📥 Import","📊 Export","👤 Utilisateurs"])

    with adm1:
        with st.form("fcl"):
            nn = st.text_input("Nouvelle classe", placeholder="ex: L3 IA")
            if st.form_submit_button("➕ Ajouter"):
                if nn.strip():
                    try: api.create_classe(dept, nn.strip()); st.success(f"'{nn}' créée."); st.rerun()
                    except Exception as e: st.error(f"Erreur : {e}")
        cd = api.get_classes(dept)
        if cd:
            st.dataframe(pd.DataFrame(cd)[["id","nom"]].rename(columns={"id":"ID","nom":"Classe"}),
                         use_container_width=True, hide_index=True)
            with st.expander("🗑️ Supprimer"):
                dc = st.selectbox("Classe", [c["id"] for c in cd],
                    format_func=lambda i: next(c["nom"] for c in cd if c["id"]==i))
                if st.button("Supprimer la classe", type="secondary"):
                    api.delete_classe(dc); st.cache_data.clear(); st.rerun()

    with adm2:
        cd2 = api.get_classes(dept)
        if not cd2:
            st.warning("Crée d'abord une classe.")
        else:
            cm = {c["nom"]:c["id"] for c in cd2}
            with st.form("fens"):
                fe1,fe2 = st.columns(2)
                with fe1:
                    fec=st.selectbox("Classe",list(cm.keys())); fem=st.text_input("Matière *")
                    fev=st.number_input("VHP *",min_value=0.0,value=30.0,step=1.0)
                    fes=st.selectbox("Semestre",["","S1","S2","S3","S4","S5","S6","S7","S8"])
                with fe2:
                    fer=st.text_input("Responsable"); fee=st.text_input("Email")
                    fed=st.text_input("Début prévu",placeholder="2024-10-01")
                    fef=st.text_input("Fin prévue",placeholder="2025-01-31")
                feo=st.text_area("Observations",height=60)
                if st.form_submit_button("➕ Ajouter"):
                    if fem.strip() and fev>0:
                        try:
                            api.create_enseignement({"classe_id":cm[fec],"matiere":fem,"vhp":fev,
                                "responsable":fer,"email":fee,"semestre":fes,
                                "debut_prevu":fed,"fin_prevue":fef,"observations":feo})
                            st.success(f"'{fem}' ajouté."); st.cache_data.clear(); st.rerun()
                        except Exception as e: st.error(f"Erreur : {e}")
            scle = st.selectbox("Filtrer", ["Toutes"]+list(cm.keys()), key="adcl")
            cid = cm.get(scle) if scle!="Toutes" else None
            try:
                el = api.get_enseignements(dept, cid)
                if el:
                    de = pd.DataFrame(el)[["id","matiere","vhp","vhr","taux","responsable","semestre"]]
                    de["taux"] = de["taux"].apply(lambda x: f"{x*100:.1f}%")
                    de = de.rename(columns={"id":"ID","matiere":"Matière","vhp":"VHP","vhr":"VHR",
                                            "taux":"Taux","responsable":"Responsable","semestre":"Sem."})
                    st.dataframe(de, use_container_width=True, hide_index=True)
                    with st.expander("🗑️ Supprimer"):
                        dei = st.selectbox("Enseignement", [e["id"] for e in el],
                            format_func=lambda i: f"#{i} — {next(e['matiere'] for e in el if e['id']==i)}")
                        if st.button("Supprimer l'enseignement", type="secondary"):
                            api.delete_enseignement(dei); st.cache_data.clear(); st.rerun()
            except Exception as e: st.error(f"Erreur : {e}")

    with adm3:
        st.write("### Import Excel idempotent")
        ii1,ii2 = st.tabs(["📁 Fichier local","🌐 URL"])
        with ii1:
            impf = st.file_uploader("Fichier Excel", type=["xlsx","xls"])
            impa = st.text_input("Année", value="2024-2025", key="impa")
            impc = st.checkbox("Effacer avant import", key="impc")
            if impf and st.button("📥 Importer", type="primary"):
                with st.spinner("Import…"):
                    try:
                        r = api.import_excel_bytes(impf.read(), impf.name, dept, impa, impc)
                        st.success(f"✅ {r['created']} créés, {r['updated']} MàJ, {r['unchanged']} inchangés")
                        if r.get("errors"): st.warning(f"⚠️ {len(r['errors'])} erreur(s)")
                        st.cache_data.clear(); st.rerun()
                    except Exception as e: st.error(f"Erreur: {e}")
        with ii2:
            impu = st.text_input("URL", key="impu")
            impa2 = st.text_input("Année", value="2024-2025", key="impa2")
            impc2 = st.checkbox("Effacer avant", key="impc2")
            if impu and st.button("📥 Importer URL", type="primary"):
                with st.spinner("…"):
                    try:
                        r = api.import_excel_url(impu, dept, impa2, impc2)
                        st.success(f"✅ {r['created']} créés, {r['updated']} MàJ, {r['unchanged']} inchangés")
                        st.cache_data.clear(); st.rerun()
                    except Exception as e: st.error(f"Erreur: {e}")

    with adm4:
        st.write("### Export")
        mex = dashboard_data.get("matieres",[])
        if not mex:
            st.info("Aucune donnée.")
        else:
            ex1,ex2 = st.columns(2)
            with ex1:
                if st.button("📊 Excel", use_container_width=True):
                    with st.spinner("…"):
                        try:
                            from backend.services.export import export_dashboard_excel
                            xb = export_dashboard_excel(mex, dept)
                            st.download_button("⬇️ Excel", xb, file_name=f"isi_{dept}_{date.today()}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                        except Exception as e: st.error(f"Erreur: {e}")
            with ex2:
                if st.button("📄 PDF", use_container_width=True):
                    with st.spinner("…"):
                        try:
                            from backend.services.export import export_dashboard_pdf
                            from config.departments import get_department_config
                            pb = export_dashboard_pdf(mex, dept, get_department_config(dept))
                            st.download_button("⬇️ PDF", pb, file_name=f"isi_{dept}_{date.today()}.pdf",
                                mime="application/pdf", use_container_width=True)
                        except Exception as e: st.error(f"Erreur: {e}")

    with adm5:
        if role != "admin":
            st.warning("Réservé aux administrateurs.")
        else:
            st.write("### Utilisateurs")
            try:
                users = api.list_users()
                if users:
                    st.dataframe(pd.DataFrame(users)[["id","username","email","role","departement_code","est_actif"]],
                                 use_container_width=True, hide_index=True)
            except Exception as e: st.error(f"Erreur : {e}")
            with st.form("fuser"):
                uu=st.text_input("Username"); ue2=st.text_input("Email")
                up=st.text_input("Mot de passe",type="password")
                ur=st.selectbox("Rôle",["viewer","chef_dept","admin"]); ud=st.selectbox("Dépt",DEPT_OPTIONS)
                if st.form_submit_button("➕ Créer"):
                    if uu and ue2 and up:
                        try: api.create_user(uu,ue2,up,ur,ud); st.success("Utilisateur créé."); st.rerun()
                        except Exception as e: st.error(f"Erreur : {e}")
            st.divider()
            st.write("### Mon mot de passe")
            with st.form("fpwd"):
                op=st.text_input("Actuel",type="password"); np1=st.text_input("Nouveau",type="password"); np2=st.text_input("Confirmer",type="password")
                if st.form_submit_button("🔒 Changer"):
                    if np1!=np2: st.error("Ne correspondent pas.")
                    elif len(np1)<8: st.error("Min 8 caractères.")
                    else:
                        try: api.change_password(op,np1); st.success("Modifié.")
                        except Exception as e: st.error(f"Erreur : {e}")
