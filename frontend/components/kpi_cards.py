"""
Composants KPI Cards — Dashboard ISI v3
"""
from __future__ import annotations

import streamlit as st

COLORS = {
    "bg_card":    "#1A2740",
    "success":    "#00C97A",
    "warning":    "#FF9500",
    "danger":     "#FF3B3B",
    "info":       "#1F6FEB",
    "muted":      "#7A90B8",
    "text":       "#E8EDF5",
}


def kpi_card(title: str, value: str, color: str = "info",
             subtitle: str = "", delta: str = "", icon: str = "") -> str:
    """Retourne le HTML d'une carte KPI."""
    c = COLORS.get(color, COLORS["info"])
    delta_html = f'<div style="color:{c};font-size:13px;font-weight:700;margin-top:4px">{delta}</div>' if delta else ""
    sub_html = f'<div style="color:{COLORS["muted"]};font-size:12px;margin-top:2px">{subtitle}</div>' if subtitle else ""
    return f"""
<div style="
  background:{COLORS['bg_card']};
  border-left:4px solid {c};
  border-radius:12px;
  padding:16px 18px;
  box-shadow:0 4px 20px rgba(0,0,0,.30);
">
  <div style="color:{COLORS['muted']};font-size:11px;text-transform:uppercase;letter-spacing:.8px;font-weight:700">
    {icon} {title}
  </div>
  <div style="color:{c};font-size:30px;font-weight:900;letter-spacing:-.5px;margin-top:6px;line-height:1">
    {value}
  </div>
  {delta_html}
  {sub_html}
</div>"""


def kpi_band(metrics: dict):
    """
    Bandeau 4 KPIs horizontal en haut de page.
    metrics = dict avec: taux_global, nb_critiques, nb_non_demarrees, retard_cumule
    """
    taux = metrics.get("taux_global", 0)
    nb_crit = metrics.get("nb_critiques", 0)
    nb_nd = metrics.get("nb_non_demarrees", 0)
    retard = metrics.get("retard_cumule", 0)
    total = metrics.get("total_matieres", 0)
    nb_term = metrics.get("nb_terminees", 0)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        color = "success" if taux >= 0.9 else ("warning" if taux >= 0.6 else "danger")
        st.markdown(kpi_card("Taux Global", f"{taux*100:.1f}%", color,
                             subtitle=f"{nb_term}/{total} matières terminées", icon="📈"),
                    unsafe_allow_html=True)
    with c2:
        color = "danger" if nb_crit > 0 else "success"
        st.markdown(kpi_card("Alertes Critiques", str(nb_crit), color,
                             subtitle="matières à risque élevé", icon="🚨"),
                    unsafe_allow_html=True)
    with c3:
        color = "warning" if nb_nd > 0 else "success"
        st.markdown(kpi_card("Non Démarrées", str(nb_nd), color,
                             subtitle="matières sans séance", icon="⏸️"),
                    unsafe_allow_html=True)
    with c4:
        color = "danger" if retard < -20 else ("warning" if retard < 0 else "success")
        st.markdown(kpi_card("Retard Cumulé", f"{retard:.0f}h", color,
                             subtitle="heures de retard global", icon="⏱️"),
                    unsafe_allow_html=True)


def kpi_band_6(metrics: dict):
    """Bandeau 6 KPIs — version étendue."""
    taux = metrics.get("taux_global", 0)
    nb_crit = metrics.get("nb_critiques", 0)
    nb_nd = metrics.get("nb_non_demarrees", 0)
    retard = metrics.get("retard_cumule", 0)
    vhp = metrics.get("total_vhp", 0)
    vhr = metrics.get("total_vhr", 0)
    nb_enc = metrics.get("nb_en_cours", 0)
    total = metrics.get("total_matieres", 0)

    c1, c2, c3 = st.columns(3)
    c4, c5, c6 = st.columns(3)

    with c1:
        color = "success" if taux >= 0.9 else ("warning" if taux >= 0.6 else "danger")
        st.markdown(kpi_card("Taux Global", f"{taux*100:.1f}%", color, icon="📊"), unsafe_allow_html=True)
    with c2:
        st.markdown(kpi_card("Total matières", str(total), "info", icon="📚"), unsafe_allow_html=True)
    with c3:
        color = "danger" if nb_crit > 0 else "success"
        st.markdown(kpi_card("Alertes Critiques", str(nb_crit), color, icon="🚨"), unsafe_allow_html=True)
    with c4:
        st.markdown(kpi_card("VHP total", f"{vhp:.0f}h", "info", icon="🎯"), unsafe_allow_html=True)
    with c5:
        st.markdown(kpi_card("VHR total", f"{vhr:.0f}h", "success" if vhr >= vhp*0.8 else "warning", icon="✅"), unsafe_allow_html=True)
    with c6:
        color = "danger" if retard < -20 else ("warning" if retard < 0 else "success")
        st.markdown(kpi_card("Retard cumulé", f"{retard:.0f}h", color, icon="⏱️"), unsafe_allow_html=True)
