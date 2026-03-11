"""
Composants Charts — Dashboard ISI v3
Graphiques Plotly réutilisables avec thème dark.
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import List, Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#E8EDF5", family="Space Grotesk"),
    margin=dict(l=0, r=0, t=36, b=0),
)

COLORS = {
    "success": "#00C97A",
    "warning": "#FF9500",
    "danger":  "#FF3B3B",
    "blue":    "#1F6FEB",
    "blue2":   "#5AA2FF",
    "muted":   "#7A90B8",
}


def color_by_taux(taux: float) -> str:
    if taux >= 0.9:
        return COLORS["success"]
    if taux >= 0.6:
        return COLORS["warning"]
    return COLORS["danger"]


def bar_avancement_classes(df_classes: pd.DataFrame) -> go.Figure:
    """Barres horizontales triées par taux — classement des classes."""
    df = df_classes.copy().sort_values("taux_moy")
    df["taux_pct"] = (df["taux_moy"] * 100).round(1)
    df["color"] = df["taux_moy"].apply(color_by_taux)

    fig = go.Figure(go.Bar(
        x=df["taux_pct"],
        y=df["classe"],
        orientation="h",
        marker=dict(
            color=df["color"],
            line=dict(width=0),
        ),
        text=df["taux_pct"].apply(lambda x: f"{x:.1f}%"),
        textposition="outside",
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Taux: %{x:.1f}%<br>"
            "<extra></extra>"
        ),
    ))
    fig.add_vline(x=90, line_dash="dash", line_color=COLORS["success"],
                  annotation_text="Objectif 90%", annotation_position="top right")
    fig.add_vline(x=60, line_dash="dot", line_color=COLORS["warning"])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Avancement par classe", font=dict(size=14, color=COLORS["blue2"])),
        xaxis=dict(title="Taux de réalisation (%)", range=[0, 115],
                   gridcolor="rgba(255,255,255,.06)"),
        yaxis=dict(automargin=True, gridcolor="rgba(255,255,255,.06)"),
        height=max(300, len(df) * 38),
    )
    return fig


def bar_vhp_vhr_classes(df_classes: pd.DataFrame) -> go.Figure:
    """Barres groupées VHP vs VHR par classe."""
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="VHP (prévu)", x=df_classes["classe"], y=df_classes["vhp_total"],
        marker_color=COLORS["blue"], opacity=0.75,
    ))
    fig.add_trace(go.Bar(
        name="VHR (réalisé)", x=df_classes["classe"], y=df_classes["vhr_total"],
        marker_color=COLORS["success"],
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        barmode="group",
        title=dict(text="VHP vs VHR par classe", font=dict(size=14, color=COLORS["blue2"])),
        height=380,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis=dict(tickangle=-30, gridcolor="rgba(255,255,255,.06)"),
        yaxis=dict(title="Heures", gridcolor="rgba(255,255,255,.06)"),
    )
    return fig


def heatmap_classes_mois(pivot: pd.DataFrame, mois_cols: list) -> go.Figure:
    """Heatmap classe × mois — densité d'activité."""
    if pivot.empty:
        return go.Figure()
    fig = px.imshow(
        pivot.values,
        x=mois_cols,
        y=pivot.index.tolist(),
        color_continuous_scale=[
            [0, "#111B2E"], [0.2, "#1A3060"],
            [0.6, "#1F6FEB"], [1, "#00C97A"],
        ],
        labels={"color": "Heures"},
        aspect="auto",
    )
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Activité par classe et par mois", font=dict(size=14, color=COLORS["blue2"])),
        height=max(300, len(pivot) * 40),
        coloraxis_colorbar=dict(title="Heures", tickfont=dict(color="#E8EDF5")),
    )
    fig.update_traces(
        hovertemplate="<b>%{y}</b> — %{x}<br>%{z:.0f}h<extra></extra>"
    )
    return fig


def line_mensuel(totaux: dict, mois_order: list) -> go.Figure:
    """Courbe d'évolution mensuelle avec marqueurs."""
    x = mois_order
    y = [totaux.get(m, 0) for m in mois_order]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x, y=y, mode="lines+markers+text",
        line=dict(color=COLORS["blue"], width=3),
        marker=dict(size=9, color=COLORS["blue2"], line=dict(width=2, color=COLORS["blue"])),
        text=[f"{v:.0f}h" if v > 0 else "" for v in y],
        textposition="top center",
        textfont=dict(size=11, color="#E8EDF5"),
        fill="tozeroy",
        fillcolor="rgba(31,111,235,.12)",
        hovertemplate="<b>%{x}</b><br>%{y:.1f}h<extra></extra>",
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Heures réalisées par mois", font=dict(size=14, color=COLORS["blue2"])),
        height=340,
        xaxis=dict(gridcolor="rgba(255,255,255,.06)"),
        yaxis=dict(title="Heures", gridcolor="rgba(255,255,255,.06)", rangemode="tozero"),
    )
    return fig


def scatter_vhp_vhr(df: pd.DataFrame) -> go.Figure:
    """Scatter VHP vs VHR — détection visuelle des écarts."""
    df = df.copy()
    df["color_val"] = df.get("taux", df.get("Taux", 0.5))
    df["text"] = df.get("matiere", df.get("Matière", ""))

    max_val = max(df.get("vhp", df.get("VHP", pd.Series([100]))).max(),
                  df.get("vhr", df.get("VHR", pd.Series([100]))).max(), 1)

    vhp_col = "vhp" if "vhp" in df.columns else "VHP"
    vhr_col = "vhr" if "vhr" in df.columns else "VHR"
    mat_col = "matiere" if "matiere" in df.columns else "Matière"

    fig = go.Figure()
    # Ligne diagonale (objectif = VHR == VHP)
    fig.add_trace(go.Scatter(
        x=[0, max_val * 1.1], y=[0, max_val * 1.1],
        mode="lines", line=dict(color="rgba(0,201,122,.3)", dash="dash", width=1.5),
        name="Objectif (VHR=VHP)", hoverinfo="skip",
    ))
    fig.add_trace(go.Scatter(
        x=df[vhp_col], y=df[vhr_col],
        mode="markers",
        marker=dict(
            size=9,
            color=df["color_val"],
            colorscale=[[0, COLORS["danger"]], [0.6, COLORS["warning"]], [1, COLORS["success"]]],
            showscale=True,
            colorbar=dict(title="Taux", tickformat=".0%", tickfont=dict(color="#E8EDF5")),
            line=dict(width=1, color="rgba(255,255,255,.15)"),
        ),
        text=df[mat_col],
        hovertemplate="<b>%{text}</b><br>VHP: %{x:.0f}h<br>VHR: %{y:.0f}h<extra></extra>",
        name="Matières",
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="VHP vs VHR — Écarts visuels", font=dict(size=14, color=COLORS["blue2"])),
        height=420,
        xaxis=dict(title="VHP (prévu)", gridcolor="rgba(255,255,255,.06)"),
        yaxis=dict(title="VHR (réalisé)", gridcolor="rgba(255,255,255,.06)"),
    )
    return fig


def pie_statuts(counts: dict) -> go.Figure:
    """Camembert de répartition des statuts."""
    labels = list(counts.keys())
    values = list(counts.values())
    colors_map = {
        "Terminé": COLORS["success"],
        "En cours": COLORS["warning"],
        "Non démarré": "#4B5563",
    }
    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        marker=dict(colors=[colors_map.get(l, COLORS["blue"]) for l in labels],
                    line=dict(color="#080E1A", width=2)),
        hole=0.45,
        textinfo="label+percent",
        textfont=dict(size=12, color="#E8EDF5"),
        hovertemplate="<b>%{label}</b><br>%{value} matières (%{percent})<extra></extra>",
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Répartition des statuts", font=dict(size=14, color=COLORS["blue2"])),
        height=320,
        showlegend=True,
        legend=dict(font=dict(color="#E8EDF5")),
    )
    return fig


def gantt_planning(df_matieres: pd.DataFrame) -> Optional[go.Figure]:
    """
    Gantt interactif des matières avec barre de progression réelle.
    Requiert les colonnes: matiere, debut_prevu, fin_prevue, taux, statut.
    """
    req = ["matiere", "debut_prevu", "fin_prevue"]
    if not all(c in df_matieres.columns for c in req):
        return None

    df = df_matieres.copy()

    def _parse_date(s):
        if not s or str(s).strip() in ("", "nan", "None"):
            return None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%b %Y", "%B %Y"):
            try:
                return pd.to_datetime(s, format=fmt).date()
            except Exception:
                continue
        try:
            return pd.to_datetime(s).date()
        except Exception:
            return None

    df["start"] = df["debut_prevu"].apply(_parse_date)
    df["end"]   = df["fin_prevue"].apply(_parse_date)
    df = df.dropna(subset=["start", "end"])
    if df.empty:
        return None

    df["start"] = df["start"].apply(str)
    df["end"]   = df["end"].apply(str)

    color_map = {
        "Terminé": COLORS["success"],
        "En cours": COLORS["warning"],
        "Non démarré": "#4B5563",
    }

    fig = px.timeline(
        df,
        x_start="start",
        x_end="end",
        y="matiere",
        color="statut",
        color_discrete_map=color_map,
        hover_data={"matiere": True, "statut": True,
                    "start": False, "end": False},
        title="Planning prévisionnel",
    )

    # Ligne "aujourd'hui"
    fig.add_vline(
        x=date.today().isoformat(),
        line_dash="dash",
        line_color=COLORS["danger"],
        annotation_text=f"Aujourd'hui ({date.today().strftime('%d/%m')})",
        annotation_position="top right",
        annotation_font_color=COLORS["danger"],
    )

    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Planning prévisionnel", font=dict(size=14, color=COLORS["blue2"])),
        height=max(350, len(df) * 30),
        yaxis=dict(automargin=True),
        legend=dict(font=dict(color="#E8EDF5")),
    )
    return fig


def bar_alertes_par_classe(df_alertes: pd.DataFrame) -> go.Figure:
    """Barres empilées des alertes par classe."""
    if df_alertes.empty:
        return go.Figure()

    pivot = df_alertes.groupby(["classe", "niveau"]).size().unstack(fill_value=0).reset_index()
    fig = go.Figure()
    for niveau, color in [("critical", COLORS["danger"]), ("warning", COLORS["warning"]), ("info", COLORS["blue"])]:
        if niveau in pivot.columns:
            fig.add_trace(go.Bar(
                name=niveau.capitalize(),
                x=pivot["classe"],
                y=pivot[niveau],
                marker_color=color,
            ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        barmode="stack",
        title=dict(text="Alertes par classe", font=dict(size=14, color=COLORS["blue2"])),
        height=350,
        legend=dict(font=dict(color="#E8EDF5")),
        xaxis=dict(tickangle=-30),
    )
    return fig
