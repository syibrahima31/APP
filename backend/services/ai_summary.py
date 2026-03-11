"""
Service IA — Résumé exécutif structuré des enseignements via GPT-4o-mini.
"""
from __future__ import annotations

import json
from typing import Optional

from backend.core.config import settings


SYSTEM_PROMPT = """Tu es un assistant expert en pilotage pédagogique universitaire (ISI - Institut Supérieur Informatique, Sénégal).
Analyse les données d'enseignement fournies et génère un rapport structuré en Markdown.

Ton rapport doit contenir EXACTEMENT ces sections :
## Résumé exécutif
3 bullet points maximum, orientés décision pour le chef de département.

## Points d'attention prioritaires
Les 3 matières/enseignants les plus critiques avec raison précise.

## Tendances positives
Ce qui avance bien et mérite d'être valorisé.

## Recommandations
2-3 actions concrètes et actionnables dans les 30 prochains jours.

## Analyse des observations
Synthèse des remarques pédagogiques des enseignants (si disponibles).

Ton style : direct, factuel, sans jargon, chiffres précis, orienté action.
Langue : Français professionnel.
"""


def build_context(summary_data: dict) -> str:
    """Construit le contexte structuré pour le prompt."""
    dept = summary_data.get("dept", "")
    mois = summary_data.get("mois_courant", "")
    total = summary_data.get("total_matieres", 0)
    vhp = summary_data.get("total_vhp", 0)
    vhr = summary_data.get("total_vhr", 0)
    taux = summary_data.get("taux_global", 0)
    nb_term = summary_data.get("nb_terminees", 0)
    nb_enc = summary_data.get("nb_en_cours", 0)
    nb_nd = summary_data.get("nb_non_demarrees", 0)
    nb_crit = summary_data.get("nb_critiques", 0)
    retard = summary_data.get("retard_cumule", 0)

    context = f"""
DÉPARTEMENT : {dept} | MOIS : {mois}

STATISTIQUES GLOBALES :
- Taux d'avancement global : {taux*100:.1f}%
- Matières : {total} total ({nb_term} terminées, {nb_enc} en cours, {nb_nd} non démarrées)
- Volume horaire : VHP={vhp:.0f}h | VHR={vhr:.0f}h | Retard cumulé={retard:.0f}h
- Alertes critiques actives : {nb_crit}

MATIÈRES EN RETARD CRITIQUE (top 10) :
"""
    critiques = summary_data.get("matieres_critiques", [])
    for m in critiques[:10]:
        context += f"- [{m.get('classe','')}] {m.get('matiere','')} — "
        context += f"VHP={m.get('vhp',0):.0f}h, VHR={m.get('vhr',0):.0f}h, "
        context += f"Taux={m.get('taux',0)*100:.0f}%, Écart={m.get('ecart',0):.0f}h"
        resp = m.get('responsable','')
        if resp:
            context += f" (Responsable: {resp})"
        context += "\n"

    context += "\nANALYSE PAR CLASSE :\n"
    for cls in summary_data.get("par_classe", []):
        context += (
            f"- {cls.get('classe','')} : {cls.get('nb_matieres',0)} matières, "
            f"taux={cls.get('taux_moy',0)*100:.0f}%, retard={cls.get('retard_h',0):.0f}h\n"
        )

    obs_list = summary_data.get("observations", [])
    if obs_list:
        context += "\nOBSERVATIONS PÉDAGOGIQUES DES ENSEIGNANTS :\n"
        for obs in obs_list[:20]:
            context += f"- [{obs.get('classe','')}] {obs.get('matiere','')}: {obs.get('texte','')[:200]}\n"

    return context.strip()


async def generate_summary(summary_data: dict) -> str:
    """
    Génère un résumé IA structuré des données du département.
    Retourne un Markdown formaté.
    """
    if not settings.OPENAI_API_KEY:
        return _fallback_summary(summary_data)

    try:
        from openai import AsyncOpenAI
        client = AsyncOpenAI(api_key=settings.OPENAI_API_KEY)

        context = build_context(summary_data)

        response = await client.chat.completions.create(
            model=settings.OPENAI_MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": context},
            ],
            max_tokens=1800,
            temperature=0.3,
        )
        return response.choices[0].message.content

    except Exception as e:
        return _fallback_summary(summary_data, error=str(e))


def _fallback_summary(data: dict, error: str = "") -> str:
    """Résumé statique si l'IA n'est pas disponible."""
    dept = data.get("dept", "")
    taux = data.get("taux_global", 0)
    nb_crit = data.get("nb_critiques", 0)
    nb_nd = data.get("nb_non_demarrees", 0)
    retard = data.get("retard_cumule", 0)

    note_ia = f"\n\n> ⚠️ IA non disponible ({error})" if error else ""

    return f"""## Résumé exécutif — {dept}

- Taux d'avancement global : **{taux*100:.1f}%**
- Alertes critiques actives : **{nb_crit}** matières
- Matières non démarrées : **{nb_nd}** | Retard cumulé : **{retard:.0f}h**

## Points d'attention
{"- " + chr(10).join(
    f"[{m.get('classe','')}] {m.get('matiere','')} — Taux {m.get('taux',0)*100:.0f}%"
    for m in data.get("matieres_critiques", [])[:3]
) if data.get("matieres_critiques") else "- Aucune alerte critique détectée."}

## Recommandations
- Contacter les enseignants des matières non démarrées
- Organiser un point de situation mensuel
- Vérifier les matières avec écart > 15h{note_ia}
"""
