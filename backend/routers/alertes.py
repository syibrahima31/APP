"""
Router Alertes — Détection intelligente, priorités, marquage lu/non lu.
"""
from __future__ import annotations

from datetime import date
from typing import List, Optional

from fastapi import APIRouter, Depends, Query
from pydantic import BaseModel
from sqlalchemy.orm import Session

from db.database import get_db
from db.models import Alerte, Classe, Enseignement, Pointage

router = APIRouter()


# ── Schémas ───────────────────────────────────────────────────────────────

class AlerteOut(BaseModel):
    id: int
    enseignement_id: int
    matiere: str
    classe: str
    responsable: str
    email: str
    type_alerte: str
    message: str
    niveau: str
    est_lue: bool
    action_requise: str
    vhp: float
    vhr: float
    taux: float
    ecart: float

    model_config = {"from_attributes": True}


# ── Règles d'alertes ──────────────────────────────────────────────────────

_MONTH_TO_ACADEMIC = {10:1, 11:2, 12:3, 1:4, 2:5, 3:6, 4:7, 5:8, 6:9, 7:10, 8:11}

ALERT_RULES = [
    {
        "type": "non_demarre_tardif",
        "niveau": "critical",
        "action": "notify_chef",
        "condition": lambda vhr, vhp, taux, mois: vhr == 0 and mois >= 3,
        "message": lambda m, ec, vhp: f"'{m}' non démarré après 3 mois — action requise",
    },
    {
        "type": "ecart_critique",
        "niveau": "critical",
        "action": "notify_enseignant",
        "condition": lambda vhr, vhp, taux, mois: ec_check(vhr, vhp) < -15 and mois >= 4,
        "message": lambda m, ec, vhp: f"'{m}' : {abs(ec):.0f}h de retard critique",
    },
    {
        "type": "retard_modere",
        "niveau": "warning",
        "action": "flag_only",
        "condition": lambda vhr, vhp, taux, mois: taux < 0.5 and mois >= 5 and vhr > 0,
        "message": lambda m, ec, vhp: f"'{m}' : taux {taux_pct(taux):.0f}% insuffisant",
    },
    {
        "type": "depassement_vhp",
        "niveau": "info",
        "action": "flag_only",
        "condition": lambda vhr, vhp, taux, mois: taux > 1.1 and vhp > 0,
        "message": lambda m, ec, vhp: f"'{m}' dépasse le VHP ({taux_pct(taux):.0f}%)",
    },
]


def ec_check(vhr, vhp):
    return vhr - vhp


def taux_pct(t):
    return t * 100


def _run_rules(ens: Enseignement, vhr: float, mois_academique: int) -> list[dict]:
    """Applique toutes les règles sur un enseignement et retourne les alertes."""
    vhp = ens.vhp or 0.0
    taux = (vhr / vhp) if vhp > 0 else 0.0
    ecart = vhr - vhp
    found = []
    for rule in ALERT_RULES:
        try:
            # Règle ecart_critique a besoin d'ecart — on adapte
            args = (vhr, vhp, taux, mois_academique)
            triggered = rule["condition"](*args)
        except Exception:
            triggered = False

        if triggered:
            # Calcul message (closure avec valeurs locales)
            msg_fn = rule["message"]
            try:
                msg = msg_fn(ens.matiere, ecart, vhp)
                # Substitutions manuelles pour les closures qui référencent les vars
                msg = msg.replace(
                    f"{taux_pct(taux):.0f}%",
                    f"{taux_pct(taux):.0f}%"
                )
            except Exception:
                msg = f"Alerte {rule['type']} sur '{ens.matiere}'"

            found.append({
                "type": rule["type"],
                "niveau": rule["niveau"],
                "action": rule["action"],
                "message": msg,
            })
    return found


# ── Endpoints ─────────────────────────────────────────────────────────────

@router.get("/", response_model=List[AlerteOut])
def list_alertes(
    dept: str = Query("IAID"),
    non_lues_only: bool = Query(False),
    niveau: Optional[str] = Query(None),
    db: Session = Depends(get_db),
):
    """
    Retourne les alertes actives pour le département.
    Génère les alertes dynamiquement depuis les données courantes.
    """
    today = date.today()
    mois_acad = _MONTH_TO_ACADEMIC.get(today.month, 5)

    enseignements = (
        db.query(Enseignement)
        .join(Classe)
        .filter(Classe.departement_code == dept.upper())
        .all()
    )

    result: list[AlerteOut] = []

    for ens in enseignements:
        vhr = sum(p.heures for p in ens.pointages)
        vhp = ens.vhp or 0.0
        taux = (vhr / vhp) if vhp > 0 else 0.0
        ecart = vhr - vhp

        triggered_rules = _run_rules(ens, vhr, mois_acad)

        for r in triggered_rules:
            if niveau and r["niveau"] != niveau:
                continue

            result.append(AlerteOut(
                id=0,  # Alerte dynamique (non persistée)
                enseignement_id=ens.id,
                matiere=ens.matiere,
                classe=ens.classe.nom if ens.classe else "",
                responsable=ens.responsable or "",
                email=ens.email or "",
                type_alerte=r["type"],
                message=r["message"],
                niveau=r["niveau"],
                est_lue=False,
                action_requise=r["action"],
                vhp=round(vhp, 1),
                vhr=round(vhr, 1),
                taux=round(taux, 4),
                ecart=round(ecart, 1),
            ))

    # Tri : critical d'abord, puis warning, puis info ; puis par écart croissant
    ordre = {"critical": 0, "warning": 1, "info": 2}
    result.sort(key=lambda a: (ordre.get(a.niveau, 3), a.ecart))
    return result


@router.get("/stats")
def alertes_stats(dept: str = Query("IAID"), db: Session = Depends(get_db)):
    """Statistiques globales des alertes pour le département."""
    alertes = list_alertes(dept=dept, non_lues_only=False, niveau=None, db=db)
    by_niveau = {"critical": 0, "warning": 0, "info": 0}
    by_type: dict = {}
    for a in alertes:
        by_niveau[a.niveau] = by_niveau.get(a.niveau, 0) + 1
        by_type[a.type_alerte] = by_type.get(a.type_alerte, 0) + 1
    return {
        "total": len(alertes),
        "by_niveau": by_niveau,
        "by_type": by_type,
        "enseignants_concernes": len({a.responsable for a in alertes if a.responsable}),
        "classes_concernees": len({a.classe for a in alertes if a.classe}),
    }
