"""
Router Analytics — Métriques avancées, prédictions, résumés.
"""
from __future__ import annotations

from collections import defaultdict
from datetime import date
from typing import List, Optional

from fastapi import APIRouter, Depends, Query
from pydantic import BaseModel
from sqlalchemy.orm import Session

from db.database import get_db
from db.models import Classe, Enseignement, Pointage

router = APIRouter()

MOIS_ACADEMIQUE = {
    10: "oct", 11: "nov", 12: "dec",
    1: "jan", 2: "fev", 3: "mars",
    4: "avril", 5: "mai", 6: "juin",
    7: "juil", 8: "aout",
}


# ── Schémas ───────────────────────────────────────────────────────────────

class MatiereMetrics(BaseModel):
    enseignement_id: int
    matiere: str
    classe: str
    responsable: str
    email: str
    semestre: str
    vhp: float
    vhr: float
    taux: float
    ecart: float
    statut: str
    nb_seances: int
    # Heures par mois
    oct: float = 0.0; nov: float = 0.0; dec: float = 0.0
    jan: float = 0.0; fev: float = 0.0; mars: float = 0.0
    avril: float = 0.0; mai: float = 0.0; juin: float = 0.0
    juil: float = 0.0; aout: float = 0.0
    # Prédiction
    proba_retard: float = 0.0
    vhr_projete: float = 0.0
    prediction_label: str = ""


class DashboardSummary(BaseModel):
    dept: str
    total_matieres: int
    total_vhp: float
    total_vhr: float
    taux_global: float
    nb_terminees: int
    nb_en_cours: int
    nb_non_demarrees: int
    nb_critiques: int
    retard_cumule: float
    matieres: List[MatiereMetrics]


class ClasseSummary(BaseModel):
    classe: str
    nb_matieres: int
    taux_moy: float
    vhp_total: float
    vhr_total: float
    retard_h: float
    nb_terminees: int
    nb_en_cours: int
    nb_non_demarrees: int


class EnseignantSummary(BaseModel):
    responsable: str
    email: str
    nb_matieres: int
    nb_classes: int
    vhp_total: float
    vhr_total: float
    taux_moy: float
    retard_h: float
    nb_non_demarrees: int
    # Simulation de charge
    charge_label: str = ""


# ── Calculs ───────────────────────────────────────────────────────────────

def _compute_matiere(ens: Enseignement) -> MatiereMetrics:
    heures_mois: dict = defaultdict(float)
    total = 0.0
    nb = 0
    for p in ens.pointages:
        try:
            mois_num = int(p.date[5:7])
        except Exception:
            continue
        col = MOIS_ACADEMIQUE.get(mois_num, "")
        if col:
            heures_mois[col] += p.heures
        total += p.heures
        nb += 1

    vhp = ens.vhp or 0.0
    taux = (total / vhp) if vhp > 0 else 0.0
    ecart = total - vhp

    if total <= 0:
        statut = "Non démarré"
    elif total < vhp:
        statut = "En cours"
    else:
        statut = "Terminé"

    # Prédiction linéaire (mois courant dans l'année académique Oct→Août)
    proba, vhr_proj, pred_label = _predict_retard(total, vhp, taux)

    return MatiereMetrics(
        enseignement_id=ens.id,
        matiere=ens.matiere,
        classe=ens.classe.nom if ens.classe else "",
        responsable=ens.responsable or "",
        email=ens.email or "",
        semestre=ens.semestre or "",
        vhp=vhp,
        vhr=round(total, 2),
        taux=round(taux, 4),
        ecart=round(ecart, 2),
        statut=statut,
        nb_seances=nb,
        proba_retard=proba,
        vhr_projete=round(vhr_proj, 1),
        prediction_label=pred_label,
        **{col: round(h, 2) for col, h in heures_mois.items()},
    )


def _predict_retard(vhr: float, vhp: float, taux: float) -> tuple[float, float, str]:
    """
    Projection linéaire basée sur la vélocité actuelle.
    Retourne (proba_retard 0-1, vhr_projeté, label).
    """
    if vhp <= 0:
        return 0.0, 0.0, ""

    today = date.today()
    # Mois académique courant (Oct=1 ... Août=11)
    month_to_academic = {10:1, 11:2, 12:3, 1:4, 2:5, 3:6, 4:7, 5:8, 6:9, 7:10, 8:11}
    mois_ecoule = month_to_academic.get(today.month, 5)
    mois_restants = 11 - mois_ecoule

    if mois_ecoule <= 0:
        return 0.5, vhr, "Semestre non commencé"

    velocite = vhr / mois_ecoule
    vhr_proj = vhr + velocite * mois_restants

    if taux >= 1.0:
        return 0.0, vhr_proj, "✅ Terminé"

    deficit = vhp - vhr_proj
    proba = max(0.0, min(1.0, deficit / vhp))

    if proba > 0.5:
        label = f"⚠️ Risque élevé — projection {vhr_proj:.0f}h/{vhp:.0f}h"
    elif proba > 0.2:
        label = f"🟡 Risque modéré — vérifier le rythme"
    else:
        label = f"🟢 Trajectoire correcte"

    return round(proba, 3), vhr_proj, label


# ── Endpoints ─────────────────────────────────────────────────────────────

@router.get("/dashboard", response_model=DashboardSummary)
def get_dashboard(dept: str = Query("IAID"), db: Session = Depends(get_db)):
    """Résumé complet du département avec métriques calculées."""
    enseignements = (
        db.query(Enseignement)
        .join(Classe)
        .filter(Classe.departement_code == dept.upper())
        .all()
    )

    matieres = [_compute_matiere(e) for e in enseignements]

    total_vhp = sum(m.vhp for m in matieres)
    total_vhr = sum(m.vhr for m in matieres)
    taux_global = (total_vhr / total_vhp) if total_vhp > 0 else 0.0

    return DashboardSummary(
        dept=dept.upper(),
        total_matieres=len(matieres),
        total_vhp=round(total_vhp, 1),
        total_vhr=round(total_vhr, 1),
        taux_global=round(taux_global, 4),
        nb_terminees=sum(1 for m in matieres if m.statut == "Terminé"),
        nb_en_cours=sum(1 for m in matieres if m.statut == "En cours"),
        nb_non_demarrees=sum(1 for m in matieres if m.statut == "Non démarré"),
        nb_critiques=sum(1 for m in matieres if m.proba_retard > 0.5),
        retard_cumule=round(sum(m.ecart for m in matieres if m.ecart < 0), 1),
        matieres=matieres,
    )


@router.get("/classes", response_model=List[ClasseSummary])
def get_classes_summary(dept: str = Query("IAID"), db: Session = Depends(get_db)):
    """Synthèse par classe pour le département."""
    enseignements = (
        db.query(Enseignement)
        .join(Classe)
        .filter(Classe.departement_code == dept.upper())
        .all()
    )

    by_class: dict = defaultdict(list)
    for ens in enseignements:
        by_class[ens.classe.nom].append(_compute_matiere(ens))

    result = []
    for classe_nom, matieres in by_class.items():
        vhp_t = sum(m.vhp for m in matieres)
        vhr_t = sum(m.vhr for m in matieres)
        result.append(ClasseSummary(
            classe=classe_nom,
            nb_matieres=len(matieres),
            taux_moy=round((vhr_t / vhp_t) if vhp_t > 0 else 0.0, 4),
            vhp_total=round(vhp_t, 1),
            vhr_total=round(vhr_t, 1),
            retard_h=round(sum(m.ecart for m in matieres if m.ecart < 0), 1),
            nb_terminees=sum(1 for m in matieres if m.statut == "Terminé"),
            nb_en_cours=sum(1 for m in matieres if m.statut == "En cours"),
            nb_non_demarrees=sum(1 for m in matieres if m.statut == "Non démarré"),
        ))
    return sorted(result, key=lambda c: c.taux_moy)


@router.get("/enseignants", response_model=List[EnseignantSummary])
def get_enseignants_summary(dept: str = Query("IAID"), db: Session = Depends(get_db)):
    """Synthèse par enseignant avec simulation de charge."""
    enseignements = (
        db.query(Enseignement)
        .join(Classe)
        .filter(Classe.departement_code == dept.upper())
        .all()
    )

    by_prof: dict = defaultdict(list)
    for ens in enseignements:
        key = (ens.responsable or "Non affecté", ens.email or "")
        by_prof[key].append(_compute_matiere(ens))

    SEUIL_CHARGE = 300  # heures/semestre raisonnable
    result = []
    for (prof, email), matieres in by_prof.items():
        vhp_t = sum(m.vhp for m in matieres)
        vhr_t = sum(m.vhr for m in matieres)
        classes_uniq = len({m.classe for m in matieres})

        if vhp_t > SEUIL_CHARGE * 1.2:
            charge_label = "🔴 Surcharge"
        elif vhp_t > SEUIL_CHARGE:
            charge_label = "🟡 Charge élevée"
        else:
            charge_label = "🟢 Charge normale"

        result.append(EnseignantSummary(
            responsable=prof,
            email=email,
            nb_matieres=len(matieres),
            nb_classes=classes_uniq,
            vhp_total=round(vhp_t, 1),
            vhr_total=round(vhr_t, 1),
            taux_moy=round((vhr_t / vhp_t) if vhp_t > 0 else 0.0, 4),
            retard_h=round(sum(m.ecart for m in matieres if m.ecart < 0), 1),
            nb_non_demarrees=sum(1 for m in matieres if m.statut == "Non démarré"),
            charge_label=charge_label,
        ))
    return sorted(result, key=lambda e: e.taux_moy)


@router.get("/mensuel")
def get_mensuel(dept: str = Query("IAID"), db: Session = Depends(get_db)):
    """Heures totales par mois pour le département."""
    enseignements = (
        db.query(Enseignement)
        .join(Classe)
        .filter(Classe.departement_code == dept.upper())
        .all()
    )

    totaux: dict = defaultdict(float)
    by_classe: dict = defaultdict(lambda: defaultdict(float))

    for ens in enseignements:
        for p in ens.pointages:
            try:
                mois_num = int(p.date[5:7])
            except Exception:
                continue
            col = MOIS_ACADEMIQUE.get(mois_num)
            if col:
                totaux[col] += p.heures
                by_classe[ens.classe.nom][col] += p.heures

    ordre_mois = ["oct", "nov", "dec", "jan", "fev", "mars", "avril", "mai", "juin", "juil", "aout"]
    return {
        "totaux": {m: round(totaux.get(m, 0), 1) for m in ordre_mois},
        "par_classe": {
            cls: {m: round(vals.get(m, 0), 1) for m in ordre_mois}
            for cls, vals in by_classe.items()
        },
    }


@router.get("/simulation-charge")
def simulate_charge(
    dept: str = Query("IAID"),
    enseignant: str = Query(...),
    nouvelles_matieres: int = Query(0),
    vhp_moyen: float = Query(30.0),
    db: Session = Depends(get_db),
):
    """Simule la charge si l'on assigne de nouvelles matières à un enseignant."""
    SEUIL = 300.0
    enseignements = (
        db.query(Enseignement)
        .join(Classe)
        .filter(
            Classe.departement_code == dept.upper(),
            Enseignement.responsable == enseignant,
        )
        .all()
    )
    charge_actuelle = sum(e.vhp for e in enseignements)
    charge_simulee = charge_actuelle + nouvelles_matieres * vhp_moyen

    return {
        "enseignant": enseignant,
        "charge_actuelle_h": round(charge_actuelle, 1),
        "charge_simulee_h": round(charge_simulee, 1),
        "delta_h": round(nouvelles_matieres * vhp_moyen, 1),
        "seuil_recommande_h": SEUIL,
        "statut_actuel": "surcharge" if charge_actuelle > SEUIL else "ok",
        "statut_simule": "surcharge" if charge_simulee > SEUIL else "ok",
        "avertissement": charge_simulee > SEUIL,
    }
