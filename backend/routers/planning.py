"""
Router Planning — Programmation et pointage des séances.
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import List, Optional

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session

from db.database import get_db
from db.models import Classe, Enseignement, Pointage, SeancePlanifiee
from backend.schemas.seance import SeanceCreate, SeancePointer, SeanceRead

router = APIRouter()


def _enrich(s: SeancePlanifiee) -> dict:
    """Ajoute les champs dénormalisés pour l'affichage."""
    ens = s.enseignement
    return {
        "id": s.id,
        "enseignement_id": s.enseignement_id,
        "date_prevue": s.date_prevue,
        "heure_debut": s.heure_debut,
        "heure_fin": s.heure_fin,
        "nb_heures": s.nb_heures,
        "salle": s.salle or "",
        "statut": s.statut,
        "pointage_id": s.pointage_id,
        "note": s.note or "",
        "created_by": s.created_by or "",
        "matiere": ens.matiere if ens else "",
        "classe_nom": ens.classe.nom if ens and ens.classe else "",
        "responsable": ens.responsable if ens else "",
    }


@router.get("/", response_model=List[SeanceRead])
def list_seances(
    dept: str = Query("IAID"),
    classe_id: Optional[int] = Query(None),
    enseignement_id: Optional[int] = Query(None),
    date_debut: Optional[date] = Query(None),
    date_fin: Optional[date] = Query(None),
    statut: Optional[str] = Query(None),
    db: Session = Depends(get_db),
):
    """Liste les séances planifiées avec filtres."""
    q = (
        db.query(SeancePlanifiee)
        .join(SeancePlanifiee.enseignement)
        .join(Enseignement.classe)
        .filter(Classe.departement_code == dept.upper())
    )
    if classe_id:
        q = q.filter(Enseignement.classe_id == classe_id)
    if enseignement_id:
        q = q.filter(SeancePlanifiee.enseignement_id == enseignement_id)
    if date_debut:
        q = q.filter(SeancePlanifiee.date_prevue >= date_debut)
    if date_fin:
        q = q.filter(SeancePlanifiee.date_prevue <= date_fin)
    if statut:
        q = q.filter(SeancePlanifiee.statut == statut)

    seances = q.order_by(SeancePlanifiee.date_prevue, SeancePlanifiee.heure_debut).all()
    return [_enrich(s) for s in seances]


@router.get("/semaine", response_model=List[SeanceRead])
def seances_semaine(
    dept: str = Query("IAID"),
    lundi: Optional[date] = Query(None),
    db: Session = Depends(get_db),
):
    """Retourne les séances de la semaine (lundi → dimanche)."""
    if not lundi:
        today = date.today()
        lundi = today - timedelta(days=today.weekday())
    dimanche = lundi + timedelta(days=6)

    q = (
        db.query(SeancePlanifiee)
        .join(SeancePlanifiee.enseignement)
        .join(Enseignement.classe)
        .filter(
            Classe.departement_code == dept.upper(),
            SeancePlanifiee.date_prevue >= lundi,
            SeancePlanifiee.date_prevue <= dimanche,
        )
        .order_by(SeancePlanifiee.date_prevue, SeancePlanifiee.heure_debut)
    )
    return [_enrich(s) for s in q.all()]


@router.post("/", response_model=SeanceRead, status_code=201)
def create_seance(payload: SeanceCreate, db: Session = Depends(get_db)):
    """Crée une séance planifiée."""
    ens = db.get(Enseignement, payload.enseignement_id)
    if not ens:
        raise HTTPException(status_code=404, detail="Enseignement introuvable")

    # Calcul auto des heures si heure_debut/fin fournies
    nb_heures = payload.nb_heures
    try:
        h1 = int(payload.heure_debut.split(":")[0]) + int(payload.heure_debut.split(":")[1]) / 60
        h2 = int(payload.heure_fin.split(":")[0]) + int(payload.heure_fin.split(":")[1]) / 60
        if h2 > h1:
            nb_heures = round(h2 - h1, 2)
    except Exception:
        pass

    s = SeancePlanifiee(
        enseignement_id=payload.enseignement_id,
        date_prevue=payload.date_prevue,
        heure_debut=payload.heure_debut,
        heure_fin=payload.heure_fin,
        nb_heures=nb_heures,
        salle=payload.salle,
        note=payload.note,
        created_by=payload.created_by,
        statut="planifie",
    )
    db.add(s)
    db.commit()
    db.refresh(s)
    return _enrich(s)


@router.post("/{seance_id}/pointer", response_model=SeanceRead)
def pointer_seance(
    seance_id: int,
    body: SeancePointer,
    db: Session = Depends(get_db),
):
    """Marque une séance comme effectuée → crée automatiquement un pointage."""
    s = db.get(SeancePlanifiee, seance_id)
    if not s:
        raise HTTPException(status_code=404, detail="Séance introuvable")
    if s.statut == "effectue":
        raise HTTPException(status_code=409, detail="Séance déjà pointée")

    heures = body.nb_heures if body.nb_heures else s.nb_heures

    # Créer le pointage automatiquement
    pt = Pointage(
        enseignement_id=s.enseignement_id,
        date=str(s.date_prevue),
        heures=heures,
        note=body.note or s.note or "",
        saisi_par=body.saisi_par or s.created_by or "",
    )
    db.add(pt)
    db.flush()  # obtenir l'id

    s.statut = "effectue"
    s.pointage_id = pt.id
    db.commit()
    db.refresh(s)
    return _enrich(s)


@router.post("/{seance_id}/annuler", response_model=SeanceRead)
def annuler_seance(seance_id: int, db: Session = Depends(get_db)):
    """Annule une séance planifiée."""
    s = db.get(SeancePlanifiee, seance_id)
    if not s:
        raise HTTPException(status_code=404, detail="Séance introuvable")
    s.statut = "annule"
    db.commit()
    db.refresh(s)
    return _enrich(s)


@router.delete("/{seance_id}", status_code=204)
def delete_seance(seance_id: int, db: Session = Depends(get_db)):
    s = db.get(SeancePlanifiee, seance_id)
    if not s:
        raise HTTPException(status_code=404, detail="Séance introuvable")
    db.delete(s)
    db.commit()
