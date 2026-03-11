from __future__ import annotations

from typing import List, Optional
from collections import defaultdict

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session

from db.database import get_db
from db.models import Classe, Enseignement, Pointage
from backend.schemas.pointage import PointageCreate, PointageRead, PointageResume

router = APIRouter()

# Mapping mois calendaire → colonne académique
_MOIS_ACADEMIQUE = {
    10: "oct", 11: "nov", 12: "dec",
    1: "jan", 2: "fev", 3: "mars",
    4: "avril", 5: "mai", 6: "juin",
    7: "juil", 8: "aout",
}


@router.get("/", response_model=List[PointageRead])
def list_pointages(
    enseignement_id: Optional[int] = Query(None),
    db: Session = Depends(get_db),
):
    q = db.query(Pointage)
    if enseignement_id:
        q = q.filter_by(enseignement_id=enseignement_id)
    return q.order_by(Pointage.date.desc()).all()


@router.post("/", response_model=PointageRead, status_code=201)
def create_pointage(payload: PointageCreate, db: Session = Depends(get_db)):
    ens = db.get(Enseignement, payload.enseignement_id)
    if not ens:
        raise HTTPException(status_code=404, detail="Enseignement introuvable")
    p = Pointage(**payload.model_dump())
    db.add(p)
    db.commit()
    db.refresh(p)
    return p


@router.delete("/{pointage_id}", status_code=204)
def delete_pointage(pointage_id: int, db: Session = Depends(get_db)):
    p = db.get(Pointage, pointage_id)
    if not p:
        raise HTTPException(status_code=404, detail="Pointage introuvable")
    db.delete(p)
    db.commit()


@router.get("/resume", response_model=List[PointageResume])
def resume_pointages(dept: str = Query("IAID"), db: Session = Depends(get_db)):
    """
    Résumé complet : pour chaque enseignement du département,
    calcule VHR (depuis les pointages), taux, écart, heures par mois.
    """
    enseignements = (
        db.query(Enseignement)
        .join(Classe)
        .filter(Classe.departement_code == dept.upper())
        .all()
    )

    result = []
    for ens in enseignements:
        heures_par_mois: dict = defaultdict(float)
        total = 0.0
        nb = 0

        for p in ens.pointages:
            try:
                mois_num = int(p.date[5:7])  # extrait MM depuis YYYY-MM-DD
            except Exception:
                continue
            col = _MOIS_ACADEMIQUE.get(mois_num, "")
            if col:
                heures_par_mois[col] += p.heures
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

        result.append(PointageResume(
            enseignement_id=ens.id,
            matiere=ens.matiere,
            classe=ens.classe.nom,
            responsable=ens.responsable or "",
            semestre=ens.semestre or "",
            email=ens.email or "",
            vhp=vhp,
            vhr=round(total, 2),
            taux=round(taux, 4),
            ecart=round(ecart, 2),
            statut=statut,
            nb_seances=nb,
            **{col: round(h, 2) for col, h in heures_par_mois.items()},
        ))

    return result
