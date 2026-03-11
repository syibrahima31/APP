from __future__ import annotations

from typing import List, Optional

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session

from db.database import get_db
from db.models import Classe, Enseignement, Pointage
from backend.schemas.enseignement import EnseignementCreate, EnseignementRead, EnseignementUpdate

router = APIRouter()


def _compute_vhr(ens: Enseignement) -> float:
    return sum(p.heures for p in ens.pointages)


def _to_read(ens: Enseignement) -> EnseignementRead:
    vhr = _compute_vhr(ens)
    taux = (vhr / ens.vhp) if ens.vhp > 0 else 0.0
    ecart = vhr - ens.vhp
    return EnseignementRead(
        id=ens.id,
        classe_id=ens.classe_id,
        matiere=ens.matiere,
        vhp=ens.vhp,
        responsable=ens.responsable or "",
        email=ens.email or "",
        semestre=ens.semestre or "",
        debut_prevu=ens.debut_prevu or "",
        fin_prevue=ens.fin_prevue or "",
        observations=ens.observations or "",
        statut=ens.statut or "",
        vhr=round(vhr, 2),
        taux=round(taux, 4),
        ecart=round(ecart, 2),
    )


@router.get("/", response_model=List[EnseignementRead])
def list_enseignements(
    dept: str = Query("IAID"),
    classe_id: Optional[int] = Query(None),
    db: Session = Depends(get_db),
):
    """Liste les enseignements, filtrables par département et/ou classe."""
    q = (
        db.query(Enseignement)
        .join(Classe)
        .filter(Classe.departement_code == dept.upper())
    )
    if classe_id:
        q = q.filter(Enseignement.classe_id == classe_id)
    return [_to_read(e) for e in q.all()]


@router.get("/{ens_id}", response_model=EnseignementRead)
def get_enseignement(ens_id: int, db: Session = Depends(get_db)):
    ens = db.get(Enseignement, ens_id)
    if not ens:
        raise HTTPException(status_code=404, detail="Enseignement introuvable")
    return _to_read(ens)


@router.post("/", response_model=EnseignementRead, status_code=201)
def create_enseignement(payload: EnseignementCreate, db: Session = Depends(get_db)):
    if not db.get(Classe, payload.classe_id):
        raise HTTPException(status_code=404, detail="Classe introuvable")
    ens = Enseignement(**payload.model_dump())
    db.add(ens)
    db.commit()
    db.refresh(ens)
    return _to_read(ens)


@router.put("/{ens_id}", response_model=EnseignementRead)
def update_enseignement(ens_id: int, payload: EnseignementUpdate, db: Session = Depends(get_db)):
    ens = db.get(Enseignement, ens_id)
    if not ens:
        raise HTTPException(status_code=404, detail="Enseignement introuvable")
    for field, value in payload.model_dump(exclude_none=True).items():
        setattr(ens, field, value)
    db.commit()
    db.refresh(ens)
    return _to_read(ens)


@router.delete("/{ens_id}", status_code=204)
def delete_enseignement(ens_id: int, db: Session = Depends(get_db)):
    ens = db.get(Enseignement, ens_id)
    if not ens:
        raise HTTPException(status_code=404, detail="Enseignement introuvable")
    db.delete(ens)
    db.commit()
