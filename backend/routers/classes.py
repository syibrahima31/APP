from __future__ import annotations

from typing import List

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from db.database import SessionLocal, init_db
from db.models import Classe
from backend.schemas.classe import ClasseCreate, ClasseRead

router = APIRouter()


def get_db():
    init_db()
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


@router.get("/", response_model=List[ClasseRead])
def list_classes(dept: str = "IAID", db: Session = Depends(get_db)):
    """Liste toutes les classes d'un département."""
    return db.query(Classe).filter_by(departement_code=dept.upper()).order_by(Classe.nom).all()


@router.get("/{classe_id}", response_model=ClasseRead)
def get_classe(classe_id: int, db: Session = Depends(get_db)):
    c = db.get(Classe, classe_id)
    if not c:
        raise HTTPException(status_code=404, detail="Classe introuvable")
    return c


@router.post("/", response_model=ClasseRead, status_code=201)
def create_classe(payload: ClasseCreate, db: Session = Depends(get_db)):
    existing = db.query(Classe).filter_by(
        departement_code=payload.departement_code.upper(), nom=payload.nom
    ).first()
    if existing:
        raise HTTPException(status_code=409, detail="Classe déjà existante")
    c = Classe(departement_code=payload.departement_code.upper(), nom=payload.nom)
    db.add(c)
    db.commit()
    db.refresh(c)
    return c


@router.delete("/{classe_id}", status_code=204)
def delete_classe(classe_id: int, db: Session = Depends(get_db)):
    c = db.get(Classe, classe_id)
    if not c:
        raise HTTPException(status_code=404, detail="Classe introuvable")
    db.delete(c)
    db.commit()
