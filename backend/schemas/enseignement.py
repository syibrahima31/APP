from __future__ import annotations
from typing import Optional
from pydantic import BaseModel


class EnseignementCreate(BaseModel):
    classe_id: int
    matiere: str
    vhp: float = 0.0
    responsable: str = ""
    email: str = ""
    semestre: str = ""
    debut_prevu: str = ""
    fin_prevue: str = ""
    observations: str = ""
    statut: str = ""


class EnseignementUpdate(BaseModel):
    matiere: Optional[str] = None
    vhp: Optional[float] = None
    responsable: Optional[str] = None
    email: Optional[str] = None
    semestre: Optional[str] = None
    debut_prevu: Optional[str] = None
    fin_prevue: Optional[str] = None
    observations: Optional[str] = None
    statut: Optional[str] = None


class EnseignementRead(BaseModel):
    id: int
    classe_id: int
    matiere: str
    vhp: float
    responsable: str
    email: str
    semestre: str
    debut_prevu: str
    fin_prevue: str
    observations: str
    statut: str
    # Heures réalisées (calculées depuis les pointages)
    vhr: float = 0.0
    taux: float = 0.0
    ecart: float = 0.0

    model_config = {"from_attributes": True}
