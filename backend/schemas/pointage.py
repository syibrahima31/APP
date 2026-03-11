from __future__ import annotations
from pydantic import BaseModel, field_validator
import re


class PointageCreate(BaseModel):
    enseignement_id: int
    date: str           # YYYY-MM-DD
    heures: float
    note: str = ""
    saisi_par: str = ""

    @field_validator("date")
    @classmethod
    def validate_date(cls, v: str) -> str:
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", v):
            raise ValueError("La date doit être au format YYYY-MM-DD")
        return v

    @field_validator("heures")
    @classmethod
    def validate_heures(cls, v: float) -> float:
        if v <= 0:
            raise ValueError("Les heures doivent être supérieures à 0")
        return v


class PointageRead(BaseModel):
    id: int
    enseignement_id: int
    date: str
    heures: float
    note: str
    saisi_par: str

    model_config = {"from_attributes": True}


class PointageResume(BaseModel):
    """Résumé de pointage pour une matière (dans le dashboard)."""
    enseignement_id: int
    matiere: str
    classe: str
    responsable: str
    semestre: str = ""
    email: str = ""
    vhp: float
    vhr: float
    taux: float
    ecart: float
    statut: str
    nb_seances: int
    # Heures par mois (Oct..Août)
    oct: float = 0.0
    nov: float = 0.0
    dec: float = 0.0
    jan: float = 0.0
    fev: float = 0.0
    mars: float = 0.0
    avril: float = 0.0
    mai: float = 0.0
    juin: float = 0.0
    juil: float = 0.0
    aout: float = 0.0
