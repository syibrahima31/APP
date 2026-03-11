from __future__ import annotations
from datetime import date
from typing import Optional
from pydantic import BaseModel


class SeanceCreate(BaseModel):
    enseignement_id: int
    date_prevue: date
    heure_debut: str = "08:00"
    heure_fin: str = "10:00"
    nb_heures: float = 2.0
    salle: str = ""
    note: str = ""
    created_by: str = ""


class SeanceRead(BaseModel):
    id: int
    enseignement_id: int
    date_prevue: date
    heure_debut: str
    heure_fin: str
    nb_heures: float
    salle: str
    statut: str
    pointage_id: Optional[int]
    note: str
    created_by: str
    # Champs dénormalisés pour l'affichage
    matiere: Optional[str] = None
    classe_nom: Optional[str] = None
    responsable: Optional[str] = None

    model_config = {"from_attributes": True}


class SeancePointer(BaseModel):
    saisi_par: str = ""
    note: str = ""
    nb_heures: Optional[float] = None  # Si None → utilise nb_heures de la séance
