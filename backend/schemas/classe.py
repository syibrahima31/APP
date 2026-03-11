from __future__ import annotations
from pydantic import BaseModel


class ClasseCreate(BaseModel):
    departement_code: str
    nom: str


class ClasseRead(BaseModel):
    id: int
    departement_code: str
    nom: str

    model_config = {"from_attributes": True}
