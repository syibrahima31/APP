"""
Router Import Excel — Import idempotent depuis un fichier uploadé.
"""
from __future__ import annotations

from typing import Optional

from fastapi import APIRouter, Depends, File, Form, HTTPException, Query, UploadFile
from pydantic import BaseModel
from sqlalchemy.orm import Session

from db.database import get_db

router = APIRouter()


class ImportResult(BaseModel):
    dept: str
    created: int
    updated: int
    unchanged: int
    errors: list
    sheets_processed: list[str]
    sheets_skipped: list[str]


@router.post("/excel", response_model=ImportResult)
async def import_excel(
    file: UploadFile = File(...),
    dept: str = Form(...),
    annee: Optional[str] = Form(None),
    clear_existing: bool = Form(False),
    db: Session = Depends(get_db),
):
    """
    Importe un fichier Excel dans la base de données.
    - Idempotent : détecte et ignore les lignes inchangées (hash SHA-256)
    - Met à jour uniquement les lignes modifiées
    - Crée les nouveaux enseignements
    - Optionnel : clear_existing supprime tout avant import
    """
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Format accepté : .xlsx ou .xls uniquement")

    content = await file.read()
    if len(content) > 20 * 1024 * 1024:  # 20 MB max
        raise HTTPException(status_code=413, detail="Fichier trop volumineux (max 20 MB)")

    try:
        from etl.importer import ExcelImporter
        importer = ExcelImporter(db=db, dept_code=dept.upper(), annee_code=annee)

        if clear_existing:
            importer.clear_department()

        stats = importer.import_bytes(content)
        return ImportResult(
            dept=dept.upper(),
            created=stats["created"],
            updated=stats["updated"],
            unchanged=stats["unchanged"],
            errors=stats["errors"],
            sheets_processed=stats.get("sheets_processed", []),
            sheets_skipped=stats.get("sheets_skipped", []),
        )
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Erreur d'import : {str(e)}")


@router.post("/excel-url")
async def import_from_url(
    url: str = Form(...),
    dept: str = Form(...),
    annee: Optional[str] = Form(None),
    clear_existing: bool = Form(False),
    db: Session = Depends(get_db),
):
    """Importe depuis une URL Google Sheets / Drive (fichier Excel)."""
    import requests as req
    try:
        r = req.get(url.strip(), timeout=45, headers={"Cache-Control": "no-cache"})
        r.raise_for_status()
        content = r.content
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Impossible de télécharger le fichier : {e}")

    try:
        from etl.importer import ExcelImporter
        importer = ExcelImporter(db=db, dept_code=dept.upper(), annee_code=annee)
        if clear_existing:
            importer.clear_department()
        stats = importer.import_bytes(content)
        return ImportResult(
            dept=dept.upper(),
            created=stats["created"],
            updated=stats["updated"],
            unchanged=stats["unchanged"],
            errors=stats["errors"],
            sheets_processed=stats.get("sheets_processed", []),
            sheets_skipped=stats.get("sheets_skipped", []),
        )
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Erreur d'import : {str(e)}")
