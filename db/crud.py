"""Opérations CRUD et chargement vers DataFrame (même format que load_excel_all_sheets)."""
from __future__ import annotations

from collections import defaultdict
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
from sqlalchemy.orm import Session

from db.database import SessionLocal, init_db
from db.models import Classe, Enseignement, Pointage

# Mapping mois calendaire → colonne DataFrame académique
_MOIS_NUM_TO_DF = {
    10: "Oct", 11: "Nov", 12: "Déc",
    1: "Jan",  2: "Fév",  3: "Mars",
    4: "Avril", 5: "Mai", 6: "Juin",
    7: "Juil", 8: "Août",
}

_MOIS_DF_COLS = ["Oct", "Nov", "Déc", "Jan", "Fév", "Mars", "Avril", "Mai", "Juin", "Juil", "Août"]


def _heures_par_mois_from_pointages(pointages: list) -> Dict[str, float]:
    """Calcule les heures par mois académique depuis la liste de pointages."""
    result: Dict[str, float] = defaultdict(float)
    for p in pointages:
        try:
            mois_num = int(p.date[5:7])
        except Exception:
            continue
        col = _MOIS_NUM_TO_DF.get(mois_num)
        if col:
            result[col] += p.heures
    return dict(result)


def load_df_from_db(dept_code: str) -> Tuple[pd.DataFrame, Dict[str, List[str]]]:
    """
    Charge les données d'un département depuis la DB et retourne
    (df, quality_issues) avec le même format que load_excel_all_sheets().
    VHR calculé depuis les pointages.
    """
    init_db()
    quality_issues: Dict[str, List[str]] = {}

    with SessionLocal() as session:
        classes = session.query(Classe).filter_by(departement_code=dept_code.upper()).all()

        if not classes:
            quality_issues["__GLOBAL__"] = [f"Aucune classe trouvée pour le département '{dept_code}'."]
            return pd.DataFrame(), quality_issues

        frames = []
        for classe in classes:
            rows = []
            for ens in classe.enseignements:
                heures_mois = _heures_par_mois_from_pointages(ens.pointages)
                row = {
                    "Classe": classe.nom,
                    "Matière": ens.matiere,
                    "VHP": ens.vhp,
                    "Responsable": ens.responsable or "",
                    "Email": ens.email or "",
                    "Semestre": ens.semestre or "",
                    "Début prévu": ens.debut_prevu or "",
                    "Fin prévue": ens.fin_prevue or "",
                    "Observations": ens.observations or "",
                    "Statut": ens.statut or "",
                    "_db_id": ens.id,
                }
                for col in _MOIS_DF_COLS:
                    row[col] = heures_mois.get(col, 0.0)
                rows.append(row)

            if not rows:
                quality_issues.setdefault(classe.nom, []).append("Aucun enseignement enregistré.")
                continue

            frames.append(pd.DataFrame(rows))

    if not frames:
        return pd.DataFrame(), quality_issues

    all_df = pd.concat(frames, ignore_index=True)

    from utils.data_pipeline import compute_metrics
    all_df = compute_metrics(all_df)
    all_df["_rowid"] = np.arange(len(all_df))

    return all_df, quality_issues


def get_or_create_classe(session: Session, dept_code: str, nom: str) -> Classe:
    """Retourne la classe existante ou en crée une nouvelle."""
    classe = (
        session.query(Classe)
        .filter_by(departement_code=dept_code.upper(), nom=nom)
        .first()
    )
    if classe is None:
        classe = Classe(departement_code=dept_code.upper(), nom=nom)
        session.add(classe)
        session.flush()
    return classe


def upsert_enseignement_from_row(session: Session, classe_id: int, data: dict) -> Enseignement:
    """
    Insère ou met à jour un enseignement depuis une ligne DataFrame.
    Note : les heures mensuelles sont maintenant dans les pointages — cette
    fonction importe les totaux mensuels comme pointages de migration (1 par mois).
    """
    db_id = data.get("_db_id")
    ens = session.get(Enseignement, db_id) if db_id else None

    if ens is None:
        ens = Enseignement(classe_id=classe_id)
        session.add(ens)

    ens.matiere = str(data.get("Matière", ""))
    ens.vhp = float(data.get("VHP", 0) or 0)
    ens.responsable = str(data.get("Responsable", "") or "")
    ens.email = str(data.get("Email", "") or "")
    ens.semestre = str(data.get("Semestre", "") or "")
    ens.debut_prevu = str(data.get("Début prévu", "") or "")
    ens.fin_prevue = str(data.get("Fin prévue", "") or "")
    ens.observations = str(data.get("Observations", "") or "")
    ens.statut = str(data.get("Statut", "") or "")

    session.flush()

    # Créer un pointage de migration pour chaque mois ayant des heures
    _MOIS_DF_TO_DATE = {
        "Oct": "2024-10-01", "Nov": "2024-11-01", "Déc": "2024-12-01",
        "Jan": "2025-01-01", "Fév": "2025-02-01", "Mars": "2025-03-01",
        "Avril": "2025-04-01", "Mai": "2025-05-01", "Juin": "2025-06-01",
        "Juil": "2025-07-01", "Août": "2025-08-01",
    }
    for df_col, date_str in _MOIS_DF_TO_DATE.items():
        heures = float(data.get(df_col, 0) or 0)
        if heures > 0:
            p = Pointage(
                enseignement_id=ens.id,
                date=date_str,
                heures=heures,
                note="Import Excel",
                saisi_par="import",
            )
            session.add(p)

    return ens


def list_classes(dept_code: str) -> List[str]:
    """Retourne la liste des noms de classes pour un département."""
    init_db()
    with SessionLocal() as session:
        classes = session.query(Classe.nom).filter_by(departement_code=dept_code.upper()).all()
        return [c.nom for c in classes]
