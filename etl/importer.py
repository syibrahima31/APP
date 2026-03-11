"""
ETL Importer — Import Excel idempotent avec SHA-256.
Crée / met à jour / ignore selon hash de chaque ligne.
"""
from __future__ import annotations

import hashlib
import io
import json
import re
from datetime import date, datetime
from typing import Optional

import pandas as pd
from sqlalchemy.orm import Session

from db.models import Classe, Enseignement, Pointage


# ─────────────────────────────────────────────────────────────────────────────
# NORMALISATION DES COLONNES
# ─────────────────────────────────────────────────────────────────────────────

RENAME_MAP: dict[str, str] = {
    # VHR
    "Vhr": "VHR", "VH Réalisé": "VHR", "VH Realise": "VHR",
    "Volume Réalisé": "VHR", "Heures réalisées": "VHR", "Heures Réalisées": "VHR",
    # VHP
    "VHP ": "VHP", "VH Prévu": "VHP", "VH Prevu": "VHP", "VH_Prévu": "VHP",
    "VH Annuel": "VHP", "VH Prévu(H)": "VHP", "Volume Prévu": "VHP",
    "Volume Horaire Prévu": "VHP", "Volume Horaire": "VHP",
    "Heures prévues": "VHP", "Heures Prévues": "VHP",
    "H. Prévu": "VHP", "H Prévu": "VHP",
    # Matière
    "Matiere": "Matière", "Matière ": "Matière", "Matieres": "Matière",
    "Matières": "Matière", "Intitulé": "Matière", "Intitule": "Matière",
    "Intitulé du cours": "Matière", "Module": "Matière", "Libellé": "Matière",
    "Libelle": "Matière", "UE": "Matière", "EC": "Matière",
    "Cours": "Matière", "Enseignement": "Matière",
    # Responsable
    "Responsable ": "Responsable", "Enseignant": "Responsable",
    "Enseignant(e)": "Responsable", "Prof": "Responsable",
    "Professeur": "Responsable", "Nom": "Responsable",
    "Nom enseignant": "Responsable", "Nom Enseignant": "Responsable",
    "Intervenant": "Responsable",
    # Semestre
    "Semestre ": "Semestre", "Semester": "Semestre", "Sem": "Semestre",
    # Observations
    "Observation": "Observations", "Observations ": "Observations",
    "Remarque": "Observations", "Remarques": "Observations",
    "Commentaire": "Observations", "Commentaires": "Observations",
    # Dates
    "Début prévu ": "Début prévu", "Debut prevu": "Début prévu",
    "Début": "Début prévu", "Date début": "Début prévu", "Date Début": "Début prévu",
    "Fin prévue ": "Fin prévue", "Fin prevue": "Fin prévue",
    "Fin": "Fin prévue", "Date fin": "Fin prévue", "Date Fin": "Fin prévue",
    # Email
    "Mail": "Email", "E-mail": "Email", "Email ": "Email",
    "Email enseignant": "Email", "Email Enseignant": "Email", "Courriel": "Email",
    # Taux (colonnes à ignorer pour éviter conflits)
    "Taux (%)": "Taux_excel", "Taux": "Taux_excel",
    "Ecart": "Écart_excel", "Écart": "Écart_excel",
}

MOIS_COLS_MAP: dict[str, tuple[str, int]] = {
    "Oct": ("oct", 10), "Octobre": ("oct", 10),
    "Nov": ("nov", 11), "Novembre": ("nov", 11),
    "Déc": ("dec", 12), "Dec": ("dec", 12), "Décembre": ("dec", 12),
    "Jan": ("jan", 1), "Janvier": ("jan", 1),
    "Fév": ("fev", 2), "Fev": ("fev", 2), "Février": ("fev", 2),
    "Mars": ("mars", 3),
    "Avril": ("avril", 4), "Avr": ("avril", 4),
    "Mai": ("mai", 5),
    "Juin": ("juin", 6),
    "Juil": ("juil", 7), "Juillet": ("juil", 7),
    "Août": ("aout", 8), "Aout": ("aout", 8),
}

ALL_MOIS = ["Oct", "Nov", "Déc", "Jan", "Fév", "Mars", "Avril", "Mai", "Juin", "Juil", "Août"]

# Mapping année académique Oct 2024 → Août 2025
ANNEE_DEFAUT = "2024-2025"
MOIS_TO_DATE: dict[str, str] = {
    "Oct": "2024-10-01", "Nov": "2024-11-01", "Déc": "2024-12-01",
    "Jan": "2025-01-01", "Fév": "2025-02-01", "Mars": "2025-03-01",
    "Avril": "2025-04-01", "Mai": "2025-05-01", "Juin": "2025-06-01",
    "Juil": "2025-07-01", "Août": "2025-08-01",
}


def _annee_to_mois_dates(annee_code: Optional[str]) -> dict[str, str]:
    """Construit le mapping mois→date pour une année académique."""
    if not annee_code:
        return MOIS_TO_DATE
    try:
        y_start = int(annee_code.split("-")[0])
        y_end = y_start + 1
        return {
            "Oct": f"{y_start}-10-01", "Nov": f"{y_start}-11-01", "Déc": f"{y_start}-12-01",
            "Jan": f"{y_end}-01-01", "Fév": f"{y_end}-02-01", "Mars": f"{y_end}-03-01",
            "Avril": f"{y_end}-04-01", "Mai": f"{y_end}-05-01", "Juin": f"{y_end}-06-01",
            "Juil": f"{y_end}-07-01", "Août": f"{y_end}-08-01",
        }
    except Exception:
        return MOIS_TO_DATE


def _clean_col(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).replace("\n", " ").replace('"', "").strip())


def _normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_clean_col(c) for c in df.columns]
    df = df.rename(columns={k: v for k, v in RENAME_MAP.items() if k in df.columns})
    return df


def _find_header_row(xls: pd.ExcelFile, sheet: str) -> pd.DataFrame:
    """Essaie différents header offsets pour trouver les colonnes Matière + VHP."""
    for h in range(6):
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=h)
        except Exception:
            continue
        df_norm = _normalize_df(df)
        if "Matière" in df_norm.columns and "VHP" in df_norm.columns:
            return df_norm
    # Fallback header=0
    return _normalize_df(pd.read_excel(xls, sheet_name=sheet, header=0))


def _to_float(x) -> float:
    if x is None or (isinstance(x, float) and x != x):
        return 0.0
    try:
        return float(str(x).replace(",", ".").strip() or "0")
    except Exception:
        return 0.0


def _compute_hash(row_data: dict) -> str:
    """SHA-256 d'un dictionnaire pour détecter les modifications."""
    canonical = json.dumps(row_data, sort_keys=True, default=str, ensure_ascii=False)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


# ─────────────────────────────────────────────────────────────────────────────
# IMPORTER PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

class ExcelImporter:
    """Import idempotent d'un fichier Excel vers la base de données."""

    def __init__(self, db: Session, dept_code: str, annee_code: Optional[str] = None):
        self.db = db
        self.dept_code = dept_code.upper()
        self.annee_code = annee_code or ANNEE_DEFAUT
        self.mois_dates = _annee_to_mois_dates(annee_code)
        self.stats = {
            "created": 0, "updated": 0, "unchanged": 0,
            "errors": [], "sheets_processed": [], "sheets_skipped": [],
        }

    def clear_department(self):
        """Supprime toutes les données du département (cascade)."""
        classes = self.db.query(Classe).filter_by(departement_code=self.dept_code).all()
        for cls in classes:
            self.db.delete(cls)
        self.db.commit()

    def import_bytes(self, file_bytes: bytes) -> dict:
        """Point d'entrée principal — lit un fichier Excel et l'importe."""
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        for sheet in xls.sheet_names:
            try:
                df = _find_header_row(xls, sheet)
                if "Matière" not in df.columns or "VHP" not in df.columns:
                    self.stats["sheets_skipped"].append(f"{sheet} (colonnes manquantes)")
                    continue
                self._import_sheet(sheet, df)
                self.stats["sheets_processed"].append(sheet)
            except Exception as e:
                self.stats["sheets_skipped"].append(f"{sheet} (erreur: {e})")

        self.db.commit()
        return self.stats

    def _import_sheet(self, classe_nom: str, df: pd.DataFrame):
        """Importe une feuille = une classe."""
        # Ensure month columns exist
        for m in ALL_MOIS:
            if m not in df.columns:
                df[m] = 0.0

        # Get or create Classe
        classe = (
            self.db.query(Classe)
            .filter_by(departement_code=self.dept_code, nom=classe_nom)
            .first()
        )
        if not classe:
            classe = Classe(departement_code=self.dept_code, nom=classe_nom)
            self.db.add(classe)
            self.db.flush()

        for _, row in df.iterrows():
            matiere = str(row.get("Matière", "") or "").strip()
            if not matiere or matiere.lower() in ("nan", "none", ""):
                continue
            vhp = _to_float(row.get("VHP", 0))
            if vhp <= 0:
                continue

            row_data = {
                "matiere": matiere,
                "vhp": round(vhp, 1),
                "responsable": str(row.get("Responsable", "") or "").strip(),
                "email": str(row.get("Email", "") or "").strip().lower(),
                "semestre": self._normalize_semestre(row.get("Semestre", "")),
                "debut_prevu": str(row.get("Début prévu", "") or "").strip(),
                "fin_prevue": str(row.get("Fin prévue", "") or "").strip(),
                "observations": str(row.get("Observations", "") or "").strip(),
                "heures_mois": {
                    m: round(_to_float(row.get(m, 0)), 1) for m in ALL_MOIS
                },
            }
            row_hash = _compute_hash({k: v for k, v in row_data.items() if k != "heures_mois"})

            # Chercher un enseignement existant par matière + classe
            existing = (
                self.db.query(Enseignement)
                .filter_by(classe_id=classe.id, matiere=matiere)
                .first()
            )

            if existing:
                if existing.import_hash == row_hash:
                    self.stats["unchanged"] += 1
                    continue
                # Mise à jour
                existing.vhp         = row_data["vhp"]
                existing.responsable = row_data["responsable"]
                existing.email       = row_data["email"]
                existing.semestre    = row_data["semestre"]
                existing.debut_prevu = row_data["debut_prevu"]
                existing.fin_prevue  = row_data["fin_prevue"]
                existing.observations = row_data["observations"]
                existing.import_hash = row_hash
                self._sync_pointages(existing, row_data["heures_mois"])
                self.stats["updated"] += 1
            else:
                ens = Enseignement(
                    classe_id=classe.id,
                    matiere=matiere,
                    vhp=row_data["vhp"],
                    responsable=row_data["responsable"],
                    email=row_data["email"],
                    semestre=row_data["semestre"],
                    debut_prevu=row_data["debut_prevu"],
                    fin_prevue=row_data["fin_prevue"],
                    observations=row_data["observations"],
                    import_hash=row_hash,
                )
                self.db.add(ens)
                self.db.flush()
                self._sync_pointages(ens, row_data["heures_mois"])
                self.stats["created"] += 1

    def _sync_pointages(self, ens: Enseignement, heures_mois: dict):
        """
        Synchronise les pointages mensuel d'import.
        Supprime les anciens pointages d'import et recrée uniquement les mois avec heures.
        """
        # Supprimer les anciens pointages d'import (saisi_par="import")
        self.db.query(Pointage).filter_by(
            enseignement_id=ens.id, saisi_par="import"
        ).delete(synchronize_session="fetch")

        # Créer les nouveaux
        for mois_col, heures in heures_mois.items():
            if heures > 0:
                date_str = self.mois_dates.get(mois_col)
                if date_str:
                    self.db.add(Pointage(
                        enseignement_id=ens.id,
                        date=date_str,
                        heures=heures,
                        note=f"Import Excel — {self.annee_code}",
                        saisi_par="import",
                    ))

    @staticmethod
    def _normalize_semestre(x) -> str:
        if x is None or (isinstance(x, float) and x != x):
            return ""
        s = str(x).strip().upper()
        if s.isdigit():
            return f"S{int(s)}"
        s = s.replace("SEMESTRE", "S").replace("SEM", "S")
        m = re.search(r"S\s*0*([1-9]\d*)", s)
        if m:
            return f"S{int(m.group(1))}"
        return s[:20] if s not in ("NAN", "NONE", "") else ""
