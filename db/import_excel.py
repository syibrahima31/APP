"""
Script de migration : importe un fichier Excel multi-feuilles dans la base de données.

Usage:
    python db/import_excel.py --dept IAID --file chemin/vers/fichier.xlsx
    python db/import_excel.py --dept KM   --file chemin/vers/fichier.xlsx --clear
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Ajouter la racine du projet au path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd

from db.database import init_db, SessionLocal
from db.models import Classe, Enseignement
from db.crud import get_or_create_classe, upsert_enseignement_from_row
from utils.data_pipeline import load_excel_all_sheets

_MOIS_DF_COLS = ["Oct", "Nov", "Déc", "Jan", "Fév", "Mars", "Avril", "Mai", "Juin", "Juil", "Août"]


def import_excel_to_db(dept_code: str, file_path: str, clear_existing: bool = False) -> None:
    """Importe un fichier Excel dans la base de données pour un département donné."""
    init_db()

    file_bytes = Path(file_path).read_bytes()
    df, quality = load_excel_all_sheets(file_bytes)

    if quality:
        print("⚠️  Problèmes qualité détectés :")
        for sheet, issues in quality.items():
            for issue in issues:
                print(f"   [{sheet}] {issue}")

    if df.empty:
        print("❌ Aucune donnée exploitable dans le fichier Excel.")
        return

    with SessionLocal() as session:
        if clear_existing:
            print(f"🗑️  Suppression des données existantes pour le département '{dept_code}'...")
            classes = session.query(Classe).filter_by(departement_code=dept_code.upper()).all()
            for c in classes:
                session.delete(c)
            session.commit()

        classes_traitees = set()
        nb_inseres = 0

        for classe_nom, group in df.groupby("Classe"):
            classe = get_or_create_classe(session, dept_code, str(classe_nom))
            classes_traitees.add(classe_nom)

            for _, row in group.iterrows():
                matiere = str(row.get("Matière", "")).strip()
                if not matiere or matiere.lower() in ("nan", ""):
                    continue

                data = {
                    "Matière": matiere,
                    "VHP": row.get("VHP", 0),
                    "Responsable": row.get("Responsable", ""),
                    "Email": row.get("Email", ""),
                    "Semestre": row.get("Semestre", ""),
                    "Début prévu": row.get("Début prévu", ""),
                    "Fin prévue": row.get("Fin prévue", ""),
                    "Observations": row.get("Observations", ""),
                    "Statut": row.get("Statut", ""),
                }
                for m in _MOIS_DF_COLS:
                    data[m] = row.get(m, 0)

                upsert_enseignement_from_row(session, classe.id, data)
                nb_inseres += 1

        session.commit()

    print(f"✅ Import terminé : {nb_inseres} enseignements importés dans {len(classes_traitees)} classe(s).")
    print(f"   Classes : {', '.join(sorted(classes_traitees))}")


def main():
    parser = argparse.ArgumentParser(description="Importer un fichier Excel dans la base de données.")
    parser.add_argument("--dept", required=True, help="Code département (IAID, KM, DRS, GI)")
    parser.add_argument("--file", required=True, help="Chemin vers le fichier Excel (.xlsx)")
    parser.add_argument("--clear", action="store_true", help="Supprimer les données existantes avant import")
    args = parser.parse_args()

    if not Path(args.file).exists():
        print(f"❌ Fichier introuvable : {args.file}")
        sys.exit(1)

    import_excel_to_db(dept_code=args.dept, file_path=args.file, clear_existing=args.clear)


if __name__ == "__main__":
    main()
