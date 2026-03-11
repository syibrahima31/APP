"""
Gestion de la connexion SQLAlchemy — Dashboard ISI v3
Supporte SQLite (dev) et PostgreSQL (prod) via DB_URL.
"""
from __future__ import annotations

import os
from pathlib import Path

from sqlalchemy import create_engine, event, text
from sqlalchemy.orm import sessionmaker, DeclarativeBase

_DEFAULT_DB = Path(__file__).parent.parent / "data" / "dashboard.db"


def get_db_url() -> str:
    url = os.environ.get("DB_URL", "")
    if url:
        return url
    db_path = Path(os.environ.get("DB_PATH", str(_DEFAULT_DB)))
    db_path.parent.mkdir(parents=True, exist_ok=True)
    return f"sqlite:///{db_path}"


_db_url = get_db_url()
_is_sqlite = _db_url.startswith("sqlite")

engine = create_engine(
    _db_url,
    connect_args={"check_same_thread": False} if _is_sqlite else {},
    echo=False,
    pool_pre_ping=not _is_sqlite,   # Inutile pour SQLite
    pool_recycle=3600,
    pool_size=10,                   # Pool de connexions réutilisées
    max_overflow=20,
)

# Active les foreign keys sur SQLite (désactivées par défaut)
if _is_sqlite:
    @event.listens_for(engine, "connect")
    def _enable_fk(dbapi_con, _):
        dbapi_con.execute("PRAGMA foreign_keys = ON")
        dbapi_con.execute("PRAGMA journal_mode = WAL")   # Meilleures perfs concurrentes

SessionLocal = sessionmaker(bind=engine, autocommit=False, autoflush=False)


class Base(DeclarativeBase):
    pass


def get_db():
    """Générateur de session pour l'injection de dépendances FastAPI."""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db() -> None:
    """Crée toutes les tables (idempotent — ne supprime pas les données)."""
    from db import models  # noqa: F401 — nécessaire pour enregistrer les modèles
    Base.metadata.create_all(bind=engine)
    _seed_default_data()


def _seed_default_data() -> None:
    """Insère les données de base (départements, admin) si la DB est vide."""
    from db.models import Departement, Utilisateur
    from backend.core.security import hash_password

    db = SessionLocal()
    try:
        # Départements par défaut
        dept_data = [
            ("IAID", "IA & Ingénierie des Données", "Ibrahima SY", "ibsy@groupeisi.com"),
            ("KM",   "Directions des Études",       "Mouhamed Gueye", "mgueye@groupeisi.com"),
            ("DRS",  "Réseaux et Systèmes",          "Latyr Ndiaye",   "landiaye@groupeisi.com"),
            ("GI",   "Génie Informatique",            "El Hadji Mor Diaw", "emdiaw@groupeisi.com"),
        ]
        for code, nom, chef, email in dept_data:
            if not db.query(Departement).filter_by(code=code).first():
                db.add(Departement(code=code, nom=nom, chef_dept=chef, email_chef=email))

        # Compte admin par défaut (à changer en production !)
        if not db.query(Utilisateur).filter_by(username="admin").first():
            db.add(Utilisateur(
                username="admin",
                email="admin@groupeisi.com",
                hashed_password=hash_password("changeme"),
                role="admin",
                departement_code="IAID",
            ))

        db.commit()
    except Exception:
        db.rollback()
    finally:
        db.close()
