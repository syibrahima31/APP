"""
Configuration centralisée — Dashboard ISI v3
Toutes les valeurs sensibles viennent de variables d'environnement ou .env
"""
from __future__ import annotations

import os
from functools import lru_cache
from typing import List


class Settings:
    """
    Configuration de l'application.
    Lit depuis les variables d'environnement avec fallbacks sécurisés.
    """

    # ── Base de données ────────────────────────────────────────────────────
    DATABASE_URL: str = os.getenv("DB_URL", "")
    DB_PATH: str = os.getenv("DB_PATH", "data/dashboard.db")

    # ── Sécurité ───────────────────────────────────────────────────────────
    SECRET_KEY: str = os.getenv(
        "SECRET_KEY",
        "isi-dashboard-dev-secret-key-CHANGE-IN-PRODUCTION-2024"
    )
    ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_HOURS: int = int(os.getenv("ACCESS_TOKEN_EXPIRE_HOURS", "8"))
    REFRESH_TOKEN_EXPIRE_DAYS: int = int(os.getenv("REFRESH_TOKEN_EXPIRE_DAYS", "7"))

    # ── CORS ───────────────────────────────────────────────────────────────
    # En prod : "https://dashboard.isi.sn,https://autre-domaine.com"
    _cors_raw: str = os.getenv("ALLOWED_ORIGINS", "http://localhost:8501,http://localhost:8502")

    @property
    def ALLOWED_ORIGINS(self) -> List[str]:
        return [o.strip() for o in self._cors_raw.split(",") if o.strip()]

    # ── OpenAI ─────────────────────────────────────────────────────────────
    OPENAI_API_KEY: str = os.getenv("OPENAI_API_KEY", "")
    OPENAI_MODEL: str = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

    # ── SMTP Email ─────────────────────────────────────────────────────────
    SMTP_HOST: str = os.getenv("SMTP_HOST", "smtp.gmail.com")
    SMTP_PORT: int = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER: str = os.getenv("SMTP_USER", "")
    SMTP_PASS: str = os.getenv("SMTP_PASS", "")
    SMTP_FROM: str = os.getenv("SMTP_FROM", "")

    # ── Application ────────────────────────────────────────────────────────
    APP_NAME: str = "Dashboard ISI — Suivi des Enseignements"
    APP_VERSION: str = "3.0.0"
    DEBUG: bool = os.getenv("DEBUG", "false").lower() == "true"
    API_BASE_URL: str = os.getenv("API_BASE_URL", "http://localhost:8000")

    # ── Seuils métier par défaut ───────────────────────────────────────────
    TAUX_VERT: float = float(os.getenv("TAUX_VERT", "0.90"))
    TAUX_ORANGE: float = float(os.getenv("TAUX_ORANGE", "0.60"))
    ECART_CRITIQUE: float = float(os.getenv("ECART_CRITIQUE", "-6"))

    def is_prod(self) -> bool:
        return not self.DEBUG and "localhost" not in self.API_BASE_URL

    def get_db_url(self) -> str:
        if self.DATABASE_URL:
            return self.DATABASE_URL
        from pathlib import Path
        db_path = Path(self.DB_PATH)
        db_path.parent.mkdir(parents=True, exist_ok=True)
        return f"sqlite:///{db_path}"


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    """Retourne l'instance singleton des settings (cachée)."""
    return Settings()


settings = get_settings()
