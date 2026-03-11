"""
Modèles SQLAlchemy — Dashboard ISI v3
Schéma étendu : Enseignant, AnneeScolaire, AuditLog, import_hash
"""
from __future__ import annotations

from datetime import datetime

from sqlalchemy import (
    Boolean, Column, Date, DateTime, Float, ForeignKey,
    Integer, JSON, String, Text, UniqueConstraint, Index
)
from sqlalchemy.orm import relationship

from db.database import Base


# ─────────────────────────────────────────────────────────────────────────────
# RÉFÉRENTIELS
# ─────────────────────────────────────────────────────────────────────────────

class Departement(Base):
    """Département universitaire (IAID, KM, DRS, GI, ...)."""
    __tablename__ = "departements"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    code            = Column(String(10), unique=True, nullable=False, index=True)
    nom             = Column(String(100), nullable=False)
    chef_dept       = Column(String(100), default="")
    email_chef      = Column(String(150), default="")
    actif           = Column(Boolean, default=True)
    created_at      = Column(DateTime, default=datetime.utcnow)

    annees          = relationship("AnneeScolaire", back_populates="departement", cascade="all, delete-orphan")
    enseignants     = relationship("Enseignant", back_populates="departement")
    classes         = relationship("Classe", back_populates="departement")

    def __repr__(self):
        return f"<Departement {self.code}>"


class AnneeScolaire(Base):
    """Année académique (ex: 2024-2025)."""
    __tablename__ = "annees_scolaires"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    code            = Column(String(9), nullable=False)       # '2024-2025'
    date_debut      = Column(Date, nullable=False)
    date_fin        = Column(Date, nullable=False)
    est_active      = Column(Boolean, default=False)
    departement_id  = Column(Integer, ForeignKey("departements.id", ondelete="CASCADE"), nullable=True)

    departement     = relationship("Departement", back_populates="annees")

    __table_args__ = (
        UniqueConstraint("departement_id", "code", name="uq_annee_dept_code"),
    )

    def __repr__(self):
        return f"<AnneeScolaire {self.code}>"


class Enseignant(Base):
    """Enseignant (entité propre, désolidarisée du string dans Enseignement)."""
    __tablename__ = "enseignants"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    nom             = Column(String(150), nullable=False)
    email           = Column(String(150), default="", index=True)
    departement_id  = Column(Integer, ForeignKey("departements.id", ondelete="SET NULL"), nullable=True)
    actif           = Column(Boolean, default=True)
    created_at      = Column(DateTime, default=datetime.utcnow)

    departement     = relationship("Departement", back_populates="enseignants")
    enseignements   = relationship("Enseignement", back_populates="enseignant")

    def __repr__(self):
        return f"<Enseignant {self.nom}>"


# ─────────────────────────────────────────────────────────────────────────────
# DONNÉES PÉDAGOGIQUES
# ─────────────────────────────────────────────────────────────────────────────

class Classe(Base):
    """Classe / section pédagogique (= une feuille Excel)."""
    __tablename__ = "classes"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    # Clé legacy (rétro-compatibilité)
    departement_code = Column(String(20), nullable=False, index=True)
    nom             = Column(String(100), nullable=False)
    niveau          = Column(String(20), default="")    # L1, L2, M1, M2 …
    # Clé FK vers la table Departement (optionnelle — peuplée par migration)
    departement_id  = Column(Integer, ForeignKey("departements.id", ondelete="SET NULL"), nullable=True)
    annee_id        = Column(Integer, ForeignKey("annees_scolaires.id", ondelete="SET NULL"), nullable=True)

    __table_args__ = (
        UniqueConstraint("departement_code", "nom", name="uq_classe_dept_nom"),
    )

    departement     = relationship("Departement", back_populates="classes")
    enseignements   = relationship("Enseignement", back_populates="classe", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Classe {self.departement_code}/{self.nom}>"


class Enseignement(Base):
    """
    Ligne de suivi : matière + heures planifiées pour une classe.
    import_hash : SHA-256 de la ligne Excel source → idempotence.
    """
    __tablename__ = "enseignements"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    classe_id       = Column(Integer, ForeignKey("classes.id", ondelete="CASCADE"), nullable=False, index=True)
    enseignant_id   = Column(Integer, ForeignKey("enseignants.id", ondelete="SET NULL"), nullable=True, index=True)

    # Champs pédagogiques
    matiere         = Column(String(200), nullable=False, default="")
    vhp             = Column(Float, nullable=False, default=0.0)   # Volume Horaire Prévu
    responsable     = Column(String(200), default="")              # Conservé pour rétro-compat.
    email           = Column(String(200), default="")
    semestre        = Column(String(20), default="")
    debut_prevu     = Column(String(50), default="")
    fin_prevue      = Column(String(50), default="")
    observations    = Column(Text, default="")
    statut          = Column(String(50), default="")

    # Idempotence import Excel
    import_hash     = Column(String(64), default="", index=True)   # SHA-256

    created_at      = Column(DateTime, default=datetime.utcnow)
    updated_at      = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    classe          = relationship("Classe", back_populates="enseignements")
    enseignant      = relationship("Enseignant", back_populates="enseignements")
    pointages       = relationship("Pointage", back_populates="enseignement", cascade="all, delete-orphan")
    alertes         = relationship("Alerte", back_populates="enseignement", cascade="all, delete-orphan")
    seances         = relationship("SeancePlanifiee", back_populates="enseignement", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Enseignement {self.matiere} (classe_id={self.classe_id})>"


class Pointage(Base):
    """Enregistrement d'une séance réalisée."""
    __tablename__ = "pointages"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    enseignement_id = Column(
        Integer, ForeignKey("enseignements.id", ondelete="CASCADE"), nullable=False, index=True
    )
    date            = Column(String(10), nullable=False)   # YYYY-MM-DD
    heures          = Column(Float, nullable=False, default=0.0)
    note            = Column(Text, default="")
    saisi_par       = Column(String(100), default="")
    created_at      = Column(DateTime, default=datetime.utcnow)

    enseignement    = relationship("Enseignement", back_populates="pointages")
    seance          = relationship("SeancePlanifiee", back_populates="pointage", uselist=False)

    __table_args__ = (
        Index("idx_pointage_ens_date", "enseignement_id", "date"),
    )

    def __repr__(self):
        return f"<Pointage {self.date} {self.heures}h>"


class SeancePlanifiee(Base):
    """Séance de cours programmée (planning prévisionnel)."""
    __tablename__ = "seances_planifiees"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    enseignement_id = Column(Integer, ForeignKey("enseignements.id", ondelete="CASCADE"), nullable=False, index=True)
    date_prevue     = Column(Date, nullable=False, index=True)
    heure_debut     = Column(String(5), default="08:00")   # HH:MM
    heure_fin       = Column(String(5), default="10:00")   # HH:MM
    nb_heures       = Column(Float, nullable=False, default=2.0)
    salle           = Column(String(80), default="")
    statut          = Column(String(20), default="planifie")  # planifie | effectue | annule
    pointage_id     = Column(Integer, ForeignKey("pointages.id", ondelete="SET NULL"), nullable=True)
    note            = Column(Text, default="")
    created_by      = Column(String(100), default="")
    created_at      = Column(DateTime, default=datetime.utcnow)

    enseignement    = relationship("Enseignement", back_populates="seances")
    pointage        = relationship("Pointage", back_populates="seance", foreign_keys=[pointage_id])

    __table_args__ = (
        Index("idx_seance_date", "date_prevue", "statut"),
        Index("idx_seance_ens", "enseignement_id", "date_prevue"),
    )

    def __repr__(self):
        return f"<SeancePlanifiee {self.date_prevue} {self.heure_debut}-{self.heure_fin}>"


# ─────────────────────────────────────────────────────────────────────────────
# ALERTES & AUDIT
# ─────────────────────────────────────────────────────────────────────────────

class Alerte(Base):
    """Alerte générée automatiquement sur une matière."""
    __tablename__ = "alertes"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    enseignement_id = Column(Integer, ForeignKey("enseignements.id", ondelete="CASCADE"), nullable=False, index=True)
    type_alerte     = Column(String(50), nullable=False)   # non_demarre_tardif, ecart_critique, retard_modere, depassement
    message         = Column(Text, default="")
    niveau          = Column(String(10), default="warning") # critical | warning | info
    est_lue         = Column(Boolean, default=False)
    action_requise  = Column(String(50), default="flag_only")  # notify_chef | notify_enseignant | flag_only
    created_at      = Column(DateTime, default=datetime.utcnow)

    enseignement    = relationship("Enseignement", back_populates="alertes")

    __table_args__ = (
        Index("idx_alerte_lue", "est_lue", "created_at"),
        Index("idx_alerte_ens", "enseignement_id", "type_alerte"),
    )

    def __repr__(self):
        return f"<Alerte {self.type_alerte} [{self.niveau}]>"


class AuditLog(Base):
    """Journal d'audit de toutes les mutations sensibles."""
    __tablename__ = "audit_logs"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    utilisateur     = Column(String(100), default="system")
    action          = Column(String(50), nullable=False)  # create | update | delete | export | login | import
    table_cible     = Column(String(50), default="")
    ligne_id        = Column(Integer, nullable=True)
    avant           = Column(JSON, nullable=True)
    apres           = Column(JSON, nullable=True)
    ip              = Column(String(45), default="")
    details         = Column(Text, default="")
    created_at      = Column(DateTime, default=datetime.utcnow)

    __table_args__ = (
        Index("idx_audit_action", "action", "created_at"),
        Index("idx_audit_user", "utilisateur", "created_at"),
    )

    def __repr__(self):
        return f"<AuditLog {self.action} par {self.utilisateur}>"


# ─────────────────────────────────────────────────────────────────────────────
# UTILISATEURS (Auth JWT)
# ─────────────────────────────────────────────────────────────────────────────

class Utilisateur(Base):
    """Compte utilisateur pour l'authentification JWT."""
    __tablename__ = "utilisateurs"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    username        = Column(String(80), unique=True, nullable=False, index=True)
    email           = Column(String(150), unique=True, nullable=False)
    hashed_password = Column(String(200), nullable=False)
    role            = Column(String(20), nullable=False, default="viewer")  # admin | chef_dept | viewer
    departement_code = Column(String(10), default="")
    est_actif       = Column(Boolean, default=True)
    created_at      = Column(DateTime, default=datetime.utcnow)
    last_login      = Column(DateTime, nullable=True)

    def __repr__(self):
        return f"<Utilisateur {self.username} [{self.role}]>"
