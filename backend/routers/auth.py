"""
Router Auth — Login JWT, refresh token, profil utilisateur.
"""
from __future__ import annotations

from datetime import datetime

from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.security import OAuth2PasswordRequestForm
from pydantic import BaseModel, EmailStr
from sqlalchemy.orm import Session

from backend.core.deps import get_current_user
from backend.core.security import (
    create_access_token, create_refresh_token,
    decode_token, hash_password, verify_password,
)
from db.database import get_db
from db.models import AuditLog, Utilisateur

router = APIRouter(prefix="/api/auth", tags=["Authentification"])


# ── Schémas ───────────────────────────────────────────────────────────────

class TokenResponse(BaseModel):
    access_token: str
    refresh_token: str
    token_type: str = "bearer"
    role: str
    dept: str
    username: str


class RefreshRequest(BaseModel):
    refresh_token: str


class UserCreate(BaseModel):
    username: str
    email: EmailStr
    password: str
    role: str = "viewer"
    departement_code: str = ""


class UserRead(BaseModel):
    id: int
    username: str
    email: str
    role: str
    departement_code: str
    est_actif: bool
    model_config = {"from_attributes": True}


class PasswordChange(BaseModel):
    current_password: str
    new_password: str


# ── Endpoints ─────────────────────────────────────────────────────────────

@router.post("/token", response_model=TokenResponse)
def login(
    form: OAuth2PasswordRequestForm = Depends(),
    db: Session = Depends(get_db),
):
    """Login avec username/password → JWT access + refresh token."""
    user = db.query(Utilisateur).filter_by(username=form.username).first()
    if not user or not verify_password(form.password, user.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Identifiants incorrects",
        )
    if not user.est_actif:
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Compte désactivé")

    user.last_login = datetime.utcnow()
    db.add(AuditLog(utilisateur=user.username, action="login", details=f"user_id={user.id}"))
    db.commit()

    return TokenResponse(
        access_token=create_access_token(user.id, user.username, user.role, user.departement_code),
        refresh_token=create_refresh_token(user.id, user.username),
        role=user.role,
        dept=user.departement_code,
        username=user.username,
    )


@router.post("/refresh", response_model=TokenResponse)
def refresh_token(body: RefreshRequest, db: Session = Depends(get_db)):
    """Renouvelle l'access token depuis un refresh token valide."""
    try:
        payload = decode_token(body.refresh_token)
    except Exception:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Refresh token invalide")

    if payload.get("type") != "refresh":
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Type de token invalide")

    user = db.get(Utilisateur, int(payload["sub"]))
    if not user or not user.est_actif:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Utilisateur introuvable")

    return TokenResponse(
        access_token=create_access_token(user.id, user.username, user.role, user.departement_code),
        refresh_token=create_refresh_token(user.id, user.username),
        role=user.role,
        dept=user.departement_code,
        username=user.username,
    )


@router.get("/me", response_model=UserRead)
def me(
    current_user: dict = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """Retourne le profil de l'utilisateur connecté."""
    user = db.get(Utilisateur, current_user["id"])
    if not user:
        raise HTTPException(status_code=404, detail="Utilisateur introuvable")
    return user


@router.post("/change-password")
def change_password(
    body: PasswordChange,
    current_user: dict = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """Change le mot de passe de l'utilisateur connecté."""
    user = db.get(Utilisateur, current_user["id"])
    if not verify_password(body.current_password, user.hashed_password):
        raise HTTPException(status_code=400, detail="Mot de passe actuel incorrect")
    if len(body.new_password) < 8:
        raise HTTPException(status_code=400, detail="Le mot de passe doit avoir au moins 8 caractères")
    user.hashed_password = hash_password(body.new_password)
    db.add(AuditLog(utilisateur=user.username, action="password_change"))
    db.commit()
    return {"message": "Mot de passe modifié avec succès"}


@router.post("/users", response_model=UserRead, status_code=201)
def create_user(
    body: UserCreate,
    current_user: dict = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """Créer un utilisateur (admin seulement)."""
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Réservé aux administrateurs")
    if db.query(Utilisateur).filter_by(username=body.username).first():
        raise HTTPException(status_code=409, detail="Username déjà utilisé")
    user = Utilisateur(
        username=body.username,
        email=body.email,
        hashed_password=hash_password(body.password),
        role=body.role,
        departement_code=body.departement_code.upper(),
    )
    db.add(user)
    db.add(AuditLog(
        utilisateur=current_user["username"],
        action="create",
        table_cible="utilisateurs",
        details=f"créé: {body.username} [{body.role}]",
    ))
    db.commit()
    db.refresh(user)
    return user


@router.get("/users", response_model=list[UserRead])
def list_users(
    current_user: dict = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """Liste tous les utilisateurs (admin seulement)."""
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Réservé aux administrateurs")
    return db.query(Utilisateur).all()
