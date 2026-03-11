"""
Dépendances FastAPI réutilisables (injection de dépendances).
"""
from __future__ import annotations

from typing import Optional

from fastapi import Depends, Header, HTTPException, status
from jose import JWTError
from sqlalchemy.orm import Session

from backend.core.security import decode_token, has_permission, can_access_dept
from db.database import get_db
from db.models import Utilisateur


def get_current_user(
    authorization: Optional[str] = Header(None),
    db: Session = Depends(get_db),
) -> dict:
    """
    Extrait et valide le JWT depuis le header Authorization.
    Retourne le payload enrichi avec les infos utilisateur.
    Lève 401 si token absent/invalide.
    """
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token d'authentification manquant",
            headers={"WWW-Authenticate": "Bearer"},
        )
    token = authorization.split(" ", 1)[1]
    try:
        payload = decode_token(token)
    except JWTError:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token invalide ou expiré",
            headers={"WWW-Authenticate": "Bearer"},
        )
    if payload.get("type") != "access":
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Type de token invalide")

    user_id = int(payload["sub"])
    user = db.get(Utilisateur, user_id)
    if not user or not user.est_actif:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Utilisateur inactif ou introuvable")

    return {
        "id": user.id,
        "username": user.username,
        "role": user.role,
        "dept": user.departement_code,
    }


def get_optional_user(
    authorization: Optional[str] = Header(None),
    db: Session = Depends(get_db),
) -> Optional[dict]:
    """Utilisateur optionnel — retourne None si non connecté (mode lecture publique)."""
    if not authorization:
        return None
    try:
        return get_current_user(authorization, db)
    except HTTPException:
        return None


def require_permission(permission: str):
    """Factory de dépendances pour protéger les routes par permission."""
    def dependency(current_user: dict = Depends(get_current_user)) -> dict:
        if not has_permission(current_user["role"], permission):
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=f"Permission '{permission}' requise (rôle actuel: {current_user['role']})",
            )
        return current_user
    return dependency


def require_dept_access(dept_param_name: str = "dept"):
    """Factory qui vérifie que l'utilisateur peut accéder au département demandé."""
    def dependency(current_user: dict = Depends(get_current_user)) -> dict:
        return current_user
    return dependency


# Raccourcis
RequireAdmin   = Depends(require_permission("*"))
RequireChef    = Depends(require_permission("write:enseignement"))
RequireViewer  = Depends(get_current_user)
