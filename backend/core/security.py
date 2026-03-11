"""
Sécurité — JWT + RBAC + hashing des mots de passe
"""
from __future__ import annotations

from datetime import datetime, timedelta
from typing import Optional

from jose import JWTError, jwt
from passlib.context import CryptContext

from backend.core.config import settings

# ── Hashing mots de passe ─────────────────────────────────────────────────
_pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")


def hash_password(password: str) -> str:
    return _pwd_context.hash(password)


def verify_password(plain: str, hashed: str) -> bool:
    return _pwd_context.verify(plain, hashed)


# ── JWT ───────────────────────────────────────────────────────────────────

def create_access_token(
    user_id: int,
    username: str,
    role: str,
    dept: str,
    expires_delta: Optional[timedelta] = None,
) -> str:
    expire = datetime.utcnow() + (
        expires_delta or timedelta(hours=settings.ACCESS_TOKEN_EXPIRE_HOURS)
    )
    payload = {
        "sub": str(user_id),
        "username": username,
        "role": role,
        "dept": dept,
        "exp": expire,
        "type": "access",
    }
    return jwt.encode(payload, settings.SECRET_KEY, algorithm=settings.ALGORITHM)


def create_refresh_token(user_id: int, username: str) -> str:
    expire = datetime.utcnow() + timedelta(days=settings.REFRESH_TOKEN_EXPIRE_DAYS)
    payload = {
        "sub": str(user_id),
        "username": username,
        "exp": expire,
        "type": "refresh",
    }
    return jwt.encode(payload, settings.SECRET_KEY, algorithm=settings.ALGORITHM)


def decode_token(token: str) -> dict:
    """Décode et valide un JWT. Lève JWTError si invalide."""
    return jwt.decode(token, settings.SECRET_KEY, algorithms=[settings.ALGORITHM])


# ── RBAC — permissions par rôle ───────────────────────────────────────────

ROLE_PERMISSIONS: dict[str, set[str]] = {
    "admin": {"*"},
    "chef_dept": {
        "read:all", "write:pointage", "write:enseignement",
        "read:export", "send:email", "read:audit",
    },
    "viewer": {
        "read:dashboard", "read:classes", "read:enseignements",
        "read:pointages", "read:alertes",
    },
}


def has_permission(role: str, permission: str) -> bool:
    perms = ROLE_PERMISSIONS.get(role, set())
    return "*" in perms or permission in perms


def can_access_dept(token_payload: dict, dept_code: str) -> bool:
    """Vérifie que l'utilisateur peut accéder à ce département."""
    if token_payload.get("role") == "admin":
        return True
    return token_payload.get("dept", "").upper() == dept_code.upper()
