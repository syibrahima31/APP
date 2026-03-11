"""
Client HTTP — Dashboard ISI v3
Gestion JWT, retry, timeout, cache session.
"""
from __future__ import annotations

import os
from typing import Any, Dict, List, Optional

import requests
import streamlit as st

API_BASE = os.environ.get("API_BASE_URL", "http://localhost:8000")
_TIMEOUT = 15


def _headers() -> dict:
    """Retourne les headers avec JWT si connecté."""
    token = st.session_state.get("access_token", "")
    h = {"Content-Type": "application/json", "Accept": "application/json"}
    if token:
        h["Authorization"] = f"Bearer {token}"
    return h


def _url(path: str) -> str:
    return f"{API_BASE}{path}"


def _get(path: str, params: dict = None) -> Any:
    r = requests.get(_url(path), params=params, headers=_headers(), timeout=_TIMEOUT)
    r.raise_for_status()
    return r.json()


def _post(path: str, data: dict = None, files=None, data_form=None) -> Any:
    if files:
        h = {k: v for k, v in _headers().items() if k != "Content-Type"}
        r = requests.post(_url(path), files=files, data=data_form, headers=h, timeout=60)
    else:
        r = requests.post(_url(path), json=data, headers=_headers(), timeout=_TIMEOUT)
    r.raise_for_status()
    return r.json()


def _put(path: str, data: dict) -> Any:
    r = requests.put(_url(path), json=data, headers=_headers(), timeout=_TIMEOUT)
    r.raise_for_status()
    return r.json()


def _delete(path: str) -> None:
    r = requests.delete(_url(path), headers=_headers(), timeout=_TIMEOUT)
    r.raise_for_status()


# ── Authentification ──────────────────────────────────────────────────────

def login(username: str, password: str) -> dict:
    """Login → retourne access_token, refresh_token, role, dept."""
    r = requests.post(
        _url("/api/auth/token"),
        data={"username": username, "password": password},
        headers={"Accept": "application/json"},
        timeout=_TIMEOUT,
    )
    r.raise_for_status()
    return r.json()


def get_me() -> dict:
    return _get("/api/auth/me")


def change_password(current: str, new: str) -> None:
    _post("/api/auth/change-password", {"current_password": current, "new_password": new})


def create_user(username: str, email: str, password: str, role: str, dept: str) -> dict:
    return _post("/api/auth/users", {
        "username": username, "email": email,
        "password": password, "role": role, "departement_code": dept,
    })


def list_users() -> List[dict]:
    return _get("/api/auth/users")


# ── Santé ─────────────────────────────────────────────────────────────────

def health() -> dict:
    r = requests.get(_url("/api/health"), timeout=5)
    r.raise_for_status()
    return r.json()


# ── Classes ───────────────────────────────────────────────────────────────

def get_classes(dept: str) -> List[dict]:
    return _get("/api/classes/", {"dept": dept})


def create_classe(dept: str, nom: str) -> dict:
    return _post("/api/classes/", {"departement_code": dept, "nom": nom})


def delete_classe(classe_id: int) -> None:
    _delete(f"/api/classes/{classe_id}")


# ── Enseignements ─────────────────────────────────────────────────────────

def get_enseignements(dept: str, classe_id: Optional[int] = None) -> List[dict]:
    params = {"dept": dept}
    if classe_id:
        params["classe_id"] = classe_id
    return _get("/api/enseignements/", params)


def create_enseignement(data: dict) -> dict:
    return _post("/api/enseignements/", data)


def update_enseignement(ens_id: int, data: dict) -> dict:
    return _put(f"/api/enseignements/{ens_id}", data)


def delete_enseignement(ens_id: int) -> None:
    _delete(f"/api/enseignements/{ens_id}")


# ── Pointage ──────────────────────────────────────────────────────────────

def get_pointages(enseignement_id: int) -> List[dict]:
    return _get("/api/pointages/", {"enseignement_id": enseignement_id})


def create_pointage(enseignement_id: int, date: str, heures: float,
                    note: str = "", saisi_par: str = "") -> dict:
    return _post("/api/pointages/", {
        "enseignement_id": enseignement_id,
        "date": date, "heures": heures,
        "note": note, "saisi_par": saisi_par,
    })


def delete_pointage(pointage_id: int) -> None:
    _delete(f"/api/pointages/{pointage_id}")


def get_resume(dept: str) -> List[dict]:
    return _get("/api/pointages/resume", {"dept": dept})


# ── Analytics ─────────────────────────────────────────────────────────────

def get_dashboard(dept: str) -> dict:
    """Résumé complet du département depuis le nouveau endpoint analytics."""
    return _get("/api/analytics/dashboard", {"dept": dept})


def get_classes_summary(dept: str) -> List[dict]:
    return _get("/api/analytics/classes", {"dept": dept})


def get_enseignants_summary(dept: str) -> List[dict]:
    return _get("/api/analytics/enseignants", {"dept": dept})


def get_mensuel(dept: str) -> dict:
    return _get("/api/analytics/mensuel", {"dept": dept})


def simulate_charge(dept: str, enseignant: str, nouvelles: int, vhp_moyen: float) -> dict:
    return _get("/api/analytics/simulation-charge", {
        "dept": dept, "enseignant": enseignant,
        "nouvelles_matieres": nouvelles, "vhp_moyen": vhp_moyen,
    })


# ── Alertes ───────────────────────────────────────────────────────────────

def get_alertes(dept: str, niveau: str = None) -> List[dict]:
    params = {"dept": dept}
    if niveau:
        params["niveau"] = niveau
    return _get("/api/alertes/", params)


def get_alertes_stats(dept: str) -> dict:
    return _get("/api/alertes/stats", {"dept": dept})


# ── Import ────────────────────────────────────────────────────────────────

def import_excel_bytes(file_bytes: bytes, filename: str, dept: str,
                       annee: str = None, clear: bool = False) -> dict:
    files = {"file": (filename, file_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    data_form = {"dept": dept, "clear_existing": str(clear).lower()}
    if annee:
        data_form["annee"] = annee
    return _post("/api/import/excel", files=files, data_form=data_form)


def import_excel_url(url: str, dept: str, annee: str = None, clear: bool = False) -> dict:
    data_form = {"url": url, "dept": dept, "clear_existing": str(clear).lower()}
    if annee:
        data_form["annee"] = annee
    return _post("/api/import/excel-url", files=None, data_form=data_form)


# ── Planning / Séances ────────────────────────────────────────────────────

def get_seances(dept: str, classe_id: int = None, enseignement_id: int = None,
                date_debut: str = None, date_fin: str = None, statut: str = None) -> list:
    params = {"dept": dept}
    if classe_id: params["classe_id"] = classe_id
    if enseignement_id: params["enseignement_id"] = enseignement_id
    if date_debut: params["date_debut"] = date_debut
    if date_fin: params["date_fin"] = date_fin
    if statut: params["statut"] = statut
    return _get("/api/seances/", params)


def get_seances_semaine(dept: str, lundi: str = None) -> list:
    params = {"dept": dept}
    if lundi: params["lundi"] = lundi
    return _get("/api/seances/semaine", params)


def create_seance(payload: dict) -> dict:
    return _post("/api/seances/", payload)


def pointer_seance(seance_id: int, saisi_par: str = "", note: str = "",
                   nb_heures: float = None) -> dict:
    data = {"saisi_par": saisi_par, "note": note}
    if nb_heures is not None:
        data["nb_heures"] = nb_heures
    return _post(f"/api/seances/{seance_id}/pointer", data)


def annuler_seance(seance_id: int) -> dict:
    return _post(f"/api/seances/{seance_id}/annuler", {})


def delete_seance(seance_id: int) -> None:
    _delete(f"/api/seances/{seance_id}")
