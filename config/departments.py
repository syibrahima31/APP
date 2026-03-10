"""Department profile registry for the dashboard app."""

from __future__ import annotations

from copy import deepcopy
from typing import Dict


BASE_PROFILE: Dict[str, object] = {
    "institution": "Institut Supérieur Informatique",
    "page_icon": "📊",
    "header_subtitle": (
        "Tableau de bord de pilotage mensuel — "
        "Suivi des enseignements par classe & par matière"
    ),
    "logo_path": "assets/logo_iaid.jpg",
    "plotly_colorway": ["#0B3D91", "#1F6FEB", "#5AA2FF", "#8EC5FF", "#BBDFFF"],
    "author_role": "Chef de Département",
    "assistant_label": "Assistante",
    "assistant_role": "Support administratif",
}


DEPARTMENT_PROFILES: Dict[str, Dict[str, object]] = {
    "IAID": {
        "dept_code": "IAID",
        "department_long": "Département IA & Ingénierie des Données (IAID)",
        "page_title": "IAID — Suivi des classes (Dashboard)",
        "header_title": "Département IA & Ingénierie des Données (IAID)",
        "author_name": "Ibrahima SY",
        "author_email": "ibsy@groupeisi.com",
        "assistant_name": "Dieynaba Barry",
        "assistant_email": "dbarry1@groupeisi.com",
        "secrets": {
            "excel_url": "IAID_EXCEL_URL",
            "dg_emails": "DG_EMAILS",
            "dashboard_url": "DASHBOARD_URL",
            "admin_pin": "ADMIN_PIN",
        },
        "email_prefix": "IAID",
    },
    "KM": {
        "dept_code": "KM",
        "department_long": "Directions des Etudes",
        "page_title": "DE KM — Suivi des classes (Dashboard)",
        "header_title": "Directions des Etudes",
        "author_name": "Mouhamed Gueye",
        "author_email": "mgueye@groupeisi.com",
        "assistant_name": "Fallou Seck",
        "assistant_email": "fseck@groupeisi.com",
        "author_role": "Directeur des Études",
        "assistant_label": "Chef de Département",
        "assistant_role": "Support administratif",
        "secrets": {
            "excel_url": "KM_EXCEL_URL",
            "dg_emails": "DG_EMAILS",
            "dashboard_url": "DASHBOARD_URL",
            "admin_pin": "ADMIN_PIN",
        },
        "email_prefix": "ISI KM",
    },
    "DRS": {
        "dept_code": "DRS",
        "department_long": "Département Réseaux et Systèmes (DRS)",
        "page_title": "DRS — Suivi des classes (Dashboard)",
        "header_title": "Département Réseaux et Systèmes(DRS)",
        "author_name": "Latyr Ndiaye",
        "author_email": "landiaye@groupeisi.com",
        "assistant_name": "Ndéye Ramatoulaye Diop",
        "assistant_email": "nrdiop@groupeisi.com",
        "secrets": {
            "excel_url": "DRS_EXCEL_URL",
            "dg_emails": "DG_EMAILS",
            "dashboard_url": "DASHBOARD_URL",
            "admin_pin": "ADMIN_PIN",
        },
        "email_prefix": "DRS",
    },
    "GI": {
        "dept_code": "GI",
        "department_long": "Departement Genie Informatique (GI)",
        "page_title": "GI - Suivi des classes (Dashboard)",
        "header_title": "Departement Genie Informatique (GI)",
        "author_name": "El Hadji Mor Diaw",
        "author_email": "emdiaw@groupeisi.com",
        "assistant_name": "Mame Anta Thiam",
        "assistant_email": "",
        "secrets": {
            "excel_url": "GI_EXCEL_URL",
            "dg_emails": "DG_EMAILS",
            "dashboard_url": "DASHBOARD_URL",
            "admin_pin": "ADMIN_PIN",
        },
        "email_prefix": "GI",
    },
}


def get_department_config(profile: str | None) -> Dict[str, object]:
    """Return merged config for the requested profile code."""
    key = (profile or "IAID").upper()
    selected = DEPARTMENT_PROFILES.get(key, DEPARTMENT_PROFILES["IAID"])
    merged = deepcopy(BASE_PROFILE)
    merged.update(deepcopy(selected))
    return merged
