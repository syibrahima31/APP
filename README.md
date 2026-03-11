# Dashboard ISI — Suivi des Enseignements

Tableau de bord de pilotage mensuel pour le suivi des enseignements par classe et par matière.
Architecture : **FastAPI (backend)** + **Streamlit (frontend)** + **SQLite/PostgreSQL (base de données)**

---

## Sommaire

1. [Architecture du projet](#1-architecture-du-projet)
2. [Installation & lancement](#2-installation--lancement)
3. [Sources de données](#3-sources-de-données)
4. [Onglets du dashboard](#4-onglets-du-dashboard)
5. [Système de pointage](#5-système-de-pointage)
6. [Administration](#6-administration)
7. [API REST (FastAPI)](#7-api-rest-fastapi)
8. [Notifications email](#8-notifications-email)
9. [Exports](#9-exports)
10. [Intelligence Artificielle](#10-intelligence-artificielle)
11. [Configuration multi-département](#11-configuration-multi-département)
12. [Variables d'environnement & Secrets](#12-variables-denvironnement--secrets)
13. [Migration depuis Excel](#13-migration-depuis-excel)
14. [Structure des fichiers](#14-structure-des-fichiers)

---

## 1. Architecture du projet

```
APP/
├── backend/                   ← API REST FastAPI
│   ├── main.py                ← Point d'entrée FastAPI
│   ├── routers/
│   │   ├── classes.py         ← Endpoints CRUD classes
│   │   ├── enseignements.py   ← Endpoints CRUD enseignements
│   │   └── pointage.py        ← Endpoints système de pointage
│   └── schemas/
│       ├── classe.py          ← Schémas Pydantic classes
│       ├── enseignement.py    ← Schémas Pydantic enseignements
│       └── pointage.py        ← Schémas Pydantic pointage
├── db/
│   ├── database.py            ← Connexion SQLAlchemy (SQLite/PostgreSQL)
│   ├── models.py              ← Modèles ORM (Classe, Enseignement, Pointage)
│   ├── crud.py                ← Lecture DB → DataFrame pandas
│   └── import_excel.py        ← Migration Excel → base de données
├── frontend/
│   ├── app.py                 ← Dashboard Streamlit (appelle FastAPI)
│   └── api_client.py          ← Client HTTP vers le backend
├── app.py                     ← Dashboard Streamlit legacy (Excel ou DB)
├── app_iaid.py                ← Lanceur profil IAID
├── app_km.py                  ← Lanceur profil KM
├── app_gi.py                  ← Lanceur profil GI
├── config/
│   └── departments.py         ← Profils des départements
├── services/
│   └── email_notifications.py ← Envoi d'emails et rappels
├── ui/
│   └── components.py          ← Composants UI réutilisables
├── utils/
│   └── data_pipeline.py       ← Pipeline de données (Excel + DB)
├── assets/
│   └── logo_iaid.jpg          ← Logo pour les exports PDF
├── data/
│   └── dashboard.db           ← Base SQLite (créée automatiquement)
└── requirements.txt
```

---

## 2. Installation & lancement

### Prérequis

```bash
pip install -r requirements.txt
```

### Mode nouveau — FastAPI + Streamlit (recommandé)

**Terminal 1 — Backend API :**
```bash
uvicorn backend.main:app --reload --port 8000
```
Documentation API interactive : `http://localhost:8000/docs`

**Terminal 2 — Frontend Dashboard :**
```bash
streamlit run frontend/app.py
```

### Mode legacy — Streamlit seul

```bash
streamlit run app.py

# Par département :
streamlit run app_iaid.py
streamlit run app_km.py
streamlit run app_gi.py
```

---

## 3. Sources de données

Le dashboard supporte trois modes d'import :

| Mode | Description |
|---|---|
| **Base de données** | Lecture depuis SQLite ou PostgreSQL via l'API FastAPI |
| **URL (auto)** | Téléchargement automatique d'un `.xlsx` depuis une URL (Google Drive, OneDrive…) avec cache ETag/Last-Modified |
| **Upload (manuel)** | Import manuel d'un fichier `.xlsx` directement dans l'interface |

### Format attendu du fichier Excel

- Chaque **feuille** = une **classe** (ex : `L3 IA`, `M1 Data`)
- Colonnes minimales requises : `Matière` + `VHP`
- Colonnes optionnelles : `Responsable`, `Email`, `Semestre`, `Observations`, `Début prévu`, `Fin prévue`, `Oct`, `Nov`, `Déc`, `Jan`, `Fév`, `Mars`, `Avril`, `Mai`, `Juin`, `Juil`, `Août`
- Le pipeline normalise automatiquement **40+ variantes** de noms de colonnes (ex : "Module", "UE", "Enseignant", "Volume Horaire Prévu"…)
- La ligne d'en-tête est détectée automatiquement (lignes de titre ignorées)

### Métriques calculées automatiquement

| Métrique | Calcul |
|---|---|
| **VHR** (heures réalisées) | Somme des colonnes mensuelles Oct → Août (ou somme des pointages en mode DB) |
| **Écart** | VHR − VHP |
| **Taux** | VHR / VHP |
| **Statut_auto** | `Non démarré` si VHR=0 · `En cours` si VHR < VHP · `Terminé` si VHR ≥ VHP |

---

## 4. Onglets du dashboard

### 🌐 Vue globale

Synthèse complète du département :

- **KPI cards animées** : matières, taux moyen, terminées, en cours, retard cumulé (h), VHP total, VHR total, non démarrées
- **Avancement par classe** : tableau avec barre de progression du taux de réalisation
- **Répartition des statuts** : graphique camembert (Terminé / En cours / Non démarré)
- **VHP vs VHR par classe** : graphique barres comparatives
- **Top retards** : 20 matières avec le plus grand retard (Écart le plus négatif)

### 🏫 Par classe

Analyse approfondie et comparaison entre deux classes :

- **Tableau de synthèse** : nb matières, taux moyen, VHP total, VHR total, retard (h), terminées, non démarrées — pour chaque classe
- **Comparaison A vs B** : sélection de deux classes et affichage KPIs côte à côte
- **Top 15 retards** : détails des matières en retard pour chaque classe sélectionnée

### 📖 Par matière

Analyse transversale des matières toutes classes confondues :

- **Tableau agrégé** : par matière — nb classes, VHP, VHR, taux moyen, retard, nb non démarrées
- **Matières en alerte** : filtre automatique selon les seuils (taux < seuil orange ou écart < seuil critique)
- **Graphique taux** : barres colorées rouge→vert selon le niveau de réalisation

### 👤 Par enseignant

Suivi de la charge et des retards par responsable :

- **Synthèse par enseignant** : nb matières, classes, VHP, VHR, taux moyen, retard (h), statuts — trié par retard décroissant
- **Top retards par enseignant** : liste paramétrable (10 → 200 lignes)
- **Non démarrés par enseignant** : graphique barres
- **Charge comparée** : VHP vs VHR et écart total
- Gestion des modules **non affectés** (`⚠️ Non affecté`)

### 📅 Analyse mensuelle

Visualisations temporelles :

- **Courbe mensuelle** : évolution des heures totales réalisées Oct → Août
- **Matrice Classe × Mois** : tableau pivot avec heures par classe et par mois
- **Heatmap** : carte de chaleur (heures/classe/mois) — désactivée si trop de données (> 250 cellules)
- **Classe la plus active** : quelle classe a réalisé le plus d'heures par mois

### 🚨 Alertes

Détection et priorisation des situations critiques :

- **KPI alertes** : total alertes, retards critiques, non démarrés
- **Règles d'alerte** :
  - `🔻 Retard critique` : Écart ≤ seuil critique (défaut −6h, configurable)
  - `⛔ Fin dépassée` : date de fin prévue dépassée et cours non terminé *(mode legacy)*
  - `🛑 Non démarré` : matière avec VHR = 0
- **Priorité** : fin dépassée > retard critique > non démarré
- **Vue priorisée** : tableau trié par priorité puis par écart
- **Graphiques** : camembert types d'alertes + barres alertes par classe
- **Envoi email** *(mode legacy)* : notification aux enseignants concernés

---

## 5. Système de pointage

Fonctionnalité permettant d'enregistrer les heures réalisées **séance par séance**.

### Nouvelle saisie

1. Sélection **classe** → **matière / responsable**
2. Saisie : **date de séance** · **heures réalisées** (0.5h → 10h, pas 0.5h) · **saisi par** · **remarque** optionnelle
3. Enregistrement via l'API FastAPI
4. Récapitulatif affiché (VHP · VHR actuel · taux · statut)

### Historique

- Consultation de tous les pointages d'une matière (ID · date · heures · remarque · saisi par)
- Métriques : total réalisé vs VHP · nb séances
- Suppression d'un pointage individuel

### Calcul du VHR depuis les pointages

Le VHR est la **somme de tous les pointages** enregistrés pour une matière :
- Répartition automatique par mois académique (Oct = mois 10, Août = mois 8)
- Mise à jour du dashboard après chaque saisie

---

## 6. Administration

### Classes

- **Créer** une classe pour un département
- **Lister** les classes (ID, nom)
- **Supprimer** une classe *(suppression en cascade : enseignements + pointages)*

### Enseignements

- **Ajouter** : classe · matière* · VHP* · responsable · email · semestre · début/fin prévus · observations
- **Lister** avec filtrage par classe (VHP · VHR · taux · statut)
- **Supprimer** *(suppression en cascade des pointages)*

---

## 7. API REST (FastAPI)

Documentation interactive : `http://localhost:8000/docs`

### `/api/classes/`

| Méthode | Endpoint | Description |
|---|---|---|
| GET | `/?dept=IAID` | Lister les classes d'un département |
| GET | `/{id}` | Détail d'une classe |
| POST | `/` | Créer une classe |
| DELETE | `/{id}` | Supprimer une classe |

### `/api/enseignements/`

| Méthode | Endpoint | Description |
|---|---|---|
| GET | `/?dept=IAID&classe_id=3` | Lister (avec VHR calculé depuis pointages) |
| GET | `/{id}` | Détail d'un enseignement |
| POST | `/` | Créer un enseignement |
| PUT | `/{id}` | Modifier |
| DELETE | `/{id}` | Supprimer |

### `/api/pointages/`

| Méthode | Endpoint | Description |
|---|---|---|
| GET | `/?enseignement_id=5` | Lister les pointages d'une matière |
| POST | `/` | Enregistrer une séance (date · heures · note · saisi_par) |
| DELETE | `/{id}` | Supprimer un pointage |
| GET | `/resume?dept=IAID` | Résumé complet : VHR · taux · heures/mois pour tout le département |

---

## 8. Notifications email

Disponible dans l'onglet **Alertes** du dashboard legacy (`app.py`).

### Rappels aux enseignants

- Sélection du **lot** : toutes alertes / non démarrés / retards critiques / fin dépassée / en cours / terminés
- **1 email par enseignant** regroupant toutes ses matières concernées
- **Template HTML** professionnel : tableau récapitulatif classe · semestre · matière · VHP · VHR · écart · statut
- Sélection/désélection manuelle des enseignants avant envoi
- Envoi en lot avec barre de progression

### Rappel mensuel DG/DGE

- Envoi automatique **1 fois par mois** au Directeur Général et Directeur des Études
- **Anti-doublon** : verrou JSON (`.streamlit/last_reminder_{DEPT}.json`) + fichier de lock
- Pièces jointes optionnelles : rapport PDF et/ou export Excel

### Fonctions du service email (`services/email_notifications.py`)

| Fonction | Description |
|---|---|
| `send_email_reminder()` | Envoi SMTP avec pièces jointes (PDF, Excel) |
| `build_prof_email_html()` | Génère le HTML de l'email enseignant |
| `get_last_reminder_month()` | Lit le dernier mois de rappel |
| `set_last_reminder_month()` | Enregistre le mois de rappel |
| `lock_is_active()` | Vérifie si un envoi est en cours |
| `set_lock()` / `clear_lock()` | Gestion du verrou anti-doublon |

---

## 9. Exports

### Export Excel consolidé

Fichier `.xlsx` multi-feuilles :
- **Consolidé** : toutes les données filtrées (colonnes mensuelles · VHR · Écart · Taux · Statut)
- **Synthese_Classes** : agrégation par classe
- **Synthese_Responsables** : agrégation par enseignant

### Export PDF — Rapport mensuel

PDF professionnel (ReportLab) :
- En-tête : logo · département · institution · auteur · date
- KPIs globaux colorisés
- Tableau détaillé : matières avec indicateurs (vert/orange/rouge selon seuils)
- Titre personnalisable · logo uploadable dans la sidebar

### Export PDF — Observations

PDF spécifique aux retours pédagogiques :
- Matières avec le champ "Observations" renseigné
- Indicateurs de performance associés

---

## 10. Intelligence Artificielle

Réservé aux **administrateurs** (connexion par PIN).

### Résumé IA des observations (OpenAI GPT-4)

- Analyse les observations pédagogiques du fichier
- Génère un **résumé synthétique en Markdown** (tendances · blocages · recommandations)
- Paramètre : nb max d'observations envoyées (50 → 800)
- Modèle : `gpt-4.1-mini`
- Résultat téléchargeable au format `.md`

---

## 11. Configuration multi-département

Quatre profils dans `config/departments.py` :

| Code | Département | Chef de département |
|---|---|---|
| **IAID** | IA & Ingénierie des Données | Ibrahima SY |
| **KM** | Directions des Études | Mouhamed Gueye |
| **DRS** | Réseaux et Systèmes | Latyr Ndiaye |
| **GI** | Génie Informatique | El Hadji Mor Diaw |

**Sélection du profil** :
```bash
APP_DEPT_PROFILE=KM streamlit run app.py
# ou via les lanceurs dédiés :
streamlit run app_km.py
```

---

## 12. Variables d'environnement & Secrets

### Base de données

| Variable | Défaut | Description |
|---|---|---|
| `DB_URL` | — | URL PostgreSQL (ex: `postgresql://user:pass@host/db`) |
| `DB_PATH` | `data/dashboard.db` | Chemin SQLite local |
| `API_BASE_URL` | `http://localhost:8000` | URL du backend FastAPI |

### Secrets Streamlit (`.streamlit/secrets.toml`)

```toml
IAID_EXCEL_URL = "https://..."
KM_EXCEL_URL   = "https://..."
DRS_EXCEL_URL  = "https://..."
GI_EXCEL_URL   = "https://..."

DG_EMAILS      = "dg@isi.sn,dge@isi.sn"
DASHBOARD_URL  = "https://votre-app.streamlit.app"
ADMIN_PIN      = "1234"

SMTP_HOST      = "smtp.gmail.com"
SMTP_PORT      = 587
SMTP_USER      = "votre@email.com"
SMTP_PASS      = "mot_de_passe_app"
SMTP_FROM      = "votre@email.com"

OPENAI_API_KEY = "sk-..."
```

---

## 13. Migration depuis Excel

```bash
# Import simple
python db/import_excel.py --dept IAID --file data/suivi_iaid.xlsx

# Réimport complet (efface les données existantes)
python db/import_excel.py --dept IAID --file data/suivi_iaid.xlsx --clear

# Plusieurs départements
python db/import_excel.py --dept KM  --file data/suivi_km.xlsx
python db/import_excel.py --dept DRS --file data/suivi_drs.xlsx
python db/import_excel.py --dept GI  --file data/suivi_gi.xlsx
```

Le script :
1. Lit toutes les feuilles du fichier Excel
2. Normalise les noms de colonnes (40+ variantes supportées)
3. Crée les classes (une par feuille)
4. Crée les enseignements
5. Convertit les heures mensuelles en **pointages de migration** (1 pointage par mois)

---

## 14. Structure des fichiers

### `utils/data_pipeline.py`

| Fonction | Description |
|---|---|
| `load_excel_all_sheets()` | Charge un Excel multi-feuilles → DataFrame |
| `load_from_db()` | Charge depuis la DB → DataFrame (même format) |
| `normalize_columns()` | Normalise 40+ variantes de noms de colonnes |
| `compute_metrics()` | Calcule VHR · Écart · Taux · Statut_auto |
| `ensure_month_cols()` | Crée les colonnes mensuelles manquantes (= 0) |
| `fetch_excel_if_changed()` | Télécharge l'Excel seulement si modifié (ETag) |
| `df_to_excel_bytes()` | Sérialise un dict de DataFrames en fichier Excel |
| `make_long()` | Convertit le DataFrame en format long (mois dépivotés) |

### `db/models.py` — Tables SQLAlchemy

| Table | Colonnes clés | Description |
|---|---|---|
| `classes` | id · departement_code · nom | Classes par département |
| `enseignements` | id · classe_id · matiere · vhp · responsable · email · semestre… | Matières planifiées |
| `pointages` | id · enseignement_id · date · heures · note · saisi_par | Séances réalisées |

### `frontend/api_client.py`

| Fonction | Description |
|---|---|
| `health()` | Vérification état API |
| `get_classes(dept)` | Liste des classes |
| `create_classe(dept, nom)` | Créer une classe |
| `delete_classe(id)` | Supprimer une classe |
| `get_enseignements(dept, classe_id)` | Liste avec VHR calculé |
| `create_enseignement(data)` | Créer un enseignement |
| `update_enseignement(id, data)` | Modifier un enseignement |
| `delete_enseignement(id)` | Supprimer un enseignement |
| `get_pointages(enseignement_id)` | Historique des pointages |
| `create_pointage(ens_id, date, heures, …)` | Enregistrer une séance |
| `delete_pointage(id)` | Supprimer un pointage |
| `get_resume(dept)` | Résumé complet (VHR · heures/mois par matière) |

---

*Dashboard ISI — v2.0.0 · Architecture FastAPI + Streamlit + SQLAlchemy*
