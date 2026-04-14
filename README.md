# Melko Energie - Automatisation TFPB

Automatisation complète de la génération des dossiers de dégrèvement TFPB (Taxe Foncière sur les Propriétés Bâties - Article 1391E du CGI) à partir d'un fichier Excel de synthèse.

Le système transforme un fichier Excel source en un ensemble structuré de dossiers contenant courriers de dégrèvement (PDF), annexes (Excel) et tableaux croisés dynamiques, organisés par commune et type de travaux.

## Fonctionnalités

### Traitement
- **Segmentation automatique** par commune et 3 types de travaux :
  - **TEE** : Travaux directs d'Économie d'Énergie (économies d'énergie et de fluides)
  - **PMR** : Travaux d'Accessibilité / Adaptation handicap et personnes âgées
  - **INDIS** : Travaux indissociablement liés (études, diagnostics, maîtrise d'oeuvre, travaux induits, déconstruction) - classés INDIS_TEE ou INDIS_PMR selon le programme dominant
- **Normalisation des communes** : fusion automatique des variantes (casse, tirets, accents, apostrophes)
- **Détection du programme dominant** : les travaux liés suivent le type dominant de leur N° Programme

### Génération de fichiers (par commune/type)
- **Courrier PDF** : réclamation contentieuse fidèle au modèle Melko Energie (logo, mise en page, références juridiques)
- **Annexe Excel** : tableau récapitulatif des factures + feuille de synthèse
- **TCD Excel** : tableau croisé dynamique (identification, données fiscales, montants, détail par programme/TVA/installateur)

### Rapports globaux
- **Recap synthèse** (`recap_synthese.xlsx`) : totaux globaux, détail par commune avec colonnes de comparaison (écart, statut OK/A vérifier), détail TVA
- **Rapport de vérification** (`verification_donnees.xlsx`) : anomalies par commune (erreurs, avertissements, suggestions) avec template coloré

### Validation des données
- **8 catégories de contrôles** : structure, communes, montants, dates, classifications, complétude, doublons, cohérence TVA
- **Page web dédiée** (`/validation`) : audit par commune avec cartes visuelles dépliables
- **Mode comparaison** : compare le fichier source avec les sorties du système

### Interface web
- **Design nature** (vert/beige) avec mode sombre
- Drag & drop, animations 3D, transitions fluides
- Console de logs en temps réel
- Téléchargement individuel, groupé (ZIP), recap et rapport de vérification

## Prérequis

- Python 3.10 à 3.13 recommandé (Python 3.14 peut poser des problèmes de venv sur Windows)
- pip

## Installation

### Option 1 : Script automatique (recommandé)

**Windows (double-clic) :**
```
setup.bat
```

**Linux / Mac :**
```bash
chmod +x setup.sh
./setup.sh
```

### Option 2 : Installation manuelle

```bash
cd excel-to-dossier

# Créer le venv
python -m venv .venv

# Si erreur "Unable to copy venvlauncher.exe" (Python 3.14 Windows) :
python -m venv .venv --without-pip
.venv\Scripts\python -m ensurepip --upgrade

# Activer le venv
.venv\Scripts\activate          # Windows CMD
source .venv/Scripts/activate   # Windows Git Bash
source .venv/bin/activate       # Linux / Mac

# Installer les dépendances
pip install -r requirements.txt
```

## Lancement

### Windows (double-clic) :
```
start.bat
```

### Manuellement :
```bash
source .venv/Scripts/activate   # ou .venv\Scripts\activate (CMD)
python -m backend.main
```

Ouvrir **http://localhost:8000** dans le navigateur.

## Utilisation

1. **Valider** (optionnel) : aller sur `/validation`, déposer le fichier Excel, corriger les erreurs signalées
2. **Traiter** : page principale `/`, déposer le fichier Excel, cliquer "Lancer le traitement"
3. **Télécharger** : ZIP complet, recap synthèse, rapport de vérification, ou fichiers individuels par commune

## Structure du projet

```
excel-to-dossier/
├── backend/
│   ├── assets/                         # Logo Melko, modèle PDF
│   ├── config/
│   │   └── constants.py                # Classifications TEE/PMR/INDIS, constantes Melko
│   ├── core/
│   │   ├── reader.py                   # Lecture Excel, validation, normalisation communes
│   │   ├── segmenter.py               # Classification TEE/PMR/INDIS, programme dominant
│   │   └── processor.py               # Pipeline principal
│   ├── generators/
│   │   ├── annexe_generator.py        # Génération annexe .xlsx
│   │   ├── courrier_generator.py      # Génération courrier .pdf
│   │   ├── tcd_generator.py           # Génération tableau croisé dynamique .xlsx
│   │   ├── recap_generator.py         # Génération recap synthèse global
│   │   └── verification_generator.py  # Génération rapport de vérification
│   ├── scripts/
│   │   └── validate.py                # Validation approfondie (8 catégories)
│   ├── utils/
│   │   └── logger.py                  # Logging console + fichier + buffer API
│   └── main.py                        # Serveur Flask + API
├── frontend/
│   ├── assets/                        # Logo
│   ├── index.html                     # Page principale (Tailwind CSS + animations 3D)
│   └── validation.html                # Page audit par commune
├── output/                            # Dossiers générés (gitignored)
├── logs/                              # Fichiers de logs (gitignored)
├── setup.bat                          # Installation auto Windows
├── setup.sh                           # Installation auto Linux/Mac
├── start.bat                          # Lancement rapide Windows
├── .gitignore
├── .env
├── requirements.txt
└── README.md
```

## Structure de sortie

```
output/
├── Amiens/
│   ├── TEE/                    # Travaux directs d'économie d'énergie
│   │   ├── annexe.xlsx
│   │   ├── courrier.pdf
│   │   └── tcd.xlsx
│   ├── PMR/                    # Travaux d'accessibilité
│   │   ├── annexe.xlsx
│   │   ├── courrier.pdf
│   │   └── tcd.xlsx
│   └── INDIS_TEE/              # Travaux indissociablement liés (études, induits...)
│       ├── annexe.xlsx
│       ├── courrier.pdf
│       └── tcd.xlsx
├── Corbie/
│   ├── TEE/
│   ├── PMR/
│   └── INDIS_TEE/
├── recap_synthese.xlsx          # Récapitulatif global
└── verification_donnees.xlsx    # Rapport de vérification
```

## API Endpoints

| Méthode | Route | Description |
|---------|-------|-------------|
| `GET` | `/` | Page principale |
| `GET` | `/validation` | Page audit des données |
| `POST` | `/api/upload` | Upload et traitement complet |
| `POST` | `/api/validate` | Validation simple du fichier |
| `POST` | `/api/validate-communes` | Validation groupée par commune |
| `POST` | `/api/validate-compare` | Validation + comparaison avec sorties système |
| `GET` | `/api/status` | Statut du traitement |
| `GET` | `/api/logs` | Logs en temps réel |
| `GET` | `/api/download` | Télécharger tout en ZIP |
| `GET` | `/api/download-recap` | Télécharger le recap synthèse |
| `GET` | `/api/download-verification` | Télécharger le rapport de vérification |
| `GET` | `/api/download/{commune}/{type}/{filename}` | Télécharger un fichier spécifique |

## Logique de classification

| Classification dans le fichier | Type assigné |
|-------------------------------|-------------|
| Économies d'énergie et de fluides | **TEE** (direct) |
| Accessibilité ou adaptation handicap | **PMR** (direct) |
| TRAVAUX D'ISOLATION, CHAUFFAGE, COUVERTURE, etc. | **INDIS_TEE** ou **INDIS_PMR** (selon programme) |
| PRESTATIONS D'ETUDES PREALABLES | **INDIS_TEE** ou **INDIS_PMR** (selon programme) |
| Travaux induits, désinstallation, rénovation | **INDIS_TEE** ou **INDIS_PMR** (selon programme) |
| Mise aux normes, Protection locataires | **TEE** ou **PMR** (mots-clés puis programme) |
| Non éligibles | **Exclu** |

## Dépannage

| Problème | Solution |
|----------|----------|
| `Unable to copy venvlauncher.exe` | `python -m venv .venv --without-pip` puis `python -m ensurepip --upgrade` |
| `DLL load failed` (pydantic) | Utilisez Python 3.12 ou 3.13 au lieu de 3.14 |
| Port 8000 déjà utilisé | Modifiez le port dans `backend/main.py` |
| Module introuvable | Vérifiez que le venv est activé et que vous êtes dans `excel-to-dossier/` |

## Stack technique

- **Backend** : Python, Flask, pandas, openpyxl, ReportLab
- **Frontend** : HTML, Tailwind CSS (CDN), JavaScript vanilla
- **Fonts** : Montserrat, Merriweather, Source Code Pro
- **Design** : Glassmorphism, dark mode (toggle), animations 3D, particules
