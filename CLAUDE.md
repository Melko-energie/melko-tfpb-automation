# CLAUDE.md - Melko Energie TFPB Automation

## Projet
Automatisation de la generation des dossiers de degrevement TFPB (Taxe Fonciere sur les Proprietes Baties - Article 1391E du CGI) pour Melko Energie, mandataire de la SIP (Societe Immobiliere Picarde).

## Stack technique
- **Backend** : Python 3.10-3.13, Flask, pandas, openpyxl, ReportLab
- **Frontend** : HTML, Tailwind CSS (CDN), JavaScript vanilla
- **Pas de framework frontend, pas de bundler, pas de TypeScript**

## Commandes
```bash
# Activer le venv
source .venv/Scripts/activate   # Git Bash Windows
.venv\Scripts\activate          # CMD Windows

# Lancer le serveur
python -m backend.main
# -> http://localhost:8000

# Lancer la validation standalone
python -m backend.scripts.validate chemin/vers/fichier.xlsx
```

## Architecture
```
backend/
  main.py                        # Serveur Flask, 11 endpoints API
  config/constants.py            # Classifications TEE/PMR/INDIS, infos Melko, colonnes requises
  core/reader.py                 # Lecture Excel, normalisation communes, nettoyage
  core/segmenter.py              # Classification TEE/PMR/INDIS, programme dominant
  core/processor.py              # Pipeline principal (orchestrateur)
  generators/annexe_generator.py # Excel annexe (openpyxl)
  generators/courrier_generator.py # PDF courrier (ReportLab)
  generators/tcd_generator.py    # Excel tableau croise dynamique
  generators/recap_generator.py  # Excel recap synthese global
  generators/verification_generator.py # Excel rapport de verification
  scripts/validate.py            # Validation approfondie (8 categories)
  utils/logger.py                # Logging console + fichier + buffer API
frontend/
  index.html                     # Page principale (drag & drop, traitement, telechargement)
  validation.html                # Page audit par commune
output/                          # Dossiers generes : output/{Commune}/{Type}/
```

## Regles strictes
- **NE PAS toucher aux calculs** : toute la logique de calcul (segmenter.py, build_synthesis, generators) est validee et ne doit pas etre modifiee sans demande explicite
- **NE PAS toucher a la logique UX** : le frontend fonctionne correctement
- **NE PAS toucher aux chemins de sortie** : la structure output/{Commune}/{Type}/ est correcte
- **Modifier uniquement ce qui est demande** : pas de refactoring, pas d'ameliorations non sollicitees, pas de changement de fonctionnalite existante
- **Pas d'accents dans les noms de fichiers generes** : les noms de communes dans les chemins output/ sont normalises (accents retires)

## Logique metier cle
- **3 types de travaux** : TEE (economies d'energie), PMR (accessibilite), INDIS (travaux lies)
- **INDIS** suit le programme dominant (INDIS_TEE ou INDIS_PMR)
- **Classifications ambigues** : resolues par mots-cles PMR puis programme dominant puis defaut TEE
- **Normalisation communes** : NFD decomposition pour retirer les accents + remplacement tirets/apostrophes/underscores
- **Annees** : paiements 2023, imposition TFPB 2024 (campagne en cours)
- **Degrevement** = 25% de (montant eligible HT - subventions)

## Conventions de code
- Logging via `backend.utils.logger.get_logger()`
- Pas de docstrings a ajouter sauf si le fichier en a deja
- Montants formates en `#,##0.00 EUR` dans les Excel, `X XXX,XX EUR` dans les PDF
- Les fichiers temporaires sont nettoyes avec `shutil.rmtree` dans les blocs `finally`
