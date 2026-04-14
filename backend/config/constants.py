import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent.parent
BACKEND_DIR = BASE_DIR / "backend"
ASSETS_DIR = BACKEND_DIR / "assets"
OUTPUT_DIR = BASE_DIR / "output"
LOGO_PATH = ASSETS_DIR / "logo_melko.jpg"

REQUIRED_COLUMNS = [
    "Adresse d'imposition",
    "x",  # Adresse des travaux
    "Commune",
    "Code Postal",
    "Installateur",
    "N° Programme",
    "N°OPERATION",
    "Numéro de facture",
    "Date de facture",
    "TFPB",
    "Date de paiement",
    "Classification des travaux",
    "Nature des travaux",
    "Détail des travaux retenus",
    "Montant HT facture ",
    "Taux de TVA facture",
    "Retenue de garantie ?",
    "Taux de retenue",
    "TVA facture",
    "Montant TTC facture",
    "Montant de virement TTC",
    "Montant de virement HT",
    "Montant subventions encaisses",
    "Montant des travaux éligibles retenus H.T",
    "Montant dégrèvement demandé",
    "CDIF/ SIP",
    "Adresse CDIF/SIP",
    "N° d'avis",
    "N° Fiscal",
]

# ── Mapping Classification -> Type de travaux ──
# TEE = Travaux d'Economie d'Energie
# PMR = Travaux Personnes a Mobilite Reduite / Accessibilite

# TEE direct = UNIQUEMENT les economies d'energie directes
TEE_CLASSIFICATIONS = [
    "Économies énergie et de fluides",
    "Économies d'énergie et de fluides",
]

# INDIS = Travaux indissociablement lies (etudes, travaux techniques, induits)
# Classifies comme INDIS_TEE ou INDIS_PMR selon le programme dominant
INDIS_CLASSIFICATIONS = [
    "PRESTATIONS D'ETUDES PREALABLES  DE SUIVI OU D'EXPERTISE",
    "TRAVAUX DE CHAUFFAGE",
    "TRAVAUX DE VENTILATION & CHAUFFAGE",
    "TRAVAUX D'ELECTRICITE",
    "TRAVAUX DE COUVERTURE  ET  ETANCHEITE",
    "TRAVAUX D'ISOLATION",
    "TRAVAUX DE MENUISERIE",
    "Travaux induits figurant aux\xa0§ 50 à 90 du III du BOI-TVA-LIQ-30-20-95",
    "Travaux de désinstallation ou de déconstruction des matériaux ou équipements devant être rénovés",
    "Travaux de rénovation des systèmes de répartition des frais d'eau ou de chauffage",
    "fourniture d'équipement nécessaire à la rénovation",
    "Réhabilitation lourde",
]

PMR_CLASSIFICATIONS = [
    "Accessibilité ou adaptation handicap et personnes âgées",
]

EXCLUDED_CLASSIFICATIONS = [
    "Non éligibles",
]

# Classifications qui peuvent etre TEE ou PMR selon contexte
AMBIGUOUS_CLASSIFICATIONS = [
    "Mise aux normes minimales de confort et habitabilité",
    "Mise aux normes minimales de confort et d'habitabilité",
    "Protection des locataires sécurité et mises en sécurité",
    "Travaux imposés administrativement",
    "Protection de la population contre les risques sanitaires amiante plomb eau",
]

# ── Infos Melko Energie (expediteur) ──
MELKO_INFO = {
    "nom": "MELKO ENERGIE",
    "mandant": "la SIP",
    "mandant_complet": "la Société Immobilière Picarde (SIP)",
    "mandant_hlm": "la Société Immobilière Picarde d'HLM",
    "adresse_ligne1": "10 RUE DE PENTHIEVRE",
    "adresse_ligne2": "75008 PARIS",
    "siren": "977 606 508",
    "siret": "977 606 508 00017",
    "tva_intra": "FR92977606508",
    "capital": "5 000",
    "rcs": "RCS DE PARIS",
    "forme": "SAS",
    "signataire_nom": "Amaury Mongongu",
    "signataire_titre": "Directeur Général",
    "signataire_email": "amongongu@melko-energie.com",
    "adresse_mandant_ligne1": "13 PLACE D d'AGUESSEAU",
    "adresse_mandant_ville": "80090 AMIENS",
}

# ── TVA rates for summary ──
TVA_RATES = [0.055, 0.10, 0.20]
TVA_LABELS = {0.055: "5,5%", 0.10: "10%", 0.20: "20%"}

# Article de loi
ARTICLE_CGI = "1391E"
