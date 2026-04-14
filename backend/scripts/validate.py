"""
Script de validation approfondi des donnees Excel TFPB.
Usage standalone : python -m backend.scripts.validate chemin/vers/fichier.xlsx
Usage API : from backend.scripts.validate import validate_file
"""

import sys
import math
from pathlib import Path
from datetime import datetime

import pandas as pd

from backend.config.constants import (
    TEE_CLASSIFICATIONS,
    INDIS_CLASSIFICATIONS,
    PMR_CLASSIFICATIONS,
    EXCLUDED_CLASSIFICATIONS,
    AMBIGUOUS_CLASSIFICATIONS,
    REQUIRED_COLUMNS,
)


# ── Severites ──
ERROR = "ERREUR"
WARNING = "ATTENTION"
INFO = "INFO"

# ── Categories ──
CAT_STRUCTURE = "Structure"
CAT_COMMUNE = "Communes"
CAT_MONTANT = "Montants"
CAT_DATE = "Dates"
CAT_CLASSIFICATION = "Classifications"
CAT_COMPLETUDE = "Completude"
CAT_DOUBLON = "Doublons"
CAT_TVA = "Coherence TVA"

# Codes postaux attendus par departement Somme
SOMME_CP_PREFIX = "80"

ALL_KNOWN_CLASSIFICATIONS = (
    TEE_CLASSIFICATIONS
    + INDIS_CLASSIFICATIONS
    + PMR_CLASSIFICATIONS
    + EXCLUDED_CLASSIFICATIONS
    + AMBIGUOUS_CLASSIFICATIONS
)


def validate_file(file_path: Path) -> dict:
    """
    Validate an Excel file and return a detailed report.
    Returns {
        "summary": {"ERREUR": int, "ATTENTION": int, "INFO": int, "total_lignes": int},
        "anomalies": [ {ligne, colonne, valeur, attendu, severite, categorie, message} ]
    }
    """
    anomalies = []

    # ── 1. Structure ──
    anomalies.extend(_check_structure(file_path))
    if any(a["severite"] == ERROR and a["categorie"] == CAT_STRUCTURE for a in anomalies):
        return _build_report(anomalies, 0)

    df = pd.read_excel(file_path, sheet_name=0)
    total_lignes = len(df)

    # Rename 'x' for internal use
    if "x" in df.columns:
        df = df.rename(columns={"x": "Adresse des travaux"})

    # ── 2. Communes ──
    anomalies.extend(_check_communes(df))

    # ── 3. Montants ──
    anomalies.extend(_check_montants(df))

    # ── 4. Dates ──
    anomalies.extend(_check_dates(df))

    # ── 5. Classifications ──
    anomalies.extend(_check_classifications(df))

    # ── 6. Completude ──
    anomalies.extend(_check_completude(df))

    # ── 7. Doublons ──
    anomalies.extend(_check_doublons(df))

    # ── 8. Coherence TVA ──
    anomalies.extend(_check_tva_coherence(df))

    return _build_report(anomalies, total_lignes)


def _build_report(anomalies: list, total_lignes: int) -> dict:
    summary = {ERROR: 0, WARNING: 0, INFO: 0, "total_lignes": total_lignes}
    for a in anomalies:
        sev = a["severite"]
        if sev in summary:
            summary[sev] += 1
    summary["total_anomalies"] = summary[ERROR] + summary[WARNING] + summary[INFO]
    return {"summary": summary, "anomalies": anomalies}


def _a(ligne, colonne, valeur, attendu, severite, categorie, message):
    """Helper to create an anomaly dict."""
    return {
        "ligne": ligne,
        "colonne": colonne,
        "valeur": str(valeur)[:100] if valeur is not None else "",
        "attendu": str(attendu)[:100] if attendu is not None else "",
        "severite": severite,
        "categorie": categorie,
        "message": message,
    }


# ═══════════════════════════════════════════════════════════
# 1. STRUCTURE
# ═══════════════════════════════════════════════════════════

def _check_structure(file_path: Path) -> list:
    results = []

    if not file_path.exists():
        results.append(_a(0, "-", str(file_path), "Fichier existant", ERROR, CAT_STRUCTURE, "Fichier introuvable"))
        return results

    if file_path.suffix.lower() not in (".xlsx", ".xls"):
        results.append(_a(0, "-", file_path.suffix, ".xlsx ou .xls", ERROR, CAT_STRUCTURE, "Format de fichier non supporte"))
        return results

    try:
        df = pd.read_excel(file_path, sheet_name=0)
    except Exception as e:
        results.append(_a(0, "-", str(e)[:80], "Fichier Excel lisible", ERROR, CAT_STRUCTURE, "Impossible de lire le fichier Excel"))
        return results

    if len(df) == 0:
        results.append(_a(0, "-", "0 lignes", "> 0 lignes", ERROR, CAT_STRUCTURE, "Le fichier est vide (aucune ligne de donnees)"))
        return results

    results.append(_a(0, "-", f"{len(df)} lignes", "-", INFO, CAT_STRUCTURE, f"Fichier lu avec succes : {len(df)} lignes, {len(df.columns)} colonnes"))

    # Check required columns
    # The source has 'x' instead of 'Adresse des travaux', that's expected
    check_cols = [c for c in REQUIRED_COLUMNS]
    for col in check_cols:
        if col not in df.columns:
            results.append(_a(0, col, "Absente", "Presente", ERROR, CAT_STRUCTURE, f"Colonne obligatoire manquante : '{col}'"))

    # Check for completely empty columns
    for col in df.columns:
        if df[col].isna().all():
            results.append(_a(0, col, "Toutes vides", "Au moins 1 valeur", WARNING, CAT_STRUCTURE, f"La colonne '{col}' est entierement vide"))

    return results


# ═══════════════════════════════════════════════════════════
# 2. COMMUNES
# ═══════════════════════════════════════════════════════════

def _check_communes(df: pd.DataFrame) -> list:
    results = []

    if "Commune" not in df.columns:
        return results

    # Communes vides
    empty_mask = df["Commune"].isna() | (df["Commune"].astype(str).str.strip().isin(["", "nan"]))
    for idx in df[empty_mask].index:
        results.append(_a(idx + 2, "Commune", "", "Nom de commune", ERROR, CAT_COMMUNE, "Commune manquante"))

    # Doublons de casse
    communes = df["Commune"].dropna().astype(str).str.strip()
    communes_clean = communes[~communes.isin(["", "nan"])]
    upper_map = {}
    for c in communes_clean.unique():
        key = c.upper()
        if key not in upper_map:
            upper_map[key] = []
        upper_map[key].append(c)

    for key, variants in upper_map.items():
        if len(variants) > 1:
            results.append(_a(
                0, "Commune", " / ".join(variants), "Nom unique",
                WARNING, CAT_COMMUNE,
                f"Meme commune ecrite differemment : {', '.join(variants)} ({len(variants)} variantes)"
            ))

    # Espaces en debut/fin
    for idx, val in communes.items():
        if val != val.strip():
            results.append(_a(idx + 2, "Commune", repr(val), val.strip(), INFO, CAT_COMMUNE, "Espaces superflus dans le nom de commune"))

    # Code postal coherent avec commune
    if "Code Postal" in df.columns:
        for idx, row in df.iterrows():
            commune = str(row.get("Commune", "")).strip()
            cp = str(row.get("Code Postal", "")).strip()
            if commune in ("", "nan") or cp in ("", "nan"):
                continue
            # Verifier que le CP est un nombre de 5 chiffres
            cp_clean = cp.replace(".0", "").replace(" ", "")
            if not cp_clean.isdigit() or len(cp_clean) != 5:
                results.append(_a(idx + 2, "Code Postal", cp, "5 chiffres", WARNING, CAT_COMMUNE, f"Code postal invalide : '{cp}' pour {commune}"))
            elif not cp_clean.startswith(SOMME_CP_PREFIX) and not cp_clean.startswith("60"):
                results.append(_a(idx + 2, "Code Postal", cp_clean, f"Commence par {SOMME_CP_PREFIX}", INFO, CAT_COMMUNE, f"Code postal hors Somme/Oise : {cp_clean} pour {commune}"))

    return results


# ═══════════════════════════════════════════════════════════
# 3. MONTANTS
# ═══════════════════════════════════════════════════════════

def _check_montants(df: pd.DataFrame) -> list:
    results = []

    money_cols = [
        "Montant HT facture ", "TVA facture", "Montant TTC facture",
        "Montant de virement TTC", "Montant de virement HT",
        "Montant des travaux éligibles retenus H.T", "Montant dégrèvement demandé",
    ]

    for col in money_cols:
        if col not in df.columns:
            continue
        series = pd.to_numeric(df[col], errors="coerce")

        # Montants negatifs
        neg_mask = series < 0
        for idx in df[neg_mask].index:
            results.append(_a(idx + 2, col, series[idx], ">= 0", ERROR, CAT_MONTANT, f"Montant negatif : {series[idx]}"))

    # HT + TVA = TTC (tolerance 0.50 EUR)
    if all(c in df.columns for c in ["Montant HT facture ", "TVA facture", "Montant TTC facture"]):
        ht = pd.to_numeric(df["Montant HT facture "], errors="coerce").fillna(0)
        tva = pd.to_numeric(df["TVA facture"], errors="coerce").fillna(0)
        ttc = pd.to_numeric(df["Montant TTC facture"], errors="coerce").fillna(0)

        for idx in df.index:
            if ht[idx] == 0 and tva[idx] == 0 and ttc[idx] == 0:
                continue
            calcul = ht[idx] + tva[idx]
            ecart = abs(calcul - ttc[idx])
            if ecart > 0.50 and ttc[idx] != 0:
                results.append(_a(
                    idx + 2, "Montant TTC facture",
                    f"HT({ht[idx]:.2f}) + TVA({tva[idx]:.2f}) = {calcul:.2f}",
                    f"TTC = {ttc[idx]:.2f}",
                    WARNING, CAT_MONTANT,
                    f"Ecart HT+TVA vs TTC de {ecart:.2f} EUR"
                ))

    # Degrevement = 25% de eligible (tolerance 1 EUR)
    if all(c in df.columns for c in ["Montant des travaux éligibles retenus H.T", "Montant dégrèvement demandé"]):
        eligible = pd.to_numeric(df["Montant des travaux éligibles retenus H.T"], errors="coerce").fillna(0)
        degrev = pd.to_numeric(df["Montant dégrèvement demandé"], errors="coerce").fillna(0)

        for idx in df.index:
            if eligible[idx] == 0 and degrev[idx] == 0:
                continue
            attendu = eligible[idx] * 0.25
            ecart = abs(degrev[idx] - attendu)
            if ecart > 1.0 and eligible[idx] > 0:
                results.append(_a(
                    idx + 2, "Montant dégrèvement demandé",
                    f"{degrev[idx]:.2f}",
                    f"25% de {eligible[idx]:.2f} = {attendu:.2f}",
                    WARNING, CAT_MONTANT,
                    f"Degrevement ({degrev[idx]:.2f}) != 25% du montant eligible ({attendu:.2f}), ecart {ecart:.2f} EUR"
                ))

    # Ligne eligible mais montant 0
    if all(c in df.columns for c in ["Classification des travaux", "Montant des travaux éligibles retenus H.T"]):
        eligible = pd.to_numeric(df["Montant des travaux éligibles retenus H.T"], errors="coerce").fillna(0)
        for idx, row in df.iterrows():
            classif = str(row.get("Classification des travaux", "")).strip()
            if classif in EXCLUDED_CLASSIFICATIONS or classif in ("", "nan"):
                continue
            if eligible[idx] == 0:
                results.append(_a(
                    idx + 2, "Montant des travaux éligibles retenus H.T",
                    "0", "> 0 pour ligne eligible",
                    WARNING, CAT_MONTANT,
                    f"Ligne classee '{classif}' mais montant eligible = 0"
                ))

    # Virement HT vs HT facture
    if all(c in df.columns for c in ["Montant HT facture ", "Montant de virement HT"]):
        ht_fact = pd.to_numeric(df["Montant HT facture "], errors="coerce").fillna(0)
        ht_vir = pd.to_numeric(df["Montant de virement HT"], errors="coerce").fillna(0)
        for idx in df.index:
            if ht_fact[idx] == 0 and ht_vir[idx] == 0:
                continue
            if ht_fact[idx] > 0 and ht_vir[idx] > 0:
                ecart = abs(ht_fact[idx] - ht_vir[idx])
                if ecart > 0.50:
                    results.append(_a(
                        idx + 2, "Montant de virement HT",
                        f"Virement: {ht_vir[idx]:.2f}",
                        f"Facture HT: {ht_fact[idx]:.2f}",
                        INFO, CAT_MONTANT,
                        f"Ecart entre montant facture HT et virement HT de {ecart:.2f} EUR (retenue de garantie ?)"
                    ))

    return results


# ═══════════════════════════════════════════════════════════
# 4. DATES
# ═══════════════════════════════════════════════════════════

def _check_dates(df: pd.DataFrame) -> list:
    results = []

    date_cols = ["Date de facture", "Date de paiement"]
    for col in date_cols:
        if col not in df.columns:
            continue
        for idx, val in df[col].items():
            if pd.isna(val):
                continue
            try:
                if isinstance(val, str):
                    pd.to_datetime(val, dayfirst=True)
                # else it's already a datetime
            except Exception:
                results.append(_a(idx + 2, col, str(val)[:30], "Date valide", ERROR, CAT_DATE, f"Format de date invalide : '{val}'"))

    # Date paiement >= date facture
    if all(c in df.columns for c in ["Date de facture", "Date de paiement"]):
        for idx, row in df.iterrows():
            try:
                d_fact = pd.to_datetime(row["Date de facture"], dayfirst=True)
                d_paie = pd.to_datetime(row["Date de paiement"], dayfirst=True)
                if pd.notna(d_fact) and pd.notna(d_paie):
                    if d_paie < d_fact:
                        ecart_j = (d_fact - d_paie).days
                        results.append(_a(
                            idx + 2, "Date de paiement",
                            d_paie.strftime("%d/%m/%Y"),
                            f">= {d_fact.strftime('%d/%m/%Y')}",
                            WARNING, CAT_DATE,
                            f"Date de paiement ({d_paie.strftime('%d/%m/%Y')}) anterieure a la facture ({d_fact.strftime('%d/%m/%Y')}) de {ecart_j} jours"
                        ))
            except Exception:
                pass

    # Annee coherente avec TFPB
    if all(c in df.columns for c in ["Date de paiement", "TFPB"]):
        for idx, row in df.iterrows():
            try:
                tfpb = int(float(row.get("TFPB", 0)))
                d_paie = pd.to_datetime(row["Date de paiement"], dayfirst=True)
                if pd.notna(d_paie) and tfpb > 0:
                    annee_attendue = tfpb - 1
                    if d_paie.year != annee_attendue:
                        results.append(_a(
                            idx + 2, "Date de paiement",
                            f"Annee {d_paie.year}",
                            f"Annee {annee_attendue} (TFPB {tfpb})",
                            WARNING, CAT_DATE,
                            f"Paiement en {d_paie.year} mais TFPB {tfpb} attend des paiements en {annee_attendue}"
                        ))
            except Exception:
                pass

    return results


# ═══════════════════════════════════════════════════════════
# 5. CLASSIFICATIONS
# ═══════════════════════════════════════════════════════════

def _check_classifications(df: pd.DataFrame) -> list:
    results = []

    if "Classification des travaux" not in df.columns:
        return results

    for idx, row in df.iterrows():
        classif = row.get("Classification des travaux")
        if pd.isna(classif) or str(classif).strip() in ("", "nan"):
            # Verifier si la ligne a des montants (donc devrait avoir une classification)
            montant = pd.to_numeric(row.get("Montant des travaux éligibles retenus H.T", 0), errors="coerce")
            if pd.notna(montant) and montant > 0:
                results.append(_a(idx + 2, "Classification des travaux", "", "Classification", ERROR, CAT_CLASSIFICATION, "Classification manquante sur une ligne avec montant eligible > 0"))
            continue

        classif_str = str(classif).strip()
        if classif_str not in ALL_KNOWN_CLASSIFICATIONS:
            results.append(_a(idx + 2, "Classification des travaux", classif_str[:60], "Classification connue", WARNING, CAT_CLASSIFICATION, f"Classification inconnue : '{classif_str[:60]}'"))

    # Verifier coherence classification vs nature des travaux
    if "Nature des travaux" in df.columns:
        for idx, row in df.iterrows():
            classif = str(row.get("Classification des travaux", "")).strip()
            nature = str(row.get("Nature des travaux", "")).strip().lower()
            if classif in ("", "nan") or nature in ("", "nan"):
                continue

            # Si classification = etudes mais nature parle de travaux physiques
            if "ETUDES" in classif.upper() or "études" in classif.lower():
                physical_kw = ["isolation", "couverture", "menuiserie", "plomberie", "chauffage"]
                for kw in physical_kw:
                    if kw in nature and "étude" not in nature and "expertise" not in nature:
                        results.append(_a(
                            idx + 2, "Classification vs Nature",
                            f"Classif: etudes / Nature: {nature[:50]}",
                            "Coherence",
                            INFO, CAT_CLASSIFICATION,
                            f"Classification 'Etudes' mais nature semble etre des travaux physiques ('{kw}' detecte)"
                        ))
                        break

    return results


# ═══════════════════════════════════════════════════════════
# 6. COMPLETUDE
# ═══════════════════════════════════════════════════════════

def _check_completude(df: pd.DataFrame) -> list:
    results = []

    critical_fields = {
        "N° d'avis": "N° d'avis d'imposition",
        "N° Fiscal": "N° fiscal",
        "CDIF/ SIP": "CDIF / SIP",
        "Adresse CDIF/SIP": "Adresse CDIF/SIP",
        "Numéro de facture": "N° de facture",
        "Installateur": "Nom de l'installateur",
    }

    for col, label in critical_fields.items():
        if col not in df.columns:
            continue
        for idx, row in df.iterrows():
            val = row.get(col)
            classif = str(row.get("Classification des travaux", "")).strip()
            if classif in EXCLUDED_CLASSIFICATIONS:
                continue
            montant = pd.to_numeric(row.get("Montant des travaux éligibles retenus H.T", 0), errors="coerce")
            if pd.isna(montant) or montant == 0:
                continue
            if pd.isna(val) or str(val).strip() in ("", "nan"):
                sev = ERROR if col in ("N° d'avis", "N° Fiscal") else WARNING
                results.append(_a(idx + 2, col, "", label, sev, CAT_COMPLETUDE, f"{label} manquant sur ligne eligible (montant = {montant:.2f})"))

    # Adresse des travaux
    addr_col = "Adresse des travaux" if "Adresse des travaux" in df.columns else "x"
    if addr_col in df.columns:
        for idx, row in df.iterrows():
            val = row.get(addr_col)
            montant = pd.to_numeric(row.get("Montant des travaux éligibles retenus H.T", 0), errors="coerce")
            if pd.isna(montant) or montant == 0:
                continue
            if pd.isna(val) or str(val).strip() in ("", "nan"):
                results.append(_a(idx + 2, "Adresse des travaux", "", "Adresse", WARNING, CAT_COMPLETUDE, "Adresse des travaux manquante sur ligne eligible"))

    return results


# ═══════════════════════════════════════════════════════════
# 7. DOUBLONS
# ═══════════════════════════════════════════════════════════

def _check_doublons(df: pd.DataFrame) -> list:
    results = []

    dup_cols = []
    if "Numéro de facture" in df.columns:
        dup_cols.append("Numéro de facture")
    if "Installateur" in df.columns:
        dup_cols.append("Installateur")
    if "Montant HT facture " in df.columns:
        dup_cols.append("Montant HT facture ")

    if len(dup_cols) < 2:
        return results

    # Chercher les doublons exacts sur (facture + installateur + montant)
    df_check = df[dup_cols].copy()
    for col in dup_cols:
        df_check[col] = df_check[col].astype(str).str.strip()

    duplicated = df_check.duplicated(keep=False)
    if duplicated.any():
        dup_groups = df_check[duplicated].groupby(dup_cols).groups
        seen = set()
        for key, indices in dup_groups.items():
            key_str = str(key)
            if key_str in seen:
                continue
            seen.add(key_str)
            lignes = [str(i + 2) for i in indices]
            facture = key[0] if len(key) > 0 else "?"
            results.append(_a(
                int(indices[0]) + 2, "Doublons",
                f"Facture {facture}",
                "Unique",
                WARNING, CAT_DOUBLON,
                f"Doublon potentiel detecte sur les lignes {', '.join(lignes)} (meme facture + installateur + montant)"
            ))

    return results


# ═══════════════════════════════════════════════════════════
# 8. COHERENCE TVA
# ═══════════════════════════════════════════════════════════

def _check_tva_coherence(df: pd.DataFrame) -> list:
    results = []

    if "Taux de TVA facture" not in df.columns or "Classification des travaux" not in df.columns:
        return results

    for idx, row in df.iterrows():
        taux = row.get("Taux de TVA facture")
        classif = str(row.get("Classification des travaux", "")).strip()

        if pd.isna(taux) or classif in ("", "nan"):
            continue

        taux_num = pd.to_numeric(taux, errors="coerce")
        if pd.isna(taux_num):
            continue

        # Normaliser : si > 1 c'est un pourcentage, sinon un decimal
        if taux_num > 1:
            taux_pct = taux_num
        else:
            taux_pct = taux_num * 100

        # TVA 5.5% ne devrait etre que sur des travaux d'energie (pas sur etudes)
        if abs(taux_pct - 5.5) < 0.1:
            if "ETUDES" in classif.upper() or "études" in classif.lower() or "expertise" in classif.lower():
                results.append(_a(
                    idx + 2, "Taux de TVA facture",
                    f"{taux_pct}%", "20% pour des etudes",
                    WARNING, CAT_TVA,
                    f"TVA 5,5% appliquee sur des prestations d'etudes (habituellement 20%)"
                ))

        # TVA 20% sur travaux d'isolation/couverture = a verifier
        if abs(taux_pct - 20) < 0.1:
            energy_kw = ["isolation", "couverture", "chauffage", "ventilation"]
            classif_lower = classif.lower()
            for kw in energy_kw:
                if kw in classif_lower and "induit" not in classif_lower and "etude" not in classif_lower:
                    results.append(_a(
                        idx + 2, "Taux de TVA facture",
                        f"20%", "5,5% pour travaux energie",
                        INFO, CAT_TVA,
                        f"TVA 20% sur travaux de type '{kw}' - verifier si le taux reduit 5,5% ne s'applique pas"
                    ))
                    break

        # Taux inconnu
        known_rates = [5.5, 10, 20]
        if not any(abs(taux_pct - r) < 0.5 for r in known_rates):
            results.append(_a(
                idx + 2, "Taux de TVA facture",
                f"{taux_pct}%", "5.5%, 10% ou 20%",
                WARNING, CAT_TVA,
                f"Taux de TVA inhabituel : {taux_pct}%"
            ))

    return results


# ═══════════════════════════════════════════════════════════
# 9. REGROUPEMENT PAR COMMUNE
# ═══════════════════════════════════════════════════════════

def group_by_commune(file_path: Path) -> dict:
    """
    Run full validation then regroup results by commune.
    Returns {
        "summary": {...},
        "communes": {
            "Corbie": {
                "nb_lignes": 30,
                "types": ["TEE", "PMR"],
                "total_ht": 245967.0,
                "total_degrevement": 61491.0,
                "erreurs": 0, "attention": 2, "info": 1,
                "anomalies": [...]
            }, ...
        },
        "global_anomalies": [...]  # anomalies without specific commune (structure, etc.)
    }
    """
    report = validate_file(file_path)

    # Read the dataframe to compute per-commune stats
    try:
        df = pd.read_excel(file_path, sheet_name=0)
    except Exception:
        return {**report, "communes": {}, "global_anomalies": report["anomalies"]}

    if "x" in df.columns:
        df = df.rename(columns={"x": "Adresse des travaux"})

    # Normalize commune names
    if "Commune" in df.columns:
        df["Commune_clean"] = df["Commune"].astype(str).str.strip()
        df["Commune_clean"] = df["Commune_clean"].replace({"nan": "", "None": ""})
    else:
        df["Commune_clean"] = ""

    # Compute per-commune stats
    commune_stats = {}
    eligible_col = "Montant des travaux éligibles retenus H.T"
    degrev_col = "Montant dégrèvement demandé"
    classif_col = "Classification des travaux"

    for commune, grp in df.groupby("Commune_clean"):
        if not commune or commune in ("", "nan"):
            commune = "(Commune vide)"

        ht = pd.to_numeric(grp.get(eligible_col, 0), errors="coerce").fillna(0).sum()
        deg = pd.to_numeric(grp.get(degrev_col, 0), errors="coerce").fillna(0).sum()

        # Determine types present
        types_present = set()
        if classif_col in grp.columns:
            for _, row in grp.iterrows():
                c = str(row.get(classif_col, "")).strip()
                if c in EXCLUDED_CLASSIFICATIONS or c in ("", "nan"):
                    continue
                if c in PMR_CLASSIFICATIONS:
                    types_present.add("PMR")
                elif c in TEE_CLASSIFICATIONS:
                    types_present.add("TEE")
                elif c in INDIS_CLASSIFICATIONS:
                    types_present.add("INDIS")
                elif c in AMBIGUOUS_CLASSIFICATIONS:
                    types_present.add("TEE")
                else:
                    types_present.add("TEE")

        commune_stats[commune] = {
            "nb_lignes": int(len(grp)),
            "types": sorted(types_present),
            "total_ht": float(round(ht, 2)),
            "total_degrevement": float(round(deg, 2)),
            "erreurs": 0,
            "attention": 0,
            "info": 0,
            "anomalies": [],
        }

    # Map anomalies to communes via line numbers
    global_anomalies = []
    for a in report["anomalies"]:
        ligne = a.get("ligne", 0)
        if ligne <= 1:
            global_anomalies.append(a)
            continue

        # Find the commune for this line
        row_idx = ligne - 2  # Excel line -> dataframe index
        if 0 <= row_idx < len(df):
            commune = str(df.iloc[row_idx].get("Commune_clean", "")).strip()
            if not commune or commune in ("", "nan"):
                commune = "(Commune vide)"
        else:
            global_anomalies.append(a)
            continue

        if commune not in commune_stats:
            commune_stats[commune] = {
                "nb_lignes": 0, "types": [], "total_ht": 0, "total_degrevement": 0,
                "erreurs": 0, "attention": 0, "info": 0, "anomalies": [],
            }

        commune_stats[commune]["anomalies"].append(a)
        if a["severite"] == ERROR:
            commune_stats[commune]["erreurs"] += 1
        elif a["severite"] == WARNING:
            commune_stats[commune]["attention"] += 1
        else:
            commune_stats[commune]["info"] += 1

    return {
        "summary": report["summary"],
        "communes": commune_stats,
        "global_anomalies": global_anomalies,
    }


# ═══════════════════════════════════════════════════════════
# 10. COMPARAISON AVEC LES SORTIES SYSTEME
# ═══════════════════════════════════════════════════════════

def compare_with_system(file_path: Path, recap_path: Path) -> dict:
    """
    Compare the source file totals with the system-generated recap.
    Returns per-commune comparison results.
    """
    # Read source and compute expected totals
    commune_report = group_by_commune(file_path)

    # Read the system recap
    try:
        recap_df = pd.read_excel(recap_path, sheet_name="Synthese TFPB", header=None)
    except Exception as e:
        return {
            **commune_report,
            "comparison": {"status": "error", "message": f"Impossible de lire le recap : {e}"},
        }

    # Parse the recap: find the detail rows (after "DETAIL PAR COMMUNE" header)
    system_data = {}
    in_detail = False
    for _, row in recap_df.iterrows():
        val_b = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""

        if val_b == "DETAIL PAR COMMUNE":
            in_detail = True
            continue

        if in_detail and val_b in ("", "TOTAL", "nan"):
            if val_b == "TOTAL":
                break
            continue

        if in_detail:
            # columns: B=Commune, C=Type, D=Nb lignes, E=Nb ops, F=Total HT, G=Sub, H=Base, I=Degrev
            commune = val_b
            work_type = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
            sys_ht = pd.to_numeric(row.iloc[5], errors="coerce") if len(row) > 5 else 0
            sys_deg = pd.to_numeric(row.iloc[8], errors="coerce") if len(row) > 8 else 0
            sys_nb = pd.to_numeric(row.iloc[3], errors="coerce") if len(row) > 3 else 0

            if pd.isna(sys_ht):
                sys_ht = 0
            if pd.isna(sys_deg):
                sys_deg = 0
            if pd.isna(sys_nb):
                sys_nb = 0

            key = f"{commune}|{work_type}"
            system_data[key] = {
                "commune": commune,
                "type": work_type,
                "sys_ht": round(float(sys_ht), 2),
                "sys_degrevement": round(float(sys_deg), 2),
                "sys_nb_lignes": int(sys_nb),
            }

    # Build comparison per commune
    comparisons = {}
    for commune, stats in commune_report["communes"].items():
        if commune == "(Commune vide)":
            continue

        src_ht = stats["total_ht"]
        src_deg = stats["total_degrevement"]
        src_nb = stats["nb_lignes"]

        # Find matching system entries (could be TEE and/or PMR)
        sys_ht_total = 0
        sys_deg_total = 0
        sys_nb_total = 0
        matched_types = []

        for key, sdata in system_data.items():
            sys_commune = sdata["commune"]
            # Match by normalized name
            if sys_commune.upper().replace(" ", "").replace("-", "").replace("_", "") == commune.upper().replace(" ", "").replace("-", "").replace("_", ""):
                sys_ht_total += sdata["sys_ht"]
                sys_deg_total += sdata["sys_degrevement"]
                sys_nb_total += sdata["sys_nb_lignes"]
                matched_types.append(sdata["type"])

        ecart_ht = round(abs(src_ht - sys_ht_total), 2)
        ecart_deg = round(abs(src_deg - sys_deg_total), 2)

        # Status
        if sys_ht_total == 0 and src_ht > 0:
            statut = "NON GENERE"
        elif ecart_ht < 1 and ecart_deg < 1:
            statut = "OK"
        elif ecart_ht < 10:
            statut = "OK (arrondi)"
        else:
            statut = "ECART"

        comparisons[commune] = {
            "source_ht": float(src_ht),
            "source_degrevement": float(src_deg),
            "source_nb_lignes": int(src_nb),
            "systeme_ht": float(sys_ht_total),
            "systeme_degrevement": float(sys_deg_total),
            "systeme_nb_lignes": int(sys_nb_total),
            "ecart_ht": float(ecart_ht),
            "ecart_degrevement": float(ecart_deg),
            "statut": statut,
            "types_generes": matched_types,
        }

    return {
        "summary": commune_report["summary"],
        "communes": commune_report["communes"],
        "global_anomalies": commune_report["global_anomalies"],
        "comparison": {"status": "ok", "data": comparisons},
    }


# ═══════════════════════════════════════════════════════════
# STANDALONE
# ═══════════════════════════════════════════════════════════

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python -m backend.scripts.validate <fichier.xlsx>")
        sys.exit(1)

    file_path = Path(sys.argv[1])
    report = validate_file(file_path)

    print(f"\n{'='*60}")
    print(f"  RAPPORT DE VALIDATION")
    print(f"{'='*60}")
    print(f"  Lignes analysees : {report['summary']['total_lignes']}")
    print(f"  ERREURS   : {report['summary'][ERROR]}")
    print(f"  ATTENTION : {report['summary'][WARNING]}")
    print(f"  INFO      : {report['summary'][INFO]}")
    print(f"{'='*60}\n")

    for a in report["anomalies"]:
        icon = {"ERREUR": "X", "ATTENTION": "!", "INFO": "i"}[a["severite"]]
        print(f"  [{icon}] L{a['ligne']:>4d} | {a['categorie']:15s} | {a['message']}")

    print()
