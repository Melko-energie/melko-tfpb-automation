import pandas as pd
from pathlib import Path
from backend.config.constants import REQUIRED_COLUMNS
from backend.utils.logger import get_logger

log = get_logger()


def read_excel(file_path: Path) -> pd.DataFrame:
    """Read the input Excel file and return the main data sheet."""
    log.info(f"Lecture du fichier Excel : {file_path.name}")

    if not file_path.exists():
        raise FileNotFoundError(f"Fichier introuvable : {file_path}")

    if not file_path.suffix.lower() in (".xlsx", ".xls"):
        raise ValueError(f"Format non supporté : {file_path.suffix}")

    df = pd.read_excel(file_path, sheet_name=0)
    log.info(f"Fichier lu : {len(df)} lignes, {len(df.columns)} colonnes")
    return df


def validate_columns(df: pd.DataFrame) -> list[str]:
    """Check that all required columns exist. Returns list of missing columns."""
    missing = []
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            missing.append(col)
    if missing:
        log.warning(f"Colonnes manquantes : {missing}")
    else:
        log.info("Toutes les colonnes obligatoires sont presentes")
    return missing


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Clean and normalize the dataframe."""
    log.info("Nettoyage des donnees...")
    df = df.copy()

    # Rename column 'x' to a proper name
    if "x" in df.columns:
        df.rename(columns={"x": "Adresse des travaux"}, inplace=True)

    # Normalize commune names (strip, title case, merge variants)
    if "Commune" in df.columns:
        df["Commune"] = df["Commune"].astype(str).str.strip()

        def _normalize_key(s):
            """Create a merge key: uppercase, strip accents, replace hyphens/underscores with space."""
            import re
            import unicodedata
            s = str(s).strip().upper()
            # Remove accents: É->E, è->E, etc.
            s = unicodedata.normalize("NFD", s)
            s = "".join(c for c in s if unicodedata.category(c) != "Mn")
            s = s.replace("-", " ").replace("_", " ").replace("'", " ").replace("\u2019", " ")
            s = re.sub(r"\s+", " ", s).strip()
            return s

        commune_map = {}
        for c in df["Commune"].unique():
            c_str = str(c).strip()
            if c_str in ("", "nan", "None"):
                continue
            key = _normalize_key(c_str)
            if key not in commune_map:
                commune_map[key] = key.title()

        df["Commune_normalized"] = df["Commune"].apply(
            lambda x: commune_map.get(_normalize_key(x), None) if str(x).strip() not in ("", "nan", "None") else None
        )

    # Ensure numeric columns are numeric
    numeric_cols = [
        "Montant HT facture ",
        "TVA facture",
        "Montant TTC facture",
        "Montant de virement TTC",
        "Montant de virement HT",
        "Montant subventions encaisses",
        "Montant des travaux éligibles retenus H.T",
        "Montant dégrèvement demandé",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Ensure Taux de TVA is numeric
    if "Taux de TVA facture" in df.columns:
        df["Taux de TVA facture"] = pd.to_numeric(df["Taux de TVA facture"], errors="coerce")

    anomalies = 0
    for idx, row in df.iterrows():
        if pd.isna(row.get("Commune")) or str(row.get("Commune")).strip() in ("", "nan"):
            log.debug(f"Ligne {idx+2}: Commune manquante")
            anomalies += 1
        if pd.isna(row.get("Classification des travaux")) or str(row.get("Classification des travaux")).strip() in ("", "nan"):
            log.debug(f"Ligne {idx+2}: Classification manquante")
            anomalies += 1

    if anomalies:
        log.warning(f"{anomalies} anomalies detectees dans les donnees")
    else:
        log.info("Aucune anomalie detectee")

    return df
