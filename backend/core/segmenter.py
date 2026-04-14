import pandas as pd
from backend.config.constants import (
    TEE_CLASSIFICATIONS,
    INDIS_CLASSIFICATIONS,
    PMR_CLASSIFICATIONS,
    EXCLUDED_CLASSIFICATIONS,
    AMBIGUOUS_CLASSIFICATIONS,
)
from backend.utils.logger import get_logger

log = get_logger()

# Mots-cles PMR pour classifier les lignes ambigues
PMR_KEYWORDS = [
    "pmr", "accessibilité", "handicap", "personnes âgées",
    "rampe", "ascenseur", "main courante", "antidérapant",
    "élargissement", "porte 90", "seuil", "ressaut",
    "douche", "adaptation", "mobilité réduite",
]


def _normalize_apostrophes(s: str) -> str:
    """Replace curly apostrophes with straight ones for matching."""
    return s.replace("\u2019", "'").replace("\u2018", "'")


def build_programme_map(df: pd.DataFrame) -> dict[str, str]:
    """
    Build a map of {N° Programme -> dominant type ('TEE' or 'PMR')}.
    For each programme, sum the HT of direct TEE and direct PMR works.
    The dominant type is the one with the highest HT total.
    If no direct TEE or PMR works exist, default to 'TEE'.
    """
    prog_col = "N° Programme"
    classif_col = "Classification des travaux"
    ht_col = "Montant des travaux éligibles retenus H.T"

    if prog_col not in df.columns:
        return {}

    programme_totals = {}
    for _, row in df.iterrows():
        prog = str(row.get(prog_col, "")).strip()
        if prog in ("", "nan"):
            continue
        classif = _normalize_apostrophes(str(row.get(classif_col, "")).strip())
        ht = pd.to_numeric(row.get(ht_col, 0), errors="coerce")
        if pd.isna(ht):
            ht = 0

        if prog not in programme_totals:
            programme_totals[prog] = {"tee_ht": 0, "pmr_ht": 0}

        # Count DIRECT TEE, PMR, and ambiguous (via keywords)
        if classif in TEE_CLASSIFICATIONS:
            programme_totals[prog]["tee_ht"] += ht
        elif classif in PMR_CLASSIFICATIONS:
            programme_totals[prog]["pmr_ht"] += ht
        elif classif in AMBIGUOUS_CLASSIFICATIONS:
            # Check keywords to decide TEE or PMR
            nature = str(row.get("Nature des travaux", "")).strip().lower()
            detail = str(row.get("Détail des travaux retenus", "")).strip().lower()
            text = f"{nature} {detail}"
            is_pmr = any(kw in text for kw in PMR_KEYWORDS)
            if is_pmr:
                programme_totals[prog]["pmr_ht"] += ht
            else:
                programme_totals[prog]["tee_ht"] += ht

    # Build the map: dominant type per programme
    programme_map = {}
    for prog, totals in programme_totals.items():
        if totals["pmr_ht"] > totals["tee_ht"] and totals["pmr_ht"] > 0:
            programme_map[prog] = "PMR"
        elif totals["tee_ht"] > 0:
            programme_map[prog] = "TEE"
        else:
            # No direct works found -> default TEE
            programme_map[prog] = "TEE"

    log.info(f"Carte des programmes : {len(programme_map)} programmes "
             f"({sum(1 for v in programme_map.values() if v == 'TEE')} TEE, "
             f"{sum(1 for v in programme_map.values() if v == 'PMR')} PMR)")

    return programme_map


def classify_row(row: pd.Series, programme_map: dict | None = None) -> str | None:
    """
    Classify a single row as 'TEE', 'PMR', 'INDIS_TEE', 'INDIS_PMR', or None (excluded).
    - TEE direct = Economies d'energie
    - PMR direct = Accessibilite handicap
    - INDIS = tous les autres travaux eligibles (isolation, chauffage, etudes, induits...)
    """
    classification = _normalize_apostrophes(str(row.get("Classification des travaux", "")).strip())
    nature = str(row.get("Nature des travaux", "")).strip().lower()
    detail = str(row.get("Détail des travaux retenus", "")).strip().lower()
    programme = str(row.get("N° Programme", "")).strip()

    # Excluded
    if classification in EXCLUDED_CLASSIFICATIONS:
        return None

    # Direct TEE = seulement Economies d'energie
    if classification in TEE_CLASSIFICATIONS:
        return "TEE"

    # Direct PMR = Accessibilite handicap
    if classification in PMR_CLASSIFICATIONS:
        return "PMR"

    # Travaux indissociablement lies (INDIS): follow programme dominant
    if classification in INDIS_CLASSIFICATIONS:
        if programme_map and programme in programme_map:
            return f"INDIS_{programme_map[programme]}"
        return "INDIS_TEE"

    # Ambiguous: keywords first, then programme, then default TEE
    if classification in AMBIGUOUS_CLASSIFICATIONS:
        text = f"{nature} {detail}"
        for kw in PMR_KEYWORDS:
            if kw in text:
                return "PMR"
        if programme_map and programme in programme_map:
            return programme_map[programme]
        return "TEE"

    # Unknown classification
    if classification and classification != "nan":
        log.warning(f"Classification inconnue : '{classification}' -> classee TEE par defaut")
        return "TEE"

    return None


def segment_data(df: pd.DataFrame) -> dict[tuple[str, str], pd.DataFrame]:
    """
    Segment data by (Commune, Type).
    Returns dict of {(commune, type): filtered_dataframe}
    """
    log.info("Segmentation des donnees par Commune et Type de travaux...")

    df = df.copy()

    # Build programme context map for linked works
    programme_map = build_programme_map(df)

    df["Type_travaux"] = df.apply(lambda row: classify_row(row, programme_map), axis=1)

    # Remove excluded rows
    excluded_count = df["Type_travaux"].isna().sum()
    if excluded_count:
        log.info(f"{excluded_count} lignes exclues (non eligibles ou classification vide)")
    df_eligible = df[df["Type_travaux"].notna()].copy()

    # Use normalized commune if available
    commune_col = "Commune_normalized" if "Commune_normalized" in df_eligible.columns else "Commune"

    segments = {}
    for (commune, work_type), group in df_eligible.groupby([commune_col, "Type_travaux"]):
        commune_str = str(commune).strip()
        if commune_str in ("", "nan", "None"):
            continue
        key = (commune_str, work_type)
        segments[key] = group.copy()
        log.info(f"  Segment [{commune_str} / {work_type}] : {len(group)} lignes")

    log.info(f"Total : {len(segments)} segments identifies")
    return segments


def build_synthesis(segment_df: pd.DataFrame) -> dict:
    """Build a synthesis dict for a segment (commune + type)."""
    total_ht = segment_df["Montant des travaux éligibles retenus H.T"].sum()
    total_degrevement = segment_df["Montant dégrèvement demandé"].sum()
    total_subventions = segment_df["Montant subventions encaisses"].sum()
    total_virement_ttc = segment_df["Montant de virement TTC"].sum()

    # Group by TVA rate
    tva_groups = {}
    if "Taux de TVA facture" in segment_df.columns:
        for rate, grp in segment_df.groupby("Taux de TVA facture"):
            if pd.notna(rate) and rate > 0:
                tva_groups[float(rate)] = {
                    "montant_ht": grp["Montant des travaux éligibles retenus H.T"].sum(),
                    "nb_factures": len(grp),
                }

    # Group by operation
    operations = {}
    op_col = "N°OPERATION" if "N°OPERATION" in segment_df.columns else None
    if op_col:
        for op, grp in segment_df.groupby(op_col):
            op_str = str(op).strip()
            if op_str in ("", "nan"):
                continue
            adresses = grp["Adresse des travaux"].unique() if "Adresse des travaux" in grp.columns else []
            operations[op_str] = {
                "adresses": [str(a) for a in adresses if str(a) != "nan"],
                "montant_ht_etudes": grp[
                    grp["Classification des travaux"].str.contains("ETUDES|études|Études", case=False, na=False)
                ]["Montant des travaux éligibles retenus H.T"].sum(),
                "montant_ht_travaux": grp[
                    ~grp["Classification des travaux"].str.contains("ETUDES|études|Études", case=False, na=False)
                ]["Montant des travaux éligibles retenus H.T"].sum(),
                "montant_ht_total": grp["Montant des travaux éligibles retenus H.T"].sum(),
                "nb_logements": _extract_nb_logements(grp),
                "nb_factures": len(grp),
            }

    # N° d'avis and N° Fiscal (take the first non-null)
    avis = segment_df["N° d'avis"].dropna().unique()
    fiscal = segment_df["N° Fiscal"].dropna().unique()
    cdif = segment_df["CDIF/ SIP"].dropna().unique()
    adresse_cdif = segment_df["Adresse CDIF/SIP"].dropna().unique()

    # Programme numbers
    programmes = segment_df["N° Programme"].dropna().unique()

    return {
        "total_ht_eligible": round(total_ht, 2),
        "total_degrevement": round(total_degrevement, 2),
        "total_subventions": round(total_subventions, 2),
        "base_nette": round(total_ht - total_subventions, 2),
        "total_virement_ttc": round(total_virement_ttc, 2),
        "tva_groups": tva_groups,
        "operations": operations,
        "nb_operations": len(operations),
        "nb_factures": len(segment_df),
        "num_avis": list(avis),
        "num_fiscal": list(fiscal),
        "cdif": list(cdif),
        "adresse_cdif": list(adresse_cdif),
        "programmes": [str(p) for p in programmes if str(p) != "nan"],
    }


def _extract_nb_logements(grp: pd.DataFrame) -> str:
    """Try to extract number of logements from detail text."""
    for detail in grp.get("Détail des travaux retenus", []):
        if pd.isna(detail):
            continue
        detail = str(detail).lower()
        for pattern in ["réhabilitation de ", "rehabilitation de "]:
            if pattern in detail:
                after = detail.split(pattern)[1]
                num = ""
                for c in after:
                    if c.isdigit():
                        num += c
                    else:
                        break
                if num:
                    return num
    return ""
