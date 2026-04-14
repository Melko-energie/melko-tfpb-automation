import shutil
from pathlib import Path
from backend.core.reader import read_excel, validate_columns, clean_dataframe
from backend.core.segmenter import segment_data, build_synthesis
from backend.generators.annexe_generator import generate_annexe
from backend.generators.courrier_generator import generate_courrier
from backend.generators.recap_generator import generate_recap
from backend.generators.tcd_generator import generate_tcd
from backend.generators.verification_generator import generate_verification
from backend.scripts.validate import group_by_commune
from backend.config.constants import OUTPUT_DIR
from backend.utils.logger import get_logger

log = get_logger()


def process_file(file_path: Path) -> dict:
    """
    Main processing pipeline:
    1. Read & validate Excel
    2. Clean data
    3. Segment by Commune + Type
    4. Generate annexe + courrier for each segment
    5. Return summary
    """
    log.info("=" * 60)
    log.info("DEMARRAGE DU TRAITEMENT")
    log.info("=" * 60)

    # Step 1: Read
    df = read_excel(file_path)
    missing = validate_columns(df)
    if missing:
        log.error(f"Colonnes manquantes critiques : {missing}")

    # Step 2: Clean
    df = clean_dataframe(df)

    # Step 3: Segment
    segments = segment_data(df)

    if not segments:
        log.warning("Aucun segment eligible trouve dans le fichier")
        return {"status": "warning", "message": "Aucun segment eligible", "segments": []}

    # Step 4: Clean output dir
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Step 5: Generate files
    results = []
    demande_counter = 1

    for (commune, work_type), segment_df in sorted(segments.items()):
        log.info(f"--- Traitement : {commune} / {work_type} ({len(segment_df)} lignes) ---")

        # Build synthesis
        synthesis = build_synthesis(segment_df)

        # Output path
        commune_safe = commune.replace(" ", "_").replace("/", "-").replace("'", "")
        output_path = OUTPUT_DIR / commune_safe / work_type

        # Generate annexe
        try:
            annexe_path = generate_annexe(segment_df, synthesis, commune, work_type, output_path)
        except Exception as e:
            log.error(f"Erreur generation annexe {commune}/{work_type}: {e}")
            annexe_path = None

        # Generate courrier
        try:
            courrier_path = generate_courrier(synthesis, commune, work_type, output_path, demande_counter)
        except Exception as e:
            log.error(f"Erreur generation courrier {commune}/{work_type}: {e}")
            courrier_path = None

        # Generate TCD
        try:
            tcd_path = generate_tcd(segment_df, synthesis, commune, work_type, output_path)
        except Exception as e:
            log.error(f"Erreur generation TCD {commune}/{work_type}: {e}")
            tcd_path = None

        results.append({
            "commune": commune,
            "type": work_type,
            "nb_lignes": len(segment_df),
            "total_degrevement": synthesis["total_degrevement"],
            "total_ht": synthesis["total_ht_eligible"],
            "nb_operations": synthesis["nb_operations"],
            "annexe": str(annexe_path) if annexe_path else None,
            "courrier": str(courrier_path) if courrier_path else None,
            "synthesis": synthesis,
        })

        demande_counter += 1

    # Step 6: Generate recap synthesis
    recap_path = None
    try:
        recap_path = generate_recap(results)
    except Exception as e:
        log.error(f"Erreur generation recap : {e}")

    # Step 7: Generate verification report
    verif_path = None
    try:
        commune_report = group_by_commune(file_path)
        verif_path = generate_verification(commune_report)
    except Exception as e:
        log.error(f"Erreur generation verification : {e}")

    log.info("=" * 60)
    log.info(f"TRAITEMENT TERMINE : {len(results)} dossiers generes")
    log.info("=" * 60)

    # Remove synthesis from response (too large for JSON)
    segments_clean = []
    for r in results:
        seg = {k: v for k, v in r.items() if k != "synthesis"}
        segments_clean.append(seg)

    return {
        "status": "success",
        "total_segments": len(segments_clean),
        "segments": segments_clean,
        "recap": str(recap_path) if recap_path else None,
        "verification": str(verif_path) if verif_path else None,
    }
