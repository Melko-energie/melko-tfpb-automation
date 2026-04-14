import shutil
import tempfile
from pathlib import Path

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS

from backend.core.processor import process_file
from backend.scripts.validate import validate_file, group_by_commune, compare_with_system
from backend.config.constants import OUTPUT_DIR
from backend.utils.logger import get_logger, get_log_buffer, clear_log_buffer

log = get_logger()

app = Flask(__name__, static_folder=None)
CORS(app)

FRONTEND_DIR = Path(__file__).resolve().parent.parent / "frontend"

_state = {"status": "idle", "result": None, "progress": 0}


@app.route("/")
def serve_index():
    return send_file(str(FRONTEND_DIR / "index.html"))


@app.route("/static/<path:filename>")
def serve_static(filename):
    return send_from_directory(str(FRONTEND_DIR), filename)


@app.route("/api/upload", methods=["POST"])
def upload_and_process():
    if "file" not in request.files:
        return jsonify({"detail": "Aucun fichier envoye"}), 400

    file = request.files["file"]
    if not file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"detail": "Le fichier doit etre un fichier Excel (.xlsx ou .xls)"}), 400

    _state["status"] = "processing"
    _state["progress"] = 0
    _state["result"] = None
    clear_log_buffer()

    tmp_dir = Path(tempfile.mkdtemp())
    tmp_file = tmp_dir / file.filename
    file.save(str(tmp_file))

    log.info(f"Fichier recu : {file.filename}")

    try:
        result = process_file(tmp_file)
        _state["status"] = "done"
        _state["result"] = result
        _state["progress"] = 100
        return jsonify(result)
    except Exception as e:
        log.error(f"Erreur lors du traitement : {str(e)}")
        _state["status"] = "error"
        _state["result"] = {"error": str(e)}
        return jsonify({"detail": f"Erreur de traitement : {str(e)}"}), 500
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.route("/api/status")
def get_status():
    return jsonify({
        "status": _state["status"],
        "progress": _state["progress"],
        "result": _state["result"],
    })


@app.route("/api/logs")
def get_logs():
    return jsonify({"logs": get_log_buffer()})


@app.route("/api/download")
def download_output():
    if not OUTPUT_DIR.exists() or not any(OUTPUT_DIR.iterdir()):
        return jsonify({"detail": "Aucun fichier genere"}), 404

    zip_path = Path(tempfile.mktemp(suffix=".zip"))
    shutil.make_archive(str(zip_path.with_suffix("")), "zip", str(OUTPUT_DIR))
    return send_file(str(zip_path), as_attachment=True, download_name="dossiers_TFPB.zip")


@app.route("/api/download/<commune>/<work_type>/<filename>")
def download_file(commune, work_type, filename):
    file_path = OUTPUT_DIR / commune / work_type / filename
    if not file_path.exists():
        return jsonify({"detail": f"Fichier introuvable : {filename}"}), 404
    return send_file(str(file_path), as_attachment=True, download_name=filename)


# ── Validation routes ──

@app.route("/validation")
def serve_validation():
    return send_file(str(FRONTEND_DIR / "validation.html"))


@app.route("/api/validate", methods=["POST"])
def api_validate():
    if "file" not in request.files:
        return jsonify({"detail": "Aucun fichier envoye"}), 400

    file = request.files["file"]
    if not file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"detail": "Le fichier doit etre un fichier Excel (.xlsx ou .xls)"}), 400

    tmp_dir = Path(tempfile.mkdtemp())
    tmp_file = tmp_dir / file.filename
    file.save(str(tmp_file))

    try:
        report = validate_file(tmp_file)
        return jsonify(report)
    except Exception as e:
        return jsonify({"detail": f"Erreur de validation : {str(e)}"}), 500
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.route("/api/download-verification")
def download_verification():
    verif_path = OUTPUT_DIR / "verification_donnees.xlsx"
    if not verif_path.exists():
        return jsonify({"detail": "Rapport de verification non encore genere"}), 404
    return send_file(str(verif_path), as_attachment=True, download_name="verification_donnees.xlsx")


@app.route("/api/validate-communes", methods=["POST"])
def api_validate_communes():
    """Validate and group results by commune."""
    if "file" not in request.files:
        return jsonify({"detail": "Aucun fichier envoye"}), 400
    file = request.files["file"]
    if not file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"detail": "Format invalide"}), 400
    tmp_dir = Path(tempfile.mkdtemp())
    tmp_file = tmp_dir / file.filename
    file.save(str(tmp_file))
    try:
        report = group_by_commune(tmp_file)
        return jsonify(report)
    except Exception as e:
        return jsonify({"detail": str(e)}), 500
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.route("/api/validate-compare", methods=["POST"])
def api_validate_compare():
    """Validate source file and compare with system recap."""
    if "file" not in request.files:
        return jsonify({"detail": "Aucun fichier envoye"}), 400
    file = request.files["file"]
    if not file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"detail": "Format invalide"}), 400
    recap_path = OUTPUT_DIR / "recap_synthese.xlsx"
    if not recap_path.exists():
        return jsonify({"detail": "Lancez d'abord le traitement sur la page principale pour generer le recap"}), 404
    tmp_dir = Path(tempfile.mkdtemp())
    tmp_file = tmp_dir / file.filename
    file.save(str(tmp_file))
    try:
        report = compare_with_system(tmp_file, recap_path)
        return jsonify(report)
    except Exception as e:
        return jsonify({"detail": str(e)}), 500
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.route("/api/download-recap")
def download_recap():
    recap_path = OUTPUT_DIR / "recap_synthese.xlsx"
    if not recap_path.exists():
        return jsonify({"detail": "Recap non encore genere"}), 404
    return send_file(str(recap_path), as_attachment=True, download_name="recap_synthese.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
