from __future__ import annotations

import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path
from uuid import uuid4

try:
    from flask import Flask, jsonify, render_template, request, send_file
    from werkzeug.exceptions import RequestEntityTooLarge
    from werkzeug.utils import secure_filename
except ModuleNotFoundError as exc:  # pragma: no cover - user-facing startup helper
    missing = exc.name or "a required package"
    raise SystemExit(
        f"Missing dependency: {missing}. Run 'pip install -r requirements.txt' and then start again."
    ) from exc

from processor import ExcelProcessingError, process_excel


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "storage" / "uploads"
OUTPUT_DIR = BASE_DIR / "storage" / "outputs"
LOG_DIR = BASE_DIR / "logs"
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
    app.config["UPLOAD_DIR"] = UPLOAD_DIR
    app.config["OUTPUT_DIR"] = OUTPUT_DIR

    ensure_runtime_directories()
    configure_logging(app)

    @app.get("/")
    def index():
        return render_template("index.html")

    @app.post("/process")
    def process_upload():
        uploaded_file = request.files.get("file")
        if uploaded_file is None or uploaded_file.filename == "":
            return jsonify({"error": "Choose an Excel file before processing."}), 400

        original_filename = secure_filename(uploaded_file.filename)
        extension = Path(original_filename).suffix.lower()
        if extension not in ALLOWED_EXTENSIONS:
            return jsonify({"error": "Only .xlsx and .xlsm files are supported."}), 400

        job_id = uuid4().hex
        upload_path = UPLOAD_DIR / f"{job_id}_{original_filename}"
        output_path = OUTPUT_DIR / f"{job_id}_processed.xlsx"

        try:
            uploaded_file.save(upload_path)
            app.logger.info("Processing upload %s as job %s", original_filename, job_id)
            summary = process_excel(upload_path, output_path)
        except ExcelProcessingError as exc:
            app.logger.warning("Excel validation failed for job %s: %s", job_id, exc)
            return jsonify({"error": str(exc)}), 400
        except Exception:
            app.logger.exception("Unexpected processing failure for job %s", job_id)
            return jsonify({"error": "Processing failed. Check logs/app.log for details."}), 500

        download_name = build_download_name(original_filename)
        app.logger.info("Finished job %s", job_id)
        return jsonify(
            {
                "message": "File processed successfully.",
                "download_url": f"/download/{job_id}",
                "download_name": download_name,
                "summary": summary,
            }
        )

    @app.get("/download/<job_id>")
    def download(job_id: str):
        if not is_hex_job_id(job_id):
            return jsonify({"error": "Invalid download id."}), 400

        output_path = OUTPUT_DIR / f"{job_id}_processed.xlsx"
        if not output_path.exists():
            return jsonify({"error": "Processed file not found. Please upload and process again."}), 404

        return send_file(
            output_path,
            as_attachment=True,
            download_name=f"name_masking_processed_{job_id[:8]}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.errorhandler(RequestEntityTooLarge)
    def handle_large_file(_: RequestEntityTooLarge):
        return jsonify({"error": "File is too large. Maximum upload size is 100 MB."}), 413

    return app


def ensure_runtime_directories() -> None:
    for directory in (UPLOAD_DIR, OUTPUT_DIR, LOG_DIR):
        directory.mkdir(parents=True, exist_ok=True)


def configure_logging(app: Flask) -> None:
    log_path = LOG_DIR / "app.log"
    handler = RotatingFileHandler(log_path, maxBytes=2_000_000, backupCount=5, encoding="utf-8")
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s [%(name)s] %(message)s")
    )
    handler.setLevel(logging.INFO)
    app.logger.setLevel(logging.INFO)
    if not any(isinstance(existing, RotatingFileHandler) for existing in app.logger.handlers):
        app.logger.addHandler(handler)


def build_download_name(original_filename: str) -> str:
    stem = Path(original_filename).stem or "processed"
    return f"{stem}_processed.xlsx"


def is_hex_job_id(value: str) -> bool:
    return len(value) == 32 and all(char in "0123456789abcdef" for char in value.lower())


app = create_app()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "70"))
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
