"""
IndiaFirst Life — Operations Toolkit
Cloud-ready Flask app — deploy to Render.com
  Local  : python app.py  (uses waitress)
  Render : gunicorn via Procfile
"""

import os
import json
import queue
import shutil
import tempfile
import threading
import zipfile
from pathlib import Path

from flask import (Flask, render_template, request,
                   jsonify, Response, stream_with_context, send_file)

from coi_downloader import process_customers

app = Flask(__name__)

CONFIG_FILE   = "config.json"
_job_queue    = queue.Queue()
_job_running  = threading.Event()
_job_zip_path = None   # path to the output ZIP after a job completes

_DEFAULT_CFG = {
    "smtp_host"    : "smtp.gmail.com",
    "smtp_port"    : 587,
    "smtp_user"    : "",
    "smtp_password": "",
}


# ─────────────────────────────────────────────
#  Config — env vars take priority (Render),
#  config.json used for local saves via the UI
# ─────────────────────────────────────────────
def _load_cfg() -> dict:
    cfg = {
        "smtp_host"    : os.environ.get("SMTP_HOST",     _DEFAULT_CFG["smtp_host"]),
        "smtp_port"    : int(os.environ.get("SMTP_PORT", _DEFAULT_CFG["smtp_port"])),
        "smtp_user"    : os.environ.get("SMTP_USER",     _DEFAULT_CFG["smtp_user"]),
        "smtp_password": os.environ.get("SMTP_PASSWORD", _DEFAULT_CFG["smtp_password"]),
    }
    if os.path.isfile(CONFIG_FILE):
        with open(CONFIG_FILE) as f:
            saved = json.load(f)
            cfg.update({k: v for k, v in saved.items() if v})
    return cfg


def _save_cfg(data: dict):
    existing = _load_cfg()
    existing.update({k: v for k, v in data.items() if k in _DEFAULT_CFG})
    with open(CONFIG_FILE, "w") as f:
        json.dump(existing, f, indent=2)


# ─────────────────────────────────────────────
#  Frontend
# ─────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


# ─────────────────────────────────────────────
#  Config API
# ─────────────────────────────────────────────
@app.route("/api/config", methods=["GET"])
def get_config():
    cfg = _load_cfg()
    cfg.pop("smtp_password", None)   # never expose password to browser
    return jsonify(cfg)


@app.route("/api/config", methods=["POST"])
def save_config():
    data = request.get_json() or {}
    _save_cfg(data)
    return jsonify({"status": "ok"})


# ─────────────────────────────────────────────
#  Job status — includes downloadReady flag
# ─────────────────────────────────────────────
@app.route("/api/coi/status", methods=["GET"])
def coi_status():
    return jsonify({
        "running"      : _job_running.is_set(),
        "downloadReady": bool(_job_zip_path and os.path.exists(_job_zip_path)),
    })


# ─────────────────────────────────────────────
#  Start job — accepts multipart/form-data
#  Fields: excel (file), send_email, smtp_*
# ─────────────────────────────────────────────
@app.route("/api/coi/start", methods=["POST"])
def coi_start():
    global _job_zip_path

    if _job_running.is_set():
        return jsonify({"status": "error", "message": "A job is already running."}), 409

    if "excel" not in request.files or request.files["excel"].filename == "":
        return jsonify({"status": "error", "message": "Please upload an Excel file."}), 400

    excel_bytes = request.files["excel"].read()
    send_email  = request.form.get("send_email", "false").lower() == "true"
    try:
        column_mapping = json.loads(request.form.get("column_mapping", "{}"))
    except (json.JSONDecodeError, ValueError):
        column_mapping = {}

    saved_cfg = _load_cfg()
    smtp_cfg = {
        "host"    : request.form.get("smtp_host",     saved_cfg["smtp_host"]),
        "port"    : int(request.form.get("smtp_port", saved_cfg["smtp_port"])),
        "user"    : request.form.get("smtp_user",     saved_cfg["smtp_user"]),
        "password": request.form.get("smtp_password") or saved_cfg.get("smtp_password", ""),
    }

    # Persist updated SMTP settings locally
    _save_cfg({
        "smtp_host"    : smtp_cfg["host"],
        "smtp_port"    : smtp_cfg["port"],
        "smtp_user"    : smtp_cfg["user"],
        **({"smtp_password": smtp_cfg["password"]} if smtp_cfg["password"] else {}),
    })

    def _run():
        global _job_zip_path
        _job_running.set()

        # Flush stale queue messages from previous run
        while not _job_queue.empty():
            try: _job_queue.get_nowait()
            except queue.Empty: break

        # Remove previous ZIP
        if _job_zip_path and os.path.exists(_job_zip_path):
            try: os.remove(_job_zip_path)
            except: pass
        _job_zip_path = None

        tmp_dir = tempfile.mkdtemp(prefix="coi_")
        try:
            process_customers(
                excel_bytes    =excel_bytes,
                output_dir     =tmp_dir,
                smtp_cfg       =smtp_cfg,
                send_email     =send_email,
                log_cb         =_job_queue.put,
                column_mapping =column_mapping,
            )

            # Bundle all downloaded PDFs into a single ZIP for browser download
            pdf_files = list(Path(tmp_dir).glob("*.pdf"))
            if pdf_files:
                _, zip_path = tempfile.mkstemp(suffix=".zip", prefix="coi_out_")
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for pdf in pdf_files:
                        zf.write(pdf, pdf.name)
                _job_zip_path = zip_path   # signals download is ready

        except Exception as exc:
            _job_queue.put({"type": "error", "message": f"Fatal: {exc}"})
            _job_queue.put({"type": "done",
                            "message": "Job ended with error.",
                            "success": 0, "failed": 0, "total": 0})
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            _job_running.clear()

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"status": "started"})


# ─────────────────────────────────────────────
#  Download ZIP of all PDFs from the last job
# ─────────────────────────────────────────────
@app.route("/api/coi/download", methods=["GET"])
def coi_download():
    if not _job_zip_path or not os.path.exists(_job_zip_path):
        return jsonify({"error": "No download available — run a job first."}), 404
    return send_file(
        _job_zip_path,
        as_attachment=True,
        download_name="IndiaFirst_COI_PDFs.zip",
        mimetype="application/zip",
    )


# ─────────────────────────────────────────────
#  SSE stream — live logs pushed to the browser
# ─────────────────────────────────────────────
@app.route("/api/coi/stream", methods=["GET"])
def coi_stream():
    def _generate():
        while True:
            try:
                msg = _job_queue.get(timeout=25)
                yield f"data: {json.dumps(msg)}\n\n"
                if msg.get("type") == "done":
                    break
            except queue.Empty:
                yield 'data: {"type":"ping"}\n\n'   # keep connection alive

    return Response(
        stream_with_context(_generate()),
        content_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ─────────────────────────────────────────────
#  Entry point — local only (Render uses Procfile)
# ─────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    print(f"  Starting locally on http://localhost:{port}")
    try:
        from waitress import serve
        serve(app, host="0.0.0.0", port=port, threads=8)
    except ImportError:
        app.run(host="0.0.0.0", port=port, threaded=True)