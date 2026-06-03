"""
Zero Man — COI Downloader
Cloud-ready Flask app (Render.com)
"""

import os, json, queue, shutil, tempfile, threading, zipfile
from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request, jsonify, Response, stream_with_context, send_file

from coi_downloader import process_customers

app = Flask(__name__)

CONFIG_FILE   = "config.json"
_job_queue    = queue.Queue()
_job_running  = threading.Event()
_job_zip_path = None
_job_zip_name = "COI_Downloads.zip"
_job_log_path = None   # path to the Excel log file after job
_job_results  = []     # per-row result dicts collected during the job

_DEFAULT_CFG = {
    "smtp_host": "smtp.gmail.com", "smtp_port": 587,
    "smtp_user": "", "smtp_password": "",
}


# ─────────────────────────────────────────────
#  Config
# ─────────────────────────────────────────────
def _load_cfg():
    if os.path.isfile(CONFIG_FILE):
        with open(CONFIG_FILE) as f:
            return {**_DEFAULT_CFG, **json.load(f)}
    return _DEFAULT_CFG.copy()

def _save_cfg(data):
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
    cfg = _load_cfg(); cfg.pop("smtp_password", None)
    return jsonify(cfg)

@app.route("/api/config", methods=["POST"])
def save_config():
    _save_cfg(request.get_json() or {})
    return jsonify({"status": "ok"})


# ─────────────────────────────────────────────
#  Job Status
# ─────────────────────────────────────────────
@app.route("/api/coi/status", methods=["GET"])
def coi_status():
    return jsonify({
        "running"      : _job_running.is_set(),
        "downloadReady": bool(_job_zip_path and os.path.exists(_job_zip_path)),
        "logReady"     : bool(_job_log_path and os.path.exists(_job_log_path)),
    })


# ─────────────────────────────────────────────
#  Start Job
# ─────────────────────────────────────────────
@app.route("/api/coi/start", methods=["POST"])
def coi_start():
    global _job_zip_path, _job_zip_name, _job_log_path, _job_results

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

    raw_name   = request.form.get("batch_name", "") or "COI_Downloads"
    batch_name = "".join(c for c in raw_name.strip() if c.isalnum() or c in "._- ") or "COI_Downloads"

    saved_cfg = _load_cfg()
    smtp_cfg = {
        "host"    : request.form.get("smtp_host",     saved_cfg["smtp_host"]),
        "port"    : int(request.form.get("smtp_port", saved_cfg["smtp_port"])),
        "user"    : request.form.get("smtp_user",     saved_cfg["smtp_user"]),
        "password": request.form.get("smtp_password") or saved_cfg.get("smtp_password", ""),
    }
    _save_cfg({"smtp_host": smtp_cfg["host"], "smtp_port": smtp_cfg["port"],
               "smtp_user": smtp_cfg["user"],
               **({"smtp_password": smtp_cfg["password"]} if smtp_cfg["password"] else {})})

    def _run():
        global _job_zip_path, _job_zip_name, _job_log_path, _job_results
        _job_running.set()

        # Flush queue
        while not _job_queue.empty():
            try: _job_queue.get_nowait()
            except queue.Empty: break

        # Clean up previous outputs
        for path in [_job_zip_path, _job_log_path]:
            if path and os.path.exists(path):
                try: os.remove(path)
                except: pass
        _job_zip_path = _job_log_path = None
        _job_results  = []

        tmp_dir = tempfile.mkdtemp(prefix="coi_")
        try:
            process_customers(
                excel_bytes    = excel_bytes,
                output_dir     = tmp_dir,
                smtp_cfg       = smtp_cfg,
                send_email     = send_email,
                log_cb         = _job_queue.put,
                column_mapping = column_mapping,
                result_cb      = lambda r: _job_results.append(r),
            )

            # ── Bundle PDFs into ZIP ──
            pdf_files = list(Path(tmp_dir).glob("*.pdf"))
            if pdf_files:
                _, zip_path = tempfile.mkstemp(suffix=".zip", prefix="coi_out_")
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for pdf in pdf_files:
                        zf.write(pdf, pdf.name)
                _job_zip_path = zip_path
                _job_zip_name = f"{batch_name}_COIs.zip"

            # ── Build log Excel ──
            if _job_results:
                log_df = pd.DataFrame(_job_results)
                # Move Status and Error Message to the last two columns
                meta_cols = [c for c in ("Status", "Error Message") if c in log_df.columns]
                other_cols = [c for c in log_df.columns if c not in meta_cols]
                log_df = log_df[other_cols + meta_cols]
                _, log_path = tempfile.mkstemp(suffix=".xlsx", prefix="coi_log_")
                log_df.to_excel(log_path, index=False)
                _job_log_path = log_path

        except Exception as exc:
            _job_queue.put({"type": "error", "message": f"Fatal: {exc}"})
            _job_queue.put({"type": "done",  "message": "Job ended with error.",
                            "success": 0, "failed": 0, "total": 0})
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            _job_running.clear()

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"status": "started"})


# ─────────────────────────────────────────────
#  Download ZIP
# ─────────────────────────────────────────────
@app.route("/api/coi/download", methods=["GET"])
def coi_download():
    if not _job_zip_path or not os.path.exists(_job_zip_path):
        return jsonify({"error": "No ZIP available."}), 404
    return send_file(_job_zip_path, as_attachment=True,
                     download_name=_job_zip_name, mimetype="application/zip")


# ─────────────────────────────────────────────
#  Download Log Excel
# ─────────────────────────────────────────────
@app.route("/api/coi/download-log", methods=["GET"])
def coi_download_log():
    if not _job_log_path or not os.path.exists(_job_log_path):
        return jsonify({"error": "No log available."}), 404
    return send_file(
        _job_log_path,
        as_attachment=True,
        download_name="COI_Job_Log.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ─────────────────────────────────────────────
#  SSE Stream
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
                yield 'data: {"type":"ping"}\n\n'

    return Response(stream_with_context(_generate()),
                    content_type="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ─────────────────────────────────────────────
#  Entry Point (local only — Render uses Procfile)
# ─────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"  Starting on http://localhost:{port}")
    try:
        from waitress import serve
        serve(app, host="0.0.0.0", port=port, threads=8)
    except ImportError:
        app.run(host="0.0.0.0", port=port, threaded=True)
