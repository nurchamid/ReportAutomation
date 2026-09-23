"""
app.py
Flask web application for Report Automation.
Run: python app.py
Access: http://localhost:5000
"""

import os
import uuid
import threading
import queue
import json
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, send_file, jsonify, Response, flash
)
from werkzeug.utils import secure_filename

import pandas as pd
from report_engine import (
    load_and_process_car_sales_data,
    generate_report,
    send_report_email_multi,
    load_email_list_from_excel,
)

# ─── App Config ───────────────────────────────────────────────────────────────

app = Flask(__name__)
app.secret_key = "report-automation-secret-2026"

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR  = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR  = os.path.join(BASE_DIR, "outputs")
EMAIL_EXCEL = os.path.join(BASE_DIR, "list_user_email.xlsx")

ALLOWED_EXT = {"csv"}
MAX_CONTENT  = 50 * 1024 * 1024   # 50 MB

app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT

# In-memory stores keyed by session_id
_progress_queues: dict[str, queue.Queue] = {}
_df_store:        dict[str, pd.DataFrame] = {}
_output_store:    dict[str, dict] = {}

# ─── Credentials (hardcoded) ──────────────────────────────────────────────────

VALID_USERS = {
    "admin": "admin123",
}

# ─── Helpers ──────────────────────────────────────────────────────────────────

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXT


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


def get_sid():
    """Return (and create if needed) a stable session-scoped ID."""
    if "sid" not in session:
        session["sid"] = str(uuid.uuid4())
    return session["sid"]


# ─── Auth Routes ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return redirect(url_for("dashboard") if session.get("logged_in") else url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        if VALID_USERS.get(username) == password:
            session["logged_in"] = True
            session["username"]  = username
            return redirect(url_for("dashboard"))
        else:
            error = "Username atau password salah."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    sid = session.get("sid")
    if sid:
        _df_store.pop(sid, None)
        _output_store.pop(sid, None)
        _progress_queues.pop(sid, None)
    session.clear()
    return redirect(url_for("login"))


# ─── Dashboard ────────────────────────────────────────────────────────────────

@app.route("/dashboard")
@login_required
def dashboard():
    sid = get_sid()
    has_data    = sid in _df_store
    has_output  = sid in _output_store
    return render_template(
        "dashboard.html",
        username=session.get("username", "User"),
        has_data=has_data,
        has_output=has_output,
    )


# ─── Upload ───────────────────────────────────────────────────────────────────

@app.route("/upload", methods=["POST"])
@login_required
def upload():
    sid = get_sid()

    source_type = request.form.get("source_type", "file")  # "file" or "url"

    try:
        if source_type == "url":
            url = request.form.get("github_url", "").strip()
            if not url:
                return jsonify({"success": False, "error": "URL tidak boleh kosong."}), 400
            df = load_and_process_car_sales_data(url)
            source_label = url
        else:
            if "csv_file" not in request.files:
                return jsonify({"success": False, "error": "Tidak ada file yang diunggah."}), 400
            file = request.files["csv_file"]
            if file.filename == "":
                return jsonify({"success": False, "error": "Nama file kosong."}), 400
            if not allowed_file(file.filename):
                return jsonify({"success": False, "error": "Hanya file CSV yang diizinkan."}), 400

            os.makedirs(UPLOAD_DIR, exist_ok=True)
            fname      = secure_filename(file.filename)
            save_path  = os.path.join(UPLOAD_DIR, f"{sid}_{fname}")
            file.save(save_path)
            df = load_and_process_car_sales_data(save_path)
            source_label = file.filename

        _df_store[sid] = df

        # Build preview info
        preview = {
            "source":   source_label,
            "rows":     f"{len(df):,}",
            "columns":  list(df.columns),
            "col_count": len(df.columns),
        }

        # Detect date range if Sale_Date exists
        if "Sale_Date" in df.columns:
            try:
                df["Sale_Date"] = pd.to_datetime(df["Sale_Date"])
                preview["date_range"] = (
                    f"{df['Sale_Date'].min().strftime('%d %b %Y')} — "
                    f"{df['Sale_Date'].max().strftime('%d %b %Y')}"
                )
            except Exception:
                pass

        return jsonify({"success": True, "preview": preview})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


# ─── Generate ─────────────────────────────────────────────────────────────────

@app.route("/generate", methods=["POST"])
@login_required
def generate():
    sid = get_sid()
    if sid not in _df_store:
        return jsonify({"success": False, "error": "Data belum diupload."}), 400

    slide_option   = request.form.get("slide_option", "2")
    gemini_api_key = request.form.get("gemini_api_key", "").strip() or None

    if slide_option in ("1", "4") and not gemini_api_key:
        return jsonify({"success": False, "error": "Gemini API Key wajib diisi untuk opsi ini."}), 400

    # Create a per-session queue for SSE progress
    q = queue.Queue()
    _progress_queues[sid] = q

    df = _df_store[sid]

    def run_generate():
        def cb(msg):
            q.put(msg)

        try:
            result = generate_report(
                df_raw=df,
                slide_option=slide_option,
                output_dir=os.path.join(OUTPUT_DIR, sid),
                gemini_api_key=gemini_api_key,
                progress_cb=cb,
            )
            result["slide_option"] = slide_option
            result["gemini_api_key"] = gemini_api_key
            _output_store[sid] = result
            q.put("__DONE__")
        except Exception as e:
            q.put(f"❌ Error saat generate: {e}")
            q.put("__ERROR__")

    t = threading.Thread(target=run_generate, daemon=True)
    t.start()

    return jsonify({"success": True, "message": "Generate dimulai. Pantau progress melalui /progress."})


@app.route("/progress")
@login_required
def progress():
    """Server-Sent Events endpoint for real-time progress."""
    sid = get_sid()

    def event_stream():
        q = _progress_queues.get(sid)
        if not q:
            yield "data: Tidak ada proses berjalan.\n\n"
            return
        while True:
            try:
                msg = q.get(timeout=60)
                if msg in ("__DONE__", "__ERROR__"):
                    yield f"data: {msg}\n\n"
                    break
                # Replace non-ascii or problematic unicode symbols safely for charmap compatibility
                clean_msg = str(msg).encode("utf-8", "replace").decode("utf-8")
                yield f"data: {clean_msg}\n\n"
            except queue.Empty:
                yield "data: [TIMEOUT] Proses terlalu lama.\n\n"
                break

    return Response(event_stream(), mimetype="text/event-stream")


# ─── Download ─────────────────────────────────────────────────────────────────

@app.route("/download/<filetype>")
@login_required
def download(filetype):
    sid = get_sid()
    output = _output_store.get(sid)
    if not output:
        flash("Belum ada file yang digenerate.", "warning")
        return redirect(url_for("dashboard"))

    if filetype == "pptx":
        path = output.get("pptx")
    elif filetype == "pdf":
        path = output.get("pdf")
    else:
        return "Tipe file tidak valid.", 400

    if not path or not os.path.exists(path):
        flash(f"File {filetype.upper()} tidak tersedia.", "danger")
        return redirect(url_for("dashboard"))

    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


# Store custom uploaded email list per session
_email_list_store: dict[str, list] = {}


@app.route("/download-email-template")
@login_required
def download_email_template():
    if os.path.exists(EMAIL_EXCEL):
        return send_file(EMAIL_EXCEL, as_attachment=True, download_name="template_list_user_email.xlsx")
    return "Template file not found.", 404


@app.route("/upload-email-list", methods=["POST"])
@login_required
def upload_email_list():
    sid = get_sid()
    if "excel_file" not in request.files:
        return jsonify({"success": False, "error": "Tidak ada file Excel yang diunggah."}), 400
    file = request.files["excel_file"]
    if file.filename == "":
        return jsonify({"success": False, "error": "Nama file kosong."}), 400

    if not (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
        return jsonify({"success": False, "error": "Hanya file Excel (.xlsx / .xls) yang diizinkan."}), 400

    try:
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        fname = secure_filename(file.filename)
        save_path = os.path.join(UPLOAD_DIR, f"{sid}_email_{fname}")
        file.save(save_path)
        data = load_email_list_from_excel(save_path)
        if not data:
            return jsonify({"success": False, "error": "Format Excel tidak sesuai atau data kosong."}), 400

        _email_list_store[sid] = data
        return jsonify({"success": True, "data": data, "count": len(data)})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/email-list")
@login_required
def email_list():
    sid = get_sid()
    if sid in _email_list_store:
        data = _email_list_store[sid]
    else:
        data = load_email_list_from_excel(EMAIL_EXCEL)
    return jsonify(data)


# ─── Send Email ───────────────────────────────────────────────────────────────

@app.route("/send-email", methods=["POST"])
@login_required
def send_email():
    sid = get_sid()
    output = _output_store.get(sid)
    if not output:
        return jsonify({"success": False, "error": "Belum ada report yang digenerate."}), 400

    data = request.get_json()
    sender_email    = data.get("sender_email", "").strip()
    sender_password = data.get("sender_password", "").strip()
    selected        = data.get("selected", [])   # [{branch_name, emails:[...]}, ...]

    if not sender_email or not sender_password:
        return jsonify({"success": False, "error": "Email dan App Password wajib diisi."}), 400
    if not selected:
        return jsonify({"success": False, "error": "Pilih minimal satu penerima."}), 400

    # Collect files to send
    file_paths = []
    if output.get("pptx") and os.path.exists(output["pptx"]):
        file_paths.append(output["pptx"])
    if output.get("pdf") and os.path.exists(output["pdf"]):
        file_paths.append(output["pdf"])

    if not file_paths:
        return jsonify({"success": False, "error": "File laporan tidak ditemukan."}), 400

    # Create queue for progress
    q = queue.Queue()
    _progress_queues[sid] = q

    def run_email():
        results = []
        slide_option = output.get("slide_option", "2")
        gemini_api_key = output.get("gemini_api_key")
        df = _df_store.get(sid)

        for item in selected:
            branch = item.get("branch_name", "Unknown")
            emails = item.get("emails", [])
            q.put(f"\n--- Memproses Laporan khusus Cabang: {branch} ---")
            
            # Generate report khusus cabang jika df tersedia
            if df is not None:
                branch_res = generate_report(
                    df_raw=df,
                    slide_option=slide_option,
                    output_dir=os.path.join(OUTPUT_DIR, sid, branch.replace(" ", "_")),
                    gemini_api_key=gemini_api_key,
                    branch_name=branch,
                    progress_cb=lambda msg: q.put(msg)
                )
                branch_files = []
                if branch_res.get("pptx") and os.path.exists(branch_res["pptx"]):
                    branch_files.append(branch_res["pptx"])
                if branch_res.get("pdf") and os.path.exists(branch_res["pdf"]):
                    branch_files.append(branch_res["pdf"])
            else:
                branch_files = file_paths

            ok = send_report_email_multi(
                sender_email=sender_email,
                sender_password=sender_password,
                receiver_emails=emails,
                file_paths=branch_files,
                branch_name=branch,
                progress_cb=lambda msg: q.put(msg),
            )
            results.append({"branch": branch, "success": ok})
        q.put(json.dumps({"__RESULTS__": results}))
        q.put("__DONE__")

    t = threading.Thread(target=run_email, daemon=True)
    t.start()

    return jsonify({"success": True, "message": "Pengiriman email dimulai."})


# ─── Run ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print("=" * 55)
    print("  Report Automation Web App")
    print("  Akses: http://localhost:5000")
    print("  Login: admin / admin123")
    print("=" * 55)
    app.run(debug=True, host="0.0.0.0", port=5000, threaded=True)
