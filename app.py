from __future__ import annotations

import io
import re
import uuid
from datetime import datetime
from pathlib import Path

from flask import Flask, flash, make_response, redirect, render_template, request, send_file, url_for
from werkzeug.serving import WSGIRequestHandler

from merge_logs_to_pdf import (
    FdoClockOptions,
    build_combined_lines_with_report,
    build_docx_bytes,
    build_pdf_bytes,
    decode_text_with_fallback,
)


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
CLI_HTML_PATH = BASE_DIR / "static" / "cli" / "txt_log_converter_v20.html"

ALLOWED_TEXT_EXTENSIONS = {".log", ".txt", ".cfg", ".conf"}
ALLOWED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
app.secret_key = "gui-convert-local-secret"
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0
APP_BUILD = datetime.now().strftime("%Y%m%d%H%M%S")
REPORT_CACHE_LIMIT = 100
REPORT_CACHE: dict[str, str] = {}


class QuietRequestHandler(WSGIRequestHandler):
    def log_request(self, code="-", size="-"):
        super().log_request(code, size)


def _is_allowed(filename: str, allowed_extensions: set[str]) -> bool:
    return Path(filename).suffix.lower() in allowed_extensions


def _safe_basename(text: str) -> str:
    text = text.strip()
    if not text:
        return "output"
    sanitized = re.sub(r"[^A-Za-z0-9._-]+", "_", text).strip("._-")
    return sanitized or "output"


def _store_validation_report(report_text: str) -> str:
    report_id = uuid.uuid4().hex
    REPORT_CACHE[report_id] = report_text
    while len(REPORT_CACHE) > REPORT_CACHE_LIMIT:
        oldest = next(iter(REPORT_CACHE))
        REPORT_CACHE.pop(oldest, None)
    return report_id


def _apply_no_cache_headers(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.get("/")
def home():
    return _apply_no_cache_headers(make_response(render_template("home.html", app_build=APP_BUILD)))


@app.get("/cli")
def cli_home():
    if not CLI_HTML_PATH.exists():
        return {"error": "cli file not found"}, 404
    return _apply_no_cache_headers(make_response(send_file(CLI_HTML_PATH, mimetype="text/html")))


@app.get("/gui")
def index():
    return _apply_no_cache_headers(make_response(render_template("index.html", app_build=APP_BUILD)))


@app.post("/generate")
def generate():
    fdo_file = request.files.get("fdo_file")
    apic_file = request.files.get("apic_file")
    image_file = request.files.get("image_file")

    if not fdo_file or not apic_file or not image_file:
        flash("กรุณาอัปโหลดไฟล์ให้ครบทั้ง 3 ไฟล์", "error")
        return redirect(url_for("index"))

    if not _is_allowed(fdo_file.filename, ALLOWED_TEXT_EXTENSIONS):
        flash("ไฟล์ Config/FDO ต้องเป็น .log .txt .cfg หรือ .conf", "error")
        return redirect(url_for("index"))

    if not _is_allowed(apic_file.filename, ALLOWED_TEXT_EXTENSIONS):
        flash("ไฟล์ APIC ต้องเป็น .log .txt .cfg หรือ .conf", "error")
        return redirect(url_for("index"))

    if not _is_allowed(image_file.filename, ALLOWED_IMAGE_EXTENSIONS):
        flash("ไฟล์รูปต้องเป็น .png .jpg .jpeg .bmp .gif หรือ .webp", "error")
        return redirect(url_for("index"))

    try:
        fdo_text = decode_text_with_fallback(fdo_file.read())
        apic_text = decode_text_with_fallback(apic_file.read())
        image_bytes = image_file.read()

        output_base = _safe_basename(Path(fdo_file.filename or "config.log").stem)
        custom_mode = (request.form.get("clock_custom_mode") or "").strip().lower() in {"1", "true", "on", "yes"}
        custom_date = (request.form.get("clock_date") or "").strip() or None
        custom_start = (request.form.get("clock_start") or "08:00:00").strip()
        custom_end = (request.form.get("clock_end") or "18:00:00").strip()
        clock_options = FdoClockOptions(
            custom_mode=custom_mode,
            custom_date=custom_date,
            custom_start_time=custom_start,
            custom_end_time=custom_end,
        )

        combined_lines, validation_report = build_combined_lines_with_report(
            fdo_text,
            apic_text,
            fdo_clock_options=clock_options,
            show_log_title_present=True,
            show_log_image_present=bool(image_bytes),
        )
        combined_text = "\n".join(combined_lines) + "\n"

        output_format = (request.form.get("output_format") or "pdf").strip().lower()
        if output_format not in {"pdf", "docx"}:
            output_format = "pdf"

        txt_name = f"{output_base}.txt"
        (OUTPUT_DIR / txt_name).write_text(combined_text, encoding="utf-8")

        if output_format == "docx":
            output_name = f"{output_base}.docx"
            output_bytes = build_docx_bytes(combined_lines, image_bytes)
            mimetype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            output_name = f"{output_base}.pdf"
            output_bytes = build_pdf_bytes(combined_lines, image_bytes)
            mimetype = "application/pdf"

        (OUTPUT_DIR / output_name).write_bytes(output_bytes)

        stream = io.BytesIO(output_bytes)
        stream.seek(0)
        response = send_file(
            stream,
            as_attachment=True,
            download_name=output_name,
            mimetype=mimetype,
        )
        report_id = _store_validation_report(validation_report)
        response.headers["X-Validation-Report-Id"] = report_id
        return response
    except Exception as exc:
        flash(f"สร้างไฟล์ผลลัพธ์ไม่สำเร็จ: {exc}", "error")
        return redirect(url_for("index"))


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/validation-report/<report_id>")
def validation_report(report_id: str):
    report_text = REPORT_CACHE.get(report_id)
    if report_text is None:
        return {"error": "report not found"}, 404
    response = make_response(report_text)
    response.mimetype = "text/plain"
    response.headers["Cache-Control"] = "no-store"
    return response


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False, request_handler=QuietRequestHandler)
