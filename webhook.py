"""SAP Workforce Person replication webhook.

Ways to replicate a person into SAP S/4HANA:
  * POST /webhook               - accept one employee JSON and replicate it.
  * POST /replicate_from_excel  - accept an uploaded .xlsx, replicate every row,
                                  and return a results workbook (or ?format=json
                                  for a JSON summary).
  * GET  /healthz               - health check.

CLI:
  python webhook.py --excel path/to/book.xlsx   # bulk loader (ex-DEV behavior)
  python webhook.py                             # start the dev server

On Cloud Run this module is served by gunicorn as ``webhook:app``.
"""

import argparse
import io
import json
import logging
import os
import uuid

from flask import Flask, jsonify, request, send_file
import requests

import excel_engine
import sap_soap

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
log = logging.getLogger("webhook")


def _env_bool(name, default=False):
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in ("1", "true", "yes", "on")


def _env_int(name, default):
    try:
        return int(os.getenv(name, default))
    except (TypeError, ValueError):
        return default


def _env_float(name, default):
    try:
        return float(os.getenv(name, default))
    except (TypeError, ValueError):
        return default


# Configuration (environment-driven; previous values kept as defaults) ------- #
PORT = _env_int("PORT", 5000)
FLAG = _env_int("FLAG", 0)                    # 0: SAP (SOAP), 1: Apex (JSON REST)
DRY_RUN = _env_bool("DRY_RUN", False)

SAP_URL = os.getenv("SAP_URL", sap_soap.URL)
SAP_HEADERS = sap_soap.get_headers()
APEX_URL = os.getenv(
    "APEX_URL",
    "https://g3272891305f143-dbb27rv.adb.eu-frankfurt-1.oraclecloudapps.com/ords/orionpoc/employees/",
)
APEX_HEADERS = {
    "Authorization": os.getenv(
        "APEX_AUTH_HEADER", "Basic b3Jpb246TG9uZ25ld3Bhc3N3b3JkMTAwNQ=="),
}

DELAY_BETWEEN_ROWS = _env_float("DELAY_BETWEEN_ROWS", 2.0)
RETRIES_ON_500 = _env_int("RETRIES_ON_500", 3)
RETRY_BACKOFF = _env_float("RETRY_BACKOFF", 3.0)
MAX_ROWS = _env_int("MAX_ROWS", 0) or None
REQUEST_TIMEOUT = _env_int("REQUEST_TIMEOUT", 60)
LOG_TO_FILE = _env_bool("LOG_TO_FILE", True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = _env_int("MAX_UPLOAD_MB", 32) * 1024 * 1024


def _row_from_json(data):
    """Build the same row dict the Excel engine builds, from a JSON payload."""
    emp = sap_soap.s(data.get("person_external_id"))
    return {
        "emp": emp,
        "first": sap_soap.s(data.get("first_name")),
        "last": sap_soap.s(data.get("last_name")),
        "gender": sap_soap.s(data.get("gender_code")) or "1",
        "email": sap_soap.s(data.get("email_address")),
        "cost": sap_soap.s(data.get("cost_center")),
        "start": sap_soap.s(data.get("start_date")),
        "wa_ext": emp,
        "person_uuid": str(uuid.uuid4()),
        "wa_uuid": str(uuid.uuid4()),
    }


def _log_file(name, text):
    """Write a local log file when enabled (LOG_TO_FILE)."""
    if LOG_TO_FILE:
        try:
            with open(name, "w") as f:
                f.write(text)
        except OSError as e:
            log.warning("could not write %s: %s", name, e)


def _send_apex(data):
    try:
        resp = requests.post(APEX_URL, json=data, headers=APEX_HEADERS,
                             timeout=REQUEST_TIMEOUT)
    except requests.RequestException as e:
        return {"status": "FAILED", "message": f"Request error: {e}"}
    try:
        body = resp.json()
    except ValueError:
        body = resp.text
    return {"status": resp.status_code, "body": body}


@app.route("/healthz", methods=["GET"])
def healthz():
    return "ok", 200


@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json(silent=True)
    if not data or "id" not in data:
        return "Missing id", 400

    row = _row_from_json(data)
    payload = sap_soap.build_payload(row)
    _log_file("request_logs.txt",
              f"received: {data}\nsent:\n" + (payload if FLAG == 0 else str(data)))
    log.info("received id=%s -> %s", data.get("id"), "SAP" if FLAG == 0 else "APEX")

    if DRY_RUN:
        result = {"status": "DRY_RUN", "proc_code": "", "pwa": "",
                  "bp": "", "message": "not sent (DRY_RUN)"}
    elif FLAG == 0:
        result = sap_soap.send_soap(payload, retries=RETRIES_ON_500,
                                    backoff=RETRY_BACKOFF, timeout=REQUEST_TIMEOUT)
    else:
        result = _send_apex(data)

    _log_file("response_logs.txt", str(result))
    log.info("result: %s", result)
    return "OK", 200


def _form_bool(value, default):
    if value is None:
        return default
    return value.strip().lower() in ("1", "true", "yes", "on")


def _form_int(value, default):
    try:
        return int(value) if value not in (None, "") else default
    except ValueError:
        return default


def _form_float(value, default):
    try:
        return float(value) if value not in (None, "") else default
    except ValueError:
        return default


@app.route("/replicate_from_excel", methods=["POST"])
def replicate_from_excel():
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify({"error": "no file provided (multipart field 'file')"}), 400

    dry_run = _form_bool(request.form.get("dry_run"), DRY_RUN)
    max_rows = _form_int(request.form.get("max_rows"), MAX_ROWS)
    delay = _form_float(request.form.get("delay"), DELAY_BETWEEN_ROWS)

    try:
        out_bytes, summary = excel_engine.process_workbook_bytes(
            upload.read(), dry_run=dry_run, max_rows=max_rows, delay=delay,
            retries=RETRIES_ON_500, retry_backoff=RETRY_BACKOFF,
            request_timeout=REQUEST_TIMEOUT)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    log.info("excel upload %s: total=%s ok=%s fail=%s",
             upload.filename, summary["total"], summary["ok"], summary["fail"])

    if request.args.get("format") == "json":
        return jsonify(summary)

    name = upload.filename.rsplit(".", 1)[0] + "_RESULTS.xlsx"
    return send_file(io.BytesIO(out_bytes), mimetype=excel_engine.XLSX_MIMETYPE,
                     as_attachment=True, download_name=name)


def _run_cli(argv=None):
    parser = argparse.ArgumentParser(description="SAP workforce replication webhook")
    parser.add_argument("--excel", help="path to an .xlsx to replicate (bulk loader mode)")
    parser.add_argument("--dry-run", action="store_true",
                        help="build payloads but do not send")
    parser.add_argument("--max-rows", type=int, default=None)
    parser.add_argument("--delay", type=float, default=None,
                        help="seconds between rows")
    args = parser.parse_args(argv)

    if args.excel:
        summary = excel_engine.process_workbook_file(
            args.excel,
            dry_run=args.dry_run or DRY_RUN,
            max_rows=args.max_rows if args.max_rows is not None else MAX_ROWS,
            delay=DELAY_BETWEEN_ROWS if args.delay is None else args.delay,
            retries=RETRIES_ON_500, retry_backoff=RETRY_BACKOFF,
            request_timeout=REQUEST_TIMEOUT)
        print(json.dumps({"total": summary["total"], "ok": summary["ok"],
                          "fail": summary["fail"], "out_path": summary["out_path"]}))
        return 0

    app.run(host="0.0.0.0", port=PORT)
    return 0


if __name__ == "__main__":
    raise SystemExit(_run_cli())