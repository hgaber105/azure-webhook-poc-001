"""
Excel-driven bulk replication engine for the SAP S/4HANA Workforce Person
service.

Extracted from the original DEV_replicate_from_excel.py so the same logic can be
driven either by webhook.py's HTTP upload route (`/replicate_from_excel`) or by
its CLI mode (`python webhook.py --excel <path>`).

Requires: pip install openpyxl
"""

import io
import logging
import time
import uuid

import openpyxl
import requests

import sap_soap

log = logging.getLogger(__name__)

RESULT_COLS = ["PersonWorkAgreement", "Upload Status", "Upload Message"]
REQUIRED_COLS = [sap_soap.COL[k]
                 for k in ("emp", "first", "last", "gender", "cost", "email", "start")]

XLSX_MIMETYPE = ("application/vnd.openxmlformats-officedocument"
                 ".spreadsheetml.sheet")


def _build_row(ws, r, header):
    """Build the SOAP row dict for worksheet row r (same shape DEV built)."""
    def cell(key):
        col = header[sap_soap.COL[key]]
        return sap_soap.s(ws.cell(row=r, column=col).value)

    emp = cell("emp")
    return {
        "emp": emp,
        "first": cell("first"),
        "last": cell("last"),
        "gender": cell("gender") or "1",
        "cost": cell("cost"),
        "email": cell("email"),
        "start": cell("start"),
        "wa_ext": f"{emp}",   # derived; sheet column holds an Excel formula
        "person_uuid": str(uuid.uuid4()),
        "wa_uuid": str(uuid.uuid4()),
    }


def _ensure_result_columns(ws, header):
    next_col = ws.max_column + 1
    for h in RESULT_COLS:
        if h not in header:
            ws.cell(row=1, column=next_col, value=h)
            header[h] = next_col
            next_col += 1


def _process_sheet(ws, *, dry_run=False, max_rows=None, delay=0.0,
                   retries=0, retry_backoff=3.0, request_timeout=60):
    """Replicate every data row in the worksheet; annotate result columns."""
    header = {str(c.value).strip(): i + 1
              for i, c in enumerate(ws[1]) if c.value is not None}

    missing = [h for h in REQUIRED_COLS if h not in header]
    if missing:
        raise ValueError(f"Missing required columns in sheet: {missing}")

    _ensure_result_columns(ws, header)

    session = None if dry_run else requests.Session()
    total = ok = fail = 0
    rows_out = []

    for r in range(2, ws.max_row + 1):
        row = _build_row(ws, r, header)
        if not row["emp"]:
            continue  # skip blank rows
        if max_rows is not None and total >= max_rows:
            break
        total += 1

        payload = sap_soap.build_payload(row)

        if dry_run:
            result = {"status": "DRY_RUN", "proc_code": "", "pwa": "",
                      "bp": "", "message": "not sent (DRY_RUN)"}
        else:
            result = sap_soap.send_soap(payload, session=session, retries=retries,
                                        backoff=retry_backoff,
                                        timeout=request_timeout)
            if delay:
                time.sleep(delay)

        if result["status"] == "SUCCESS":
            ok += 1
            upload_status, upload_message = "Success", ""
        elif result["status"] == "DRY_RUN":
            upload_status, upload_message = "DRY_RUN", ""
        else:
            fail += 1
            upload_status, upload_message = "Failed", result["message"]

        ws.cell(row=r, column=header["PersonWorkAgreement"]).value = result["pwa"]
        ws.cell(row=r, column=header["Upload Status"]).value = upload_status
        ws.cell(row=r, column=header["Upload Message"]).value = upload_message

        log.info("row %s: %s %s %s -> %s PWA=%s BP=%s",
                 r, row["emp"], row["first"], row["last"],
                 result["status"], result["pwa"], result["bp"])
        rows_out.append({
            "row": r, "emp": row["emp"], "first": row["first"], "last": row["last"],
            "status": result["status"], "proc_code": result["proc_code"],
            "pwa": result["pwa"], "bp": result["bp"], "message": result["message"],
        })

    return {"total": total, "ok": ok, "fail": fail, "rows": rows_out}


def process_workbook_file(path, *, dry_run=False, max_rows=None, delay=0.0,
                          retries=0, retry_backoff=3.0, request_timeout=60,
                          fallback_path=None):
    """Process a workbook on disk, writing results back into it.

    If the source is locked (e.g. open in Excel) and a fallback path is given,
    the results are written there instead.
    """
    wb = openpyxl.load_workbook(path)
    summary = _process_sheet(wb.active, dry_run=dry_run, max_rows=max_rows,
                             delay=delay, retries=retries,
                             retry_backoff=retry_backoff,
                             request_timeout=request_timeout)

    out_path = path
    try:
        wb.save(out_path)
    except PermissionError:
        if fallback_path is None:
            root, _, ext = path.rpartition(".")
            fallback_path = (root or path) + "_RESULTS." + (ext or "xlsx")
        out_path = fallback_path
        wb.save(out_path)
        log.warning("source workbook is locked; wrote results to %s", out_path)

    summary["out_path"] = out_path
    return summary


def process_workbook_bytes(data, *, dry_run=False, max_rows=None, delay=0.0,
                           retries=0, retry_backoff=3.0, request_timeout=60):
    """Process an in-memory workbook; return (xlsx_bytes, summary)."""
    wb = openpyxl.load_workbook(io.BytesIO(data))
    summary = _process_sheet(wb.active, dry_run=dry_run, max_rows=max_rows,
                             delay=delay, retries=retries,
                             retry_backoff=retry_backoff,
                             request_timeout=request_timeout)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), summary