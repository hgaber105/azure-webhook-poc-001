"""
Shared SAP S/4HANA Workforce Person replication helpers.

Used by webhook.py (JSON receiver + Excel upload) and by the archived
_DEV_replicate_from_excel.py, so every path builds byte-identical SOAP requests
and parses responses the same way.

Values may be overridden through environment variables (handy on Cloud Run);
the previous hardcoded values are kept as defaults.

Requires: pip install requests
"""

import datetime
import os
import time
import xml.etree.ElementTree as ET

import requests


# ----------------------- Environment configuration ------------------------- #
# ----------------------- Test -----------------------------------------------#
# URL = "https://my434767-api.s4hana.cloud.sap/sap/bc/srt/scs_ext/sap/workforcepersonmasterdatarepli"
# AUTH_HEADER = "Basic SFJfSU5UX1VTRVI6Tlk5TnpAVzlBPldYJiMjRCUvdEJhN29NenFEeWFEeSVQYmlwW3l7ew=="

# ----------------------- Dev -----------------------------------------------#
URL = os.getenv(
    "SAP_URL",
    "https://my434719-api.s4hana.cloud.sap/sap/bc/srt/scs_ext/sap/workforcepersonmasterdatarepli",
)
AUTH_HEADER = os.getenv(
    "SAP_AUTH_HEADER",
    "Basic SU5UX1VTRVI6I3JxZVBlVTJNPldYRmNVeEhZJWtyYlMzL1NNJlNOW3djeXgmKG5QPA==",
)


SOAP_ACTION = os.getenv(
    "SAP_SOAP_ACTION",
    '"http://sap.com/xi/PASEIN/WorkforcePersonMasterDataReplicationRequestResponse_In'
    '/WorkforcePersonMasterDataReplicationRequestResponse_InRequest"',
)

# Defaults for fields not present in the source data ------------------------ #
DEFAULT_COMPANY_CODE = "1000"
DEFAULT_END_DATE = "9999-12-31"
DEFAULT_WEEKLY_HOURS = "40"
DEFAULT_WEEKLY_DAYS = "5"
DEFAULT_WA_TYPE = "1"       # 1/2/3
DEFAULT_WA_STATUS = "1"     # 0=withdrawn, 1=active

# Column headers expected in source data ------------------------------------ #
COL = {
    "emp":    "Emp Num",
    "first":  "First Name",
    "last":   "Last Name",
    "gender": "Gender",
    "cost":   "Cost Center",
    "email":  "Email",
    "start":  "Start Date",
}

def get_headers():
    """SOAP headers built from the current environment configuration."""
    return {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": SOAP_ACTION,
        "Authorization": AUTH_HEADER,
    }


HEADERS = get_headers()


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def s(value):
    """Normalise a value to a clean string ('' for None, dates to YYYY-MM-DD)."""
    if value is None:
        return ""
    if isinstance(value, datetime.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, datetime.date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if len(text) >= 10 and text[4] == "-" and text[7] == "-":
        return text[:10]
    return text


def build_payload(row):
    """row is a dict keyed by the COL logical names (already stringified)."""
    return f"""<soapenv:Envelope
  xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:pas="http://sap.com/xi/PASEIN">
  <soapenv:Header/>
  <soapenv:Body>
    <pas:WorkforcePersonMasterDataReplicationRequest_sync>
      <WorkforcePersonMasterData>
        <PersonUUID>{row['person_uuid']}</PersonUUID>
        <PersonExternalID>{row['emp']}</PersonExternalID>

        <PersonalInformation>
          <FirstName>{row['first']}</FirstName>
          <LastName>{row['last']}</LastName>
          <GenderCode>{row['gender']}</GenderCode>
        </PersonalInformation>

        <WorkPlaceInformation>
          <EmailAddress>{row['email']}</EmailAddress>
        </WorkPlaceInformation>

        <WorkAgreementInformation>
          <PersonWorkAgreementUUID>{row['wa_uuid']}</PersonWorkAgreementUUID>
          <PersonWorkAgreementExternalID>{row['wa_ext']}</PersonWorkAgreementExternalID>
          <PersonWorkAgreementType>{DEFAULT_WA_TYPE}</PersonWorkAgreementType>

          <WorkAgreementJobInformation>
            <ValidityPeriod>
              <StartDate>{row['start']}</StartDate>
              <EndDate>{DEFAULT_END_DATE}</EndDate>
            </ValidityPeriod>
            <WorkAgreementStatus>{DEFAULT_WA_STATUS}</WorkAgreementStatus>
            <CompanyCode>{DEFAULT_COMPANY_CODE}</CompanyCode>
            <CostCenter>{row['cost']}</CostCenter>
            <WorkingTimePercentage>100.00</WorkingTimePercentage>
            <WeeklyWorkingHours>{DEFAULT_WEEKLY_HOURS}</WeeklyWorkingHours>
            <WeeklyWorkDays>{DEFAULT_WEEKLY_DAYS}</WeeklyWorkDays>
          </WorkAgreementJobInformation>
        </WorkAgreementInformation>

      </WorkforcePersonMasterData>
    </pas:WorkforcePersonMasterDataReplicationRequest_sync>
  </soapenv:Body>
</soapenv:Envelope>"""


def _localname(tag):
    return tag.rsplit("}", 1)[-1]


def _findall_local(root, name):
    return [e for e in root.iter() if _localname(e.tag) == name]


def _first_text(root, name):
    for e in root.iter():
        if _localname(e.tag) == name:
            return (e.text or "").strip()
    return ""


def parse_response(status_code, text):
    """Return dict: status, proc_code, pwa, bp, message."""
    out = {"status": "FAILED", "proc_code": "", "pwa": "", "bp": "", "message": ""}
    try:
        root = ET.fromstring(text)
    except ET.ParseError:
        out["message"] = f"HTTP {status_code}; unparseable body: {text[:300]}"
        return out

    # SOAP Fault (transport/schema error, typically HTTP 500)
    fault = _findall_local(root, "Fault")
    if fault:
        out["message"] = f"SOAP Fault: {_first_text(fault[0], 'faultstring')}"
        return out

    out["proc_code"] = _first_text(root, "ProcessingStatusCode")
    out["pwa"] = _first_text(root, "PersonWorkAgreement").strip()

    # Collect log notes; grab business partner id if present
    notes = []
    for li in _findall_local(root, "LogItem"):
        sev = _first_text(li, "SeverityCode")
        note = _first_text(li, "Note")
        if note:
            notes.append(f"[{sev}] {note}")
        if "Business partner" in note:
            digits = "".join(c for c in note if c.isdigit())
            if digits:
                out["bp"] = digits
    out["message"] = " | ".join(notes)

    # ProcessingStatusCode 01 = success, 02 = business error
    if out["proc_code"] == "01":
        out["status"] = "SUCCESS"
    else:
        out["status"] = "FAILED"
    return out


def send_soap(payload, *, url=None, headers=None, session=None,
              retries=0, backoff=3.0, timeout=60):
    """POST a SOAP payload to SAP, returning the parsed result dict.

    The service serializes employee/business-partner creation, so rapid-fire
    calls can collide and return a transient HTTP 500; `retries` (with
    `backoff` seconds multiplied by the attempt) covers that case.
    """
    url = url or URL
    headers = headers or HEADERS
    post = session.post if session is not None else requests.post

    for attempt in range(1, retries + 2):   # 1 initial attempt + retries
        try:
            resp = post(url, data=payload.encode("utf-8"),
                        headers=headers, timeout=timeout)
        except requests.RequestException as e:
            result = {"status": "FAILED", "proc_code": "", "pwa": "",
                      "bp": "", "message": f"Request error: {e}"}
            if attempt <= retries:
                time.sleep(backoff * attempt)
                continue
            return result

        result = parse_response(resp.status_code, resp.text)
        result["message"] = f"HTTP {resp.status_code} | " + result["message"]
        if resp.status_code == 500 and attempt <= retries:
            time.sleep(backoff * attempt)
            continue
        return result

    return {"status": "FAILED", "proc_code": "", "pwa": "",
            "bp": "", "message": "exhausted retries"}