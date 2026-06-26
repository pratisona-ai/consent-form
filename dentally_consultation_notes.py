"""
dentally_consultation_notes.py

Fetches a patient's Surgical Consultation notes from Dentally (via treatment plan
items), uses Gemini to fill the patient-specific sections, then saves a completed
consent form by doing in-place text replacement on a copy of the template —
preserving the logo, procedure images, watermarks, and all formatting.

Usage:
    python dentally_consultation_notes.py "Pratibha Singh"

Auth:
    export DENTALLY_API_TOKEN="..."
    export GEMINI_API_KEY="..."
"""

import datetime
import json
import os
import re
import shutil
import sys

import requests
from bs4 import BeautifulSoup
from docx import Document
from google import genai

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

BASE_URL = "https://api.dentally.co/v1"
SITE_ID = "70596be0-e19d-43a1-86e8-e209a3c1aa92"
TEMPLATE_PATH = os.path.expanduser("~/Desktop/FAME Single Implant Template.docx")
DESKTOP = os.path.expanduser("~/Desktop")

TOKEN = os.environ.get("DENTALLY_API_TOKEN", "")
GEMINI_KEY = os.environ.get("GEMINI_API_KEY", "")

if not TOKEN:
    sys.exit("Error: DENTALLY_API_TOKEN not set.")
if not GEMINI_KEY:
    sys.exit("Error: GEMINI_API_KEY not set.")

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/json",
    "User-Agent": "DentallyConsentForm/1.0",
}


# ---------------------------------------------------------------------------
# Dentally helpers
# ---------------------------------------------------------------------------

def search_patient(name: str) -> list[dict]:
    r = requests.get(f"{BASE_URL}/patients", headers=HEADERS,
                     params={"query": name}, timeout=15)
    r.raise_for_status()
    return r.json().get("patients", [])


def get_consultation_items(patient_id: int) -> list[dict]:
    r = requests.get(f"{BASE_URL}/treatment_plans", headers=HEADERS,
                     params={"patient_id": patient_id, "site_id": SITE_ID}, timeout=15)
    r.raise_for_status()
    plans = r.json().get("treatment_plans", [])

    items = []
    for plan in plans:
        r2 = requests.get(f"{BASE_URL}/treatment_plan_items", headers=HEADERS,
                          params={"treatment_plan_id": plan["id"]}, timeout=15)
        r2.raise_for_status()
        for item in r2.json().get("treatment_plan_items", []):
            if "consultation" in (item.get("nomenclature") or "").lower():
                items.append(item)
    return items


def html_to_text(html: str) -> str:
    return BeautifulSoup(html, "html.parser").get_text("\n").strip()


# ---------------------------------------------------------------------------
# Gemini — returns structured JSON for the variable sections only
# ---------------------------------------------------------------------------

def generate_content(patient_name: str, notes: str) -> dict:
    client = genai.Client(api_key=GEMINI_KEY)
    today = datetime.date.today().strftime("%d %B %Y")

    prompt = f"""You are a treatment coordinator at FAME Dentistry Ltd in Edinburgh.

Using the clinical consultation notes below, write patient-friendly text for four
specific sections of a dental implant consent form.

PATIENT: {patient_name}
DATE: {today}
DENTIST: Dr Ferhan Ahmed

CONSULTATION NOTES:
{notes}

Return ONLY valid JSON — no markdown, no code fences, no commentary.

{{
  "clinical_explanation": "2-3 warm sentences explaining in plain English what brought the patient in and what was found. No clinical abbreviations.",
  "clinical_assessment": "1-2 sentences summarising the clinical assessment findings in plain English (fracture, infection, bone volume).",
  "opt_note": "One sentence describing what the OPT and CBCT scans showed and where the marked areas are on the attached image.",
  "appointment_1": "2-3 sentences describing Appointment 1: surgical extraction of the lower-right wisdom tooth (LR8) and implant placement for the upper-left premolar (UL4), plus assessment of the lower-right molar (LR6) site for possible immediate implant placement if conditions allow."
}}"""

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
    )

    text = re.sub(r'^```(?:json)?\s*|\s*```$', '', response.text.strip(), flags=re.MULTILINE)
    return json.loads(text)


# ---------------------------------------------------------------------------
# In-place paragraph text replacement
# ---------------------------------------------------------------------------

def _replace_in_para(para, old: str, new: str, exact: bool = False) -> bool:
    """
    Replace `old` with `new` in `para`.
    If exact=True, only replace when the full paragraph text equals `old`.
    Preserves run-level formatting where possible; falls back to merging runs.
    """
    text = para.text
    if exact:
        if text.strip() != old.strip():
            return False
        target = old
    else:
        if old not in text:
            return False
        target = old

    # Try single-run replacement first (preserves per-run formatting)
    for run in para.runs:
        if target in run.text:
            run.text = run.text.replace(target, new, 1)
            return True

    # Text is split across runs: consolidate into first run, clear the rest
    if para.runs:
        para.runs[0].text = text.replace(target, new, 1)
        for run in para.runs[1:]:
            run.text = ""
    return True


def _apply_replacements(doc: Document, substring_map: dict, exact_map: dict) -> None:
    """Apply replacements to all paragraphs in body, tables, headers, and footers."""

    def process(para):
        for old, new in exact_map.items():
            if _replace_in_para(para, old, new, exact=True):
                return  # exact match consumed the paragraph
        for old, new in substring_map.items():
            _replace_in_para(para, old, new)

    for para in doc.paragraphs:
        process(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    process(para)

    for section in doc.sections:
        for part in (section.header, section.footer,
                     section.first_page_header, section.first_page_footer):
            try:
                for para in part.paragraphs:
                    process(para)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Build and save
# ---------------------------------------------------------------------------

def save_docx(patient_name: str, address: str, content: dict) -> str:
    today = datetime.date.today().strftime("%d %B %Y")

    # Paragraphs where the ENTIRE text is the placeholder (safe exact match)
    exact_map = {
        "Date":           today,
        "Patient Name":   patient_name,
        "Patient Address": address or "",
        "Patient friendly explanation of clinical notes.":  content["clinical_explanation"],
        "Clinical assessment explained in patient friendly words.": content["clinical_assessment"],
        "OPT with area marked for patient reference.":      content["opt_note"],
        "Appointment explained.":                           content["appointment_1"],
    }

    # Substring replacements — sorted longest-first to prevent partial shadowing
    substring_map = dict(sorted({
        "Dental Implant Treatment Proposal for:  ":
            f"Dental Implant Treatment Proposal for: {patient_name}",
        "Prepared by: Dentist Name":
            "Prepared by: Dr Ferhan Ahmed",
        "Dear Patient Name,":
            f"Dear {patient_name},",
        "I, Patient name, hereby consent":
            f"I, {patient_name}, hereby consent",
        "explained the Treatment Plan to patient name and given her/him":
            f"explained the Treatment Plan to {patient_name} and given her",
        "Total Cost of Appointment 1   Cost":  "Total Cost of Appointment 1   TBC",
        "Total Cost of Appointment 2   Cost":  "Total Cost of Appointment 2   TBC",
        "Total Cost of Appointment 3   Cost":  "Total Cost of Appointment 3   TBC",
        "Total Cost of Appointment 4 Cost ":   "Total Cost of Appointment 4   TBC",
        "Total Cost of Appointment 5   Cost":  "Total Cost of Appointment 5   TBC",
        "Total Cost of Appointment 6    Cost": "Total Cost of Appointment 6   TBC",
        "Total Cost of Appointment 7   Cost":  "Total Cost of Appointment 7   TBC",
    }.items(), key=lambda x: len(x[0]), reverse=True))

    # Copy template — preserves logo, procedure images, watermarks, all formatting
    safe_name = patient_name.replace("/", "-")
    out = os.path.join(DESKTOP, f"{safe_name} - Implant Consent Form.docx")
    shutil.copy2(TEMPLATE_PATH, out)

    doc = Document(out)
    _apply_replacements(doc, substring_map, exact_map)
    doc.save(out)
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run(patient_name: str) -> None:
    print(f"Searching for patient '{patient_name}' ...")
    patients = search_patient(patient_name)
    if not patients:
        sys.exit("No patient found.")
    patient = patients[0]
    patient_id = patient["id"]
    full_name = f"{patient['first_name']} {patient['last_name']}"
    address = ", ".join(filter(None, [
        patient.get("address_line_1"), patient.get("address_line_2"),
        patient.get("town"), patient.get("postcode"),
    ]))
    print(f"  → {full_name} (ID {patient_id})")

    print("Fetching consultation notes ...")
    items = get_consultation_items(patient_id)
    if not items:
        sys.exit("No consultation treatment plan items found.")
    print(f"  → {len(items)} consultation item(s) found")

    notes_text = "\n\n".join(
        f"[{item['nomenclature']}]\n{html_to_text(item.get('notes') or '')}"
        for item in items
    )

    print("Generating patient-friendly content with Gemini ...")
    content = generate_content(full_name, notes_text)

    print("Building consent form from template ...")
    out_path = save_docx(full_name, address, content)
    print(f"\nSaved: {out_path}")
    print("\nGemini filled sections:")
    for k, v in content.items():
        print(f"  [{k}] {v[:100]}...")


if __name__ == "__main__":
    name_input = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else "Pratibha Singh"
    run(name_input)
