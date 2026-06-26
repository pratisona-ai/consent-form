"""
Streamlit Consent Form Generator
Fetches Dentally consultation notes → Gemini fills template → download .docx
"""

import datetime
import io
import json
import os
import re
import shutil
import tempfile

import requests
import streamlit as st
from bs4 import BeautifulSoup
from docx import Document
from google import genai

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

BASE_URL  = "https://api.dentally.co/v1"
SITE_ID   = "70596be0-e19d-43a1-86e8-e209a3c1aa92"
TEMPLATE  = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "FAME Single Implant Template.docx")

def _secret(key):
    try:
        return st.secrets[key]
    except Exception:
        return os.environ.get(key, "")

TOKEN = _secret("DENTALLY_API_TOKEN")

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/json",
    "User-Agent": "DentallyConsentForm/1.0",
}

# ---------------------------------------------------------------------------
# Dentally helpers
# ---------------------------------------------------------------------------

def search_patient(name):
    r = requests.get(f"{BASE_URL}/patients", headers=HEADERS,
                     params={"query": name}, timeout=15)
    r.raise_for_status()
    return r.json().get("patients", [])


def get_consultation_items(patient_id):
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


def html_to_text(html):
    return BeautifulSoup(html, "html.parser").get_text("\n").strip()

# ---------------------------------------------------------------------------
# Gemini
# ---------------------------------------------------------------------------

CONTENT_SCHEMA = """{
  "clinical_explanation": "2-3 warm sentences in plain English — what brought the patient in and what was found. No abbreviations.",
  "clinical_assessment": "1-2 sentences summarising clinical findings in plain English (fracture, infection, bone volume).",
  "opt_note": "One sentence on what the OPT and CBCT scans showed.",
  "appointment_1": "2-3 sentences: extraction of lower-right wisdom tooth (LR8), implant placement for upper-left premolar (UL4), and assessment of lower-right molar (LR6) for possible immediate implant."
}"""


@st.cache_resource
def _gemini():
    return genai.Client(api_key=_secret("GEMINI_API_KEY"))


def _gemini_call(prompt):
    resp = _gemini().models.generate_content(model="gemini-2.5-flash", contents=prompt)
    return _parse_json(resp.text)


def generate_content(patient_name, notes):
    today  = datetime.date.today().strftime("%d %B %Y")
    prompt = f"""You are a treatment coordinator at FAME Dentistry Ltd, Edinburgh.

Generate patient-friendly text for four sections of a dental implant consent form.

PATIENT: {patient_name}  DATE: {today}  DENTIST: Dr Ferhan Ahmed

CONSULTATION NOTES:
{notes}

Return ONLY valid JSON (no markdown, no code fences):
{CONTENT_SCHEMA}"""
    return _gemini_call(prompt)


def modify_content(patient_name, notes, current, instruction):
    prompt = f"""You are a treatment coordinator at FAME Dentistry Ltd, Edinburgh.

You previously generated this consent form content for {patient_name}:
{json.dumps(current, indent=2)}

The clinician has requested: "{instruction}"

Apply the change and return updated JSON with the same four keys (no markdown, no code fences):
{CONTENT_SCHEMA}"""
    return _gemini_call(prompt)


def _parse_json(text):
    text = re.sub(r'^```(?:json)?\s*|\s*```$', '', text.strip(), flags=re.MULTILINE)
    return json.loads(text)

# ---------------------------------------------------------------------------
# DOCX builder — copies template, replaces text in-place
# ---------------------------------------------------------------------------

def _replace_in_para(para, old, new, exact=False):
    text = para.text
    if exact:
        if text.strip() != old.strip():
            return False
    elif old not in text:
        return False
    for run in para.runs:
        if old in run.text:
            run.text = run.text.replace(old, new, 1)
            return True
    if para.runs:
        para.runs[0].text = text.replace(old, new, 1)
        for run in para.runs[1:]:
            run.text = ""
    return True


def _apply_replacements(doc, sub_map, exact_map):
    def process(para):
        for old, new in exact_map.items():
            if _replace_in_para(para, old, new, exact=True):
                return
        for old, new in sub_map.items():
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


def build_docx_bytes(patient_name, address, content):
    today = datetime.date.today().strftime("%d %B %Y")

    exact_map = {
        "Date":            today,
        "Patient Name":    patient_name,
        "Patient Address": address or "",
        "Patient friendly explanation of clinical notes.":      content["clinical_explanation"],
        "Clinical assessment explained in patient friendly words.": content["clinical_assessment"],
        "OPT with area marked for patient reference.":          content["opt_note"],
        "Appointment explained.":                               content["appointment_1"],
    }

    sub_map = dict(sorted({
        "Dental Implant Treatment Proposal for:  ":
            f"Dental Implant Treatment Proposal for: {patient_name}",
        "Prepared by: Dentist Name": "Prepared by: Dr Ferhan Ahmed",
        "Dear Patient Name,":        f"Dear {patient_name},",
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

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        shutil.copy2(TEMPLATE, tmp.name)
        tmp_path = tmp.name

    try:
        doc = Document(tmp_path)
        _apply_replacements(doc, sub_map, exact_map)
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()
    finally:
        os.unlink(tmp_path)


def docx_to_text(docx_bytes):
    doc = Document(io.BytesIO(docx_bytes))
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

# ---------------------------------------------------------------------------
# UI components
# ---------------------------------------------------------------------------

def render_form_card(patient_name, content, docx_bytes, key):
    today = datetime.date.today().strftime("%d %B %Y")

    st.markdown(f"""
**Patient:** {patient_name} &nbsp;|&nbsp; **Date:** {today} &nbsp;|&nbsp; **Dentist:** Dr Ferhan Ahmed

---

**Your situation**
{content["clinical_explanation"]}

**What we found**
{content["clinical_assessment"]}

**Your scans**
{content["opt_note"]}

**Appointment 1**
{content["appointment_1"]}

---
*Full document includes appointments 2–7, complications, consent & payment sections.*
""")

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="⬇️ Download .docx",
            data=docx_bytes,
            file_name=f"{patient_name} - Implant Consent Form.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key=f"dl_{key}",
        )
    with col2:
        with st.expander("📋 Copy full text"):
            st.text_area(
                label="full_text",
                value=docx_to_text(docx_bytes),
                height=260,
                label_visibility="collapsed",
                key=f"txt_{key}",
            )

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

def _init():
    defaults = dict(
        messages=[],   # list of dicts: {role, text, type?, patient_name?, content?, docx_bytes?}
        stage="init",  # "init" | "form_ready"
        patient=None,
        notes_text=None,
        content=None,
        docx_bytes=None,
        form_version=0,  # increments on each regeneration for unique widget keys
    )
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------

def _handle_patient(patient_name):
    with st.chat_message("assistant"):
        try:
            with st.spinner(f"Searching for {patient_name}…"):
                patients = search_patient(patient_name)
        except Exception as e:
            st.error(f"Dentally error: {e}")
            st.session_state.messages.append({"role": "assistant", "text": f"Error: {e}"})
            return

        if not patients:
            msg = f"No patient found matching **{patient_name}**. Please check the name and try again."
            st.markdown(msg)
            st.session_state.messages.append({"role": "assistant", "text": msg})
            return

        patient   = patients[0]
        full_name = f"{patient['first_name']} {patient['last_name']}"
        address   = ", ".join(filter(None, [
            patient.get("address_line_1"), patient.get("address_line_2"),
            patient.get("town"), patient.get("postcode"),
        ]))
        st.session_state.patient = patient

        st.markdown(f"Found **{full_name}**. Fetching consultation notes…")

        try:
            with st.spinner("Fetching notes from Dentally…"):
                items = get_consultation_items(patient["id"])
        except Exception as e:
            st.error(f"Dentally error: {e}")
            st.session_state.messages.append({"role": "assistant", "text": f"Error: {e}"})
            return

        if not items:
            msg = f"No consultation notes found for **{full_name}**."
            st.markdown(msg)
            st.session_state.messages.append({"role": "assistant", "text": msg})
            return

        notes_text = "\n\n".join(
            f"[{i['nomenclature']}]\n{html_to_text(i.get('notes') or '')}"
            for i in items
        )
        st.session_state.notes_text = notes_text

        st.markdown("Generating consent form with Gemini…")

        try:
            with st.spinner("Generating patient-friendly content…"):
                content = generate_content(full_name, notes_text)
            with st.spinner("Building document…"):
                docx_bytes = build_docx_bytes(full_name, address, content)
        except Exception as e:
            st.error(f"Generation error: {e}")
            st.session_state.messages.append({"role": "assistant", "text": f"Error: {e}"})
            return

        st.session_state.content    = content
        st.session_state.docx_bytes = docx_bytes
        st.session_state.stage      = "form_ready"
        st.session_state.form_version += 1

        st.markdown("Here's the generated consent form:")
        render_form_card(full_name, content, docx_bytes,
                         key=st.session_state.form_version)

        st.session_state.messages.append({
            "role": "assistant",
            "text": f"Found **{full_name}**. Here's the generated consent form:",
            "type": "form",
            "patient_name": full_name,
            "content": content,
            "docx_bytes": docx_bytes,
            "form_key": st.session_state.form_version,
        })


def _handle_modification(instruction):
    patient   = st.session_state.patient
    full_name = f"{patient['first_name']} {patient['last_name']}"
    address   = ", ".join(filter(None, [
        patient.get("address_line_1"), patient.get("address_line_2"),
        patient.get("town"), patient.get("postcode"),
    ]))

    with st.chat_message("assistant"):
        try:
            with st.spinner("Updating consent form…"):
                new_content = modify_content(
                    full_name, st.session_state.notes_text,
                    st.session_state.content, instruction,
                )
                docx_bytes = build_docx_bytes(full_name, address, new_content)
        except Exception as e:
            st.error(f"Error: {e}")
            st.session_state.messages.append({"role": "assistant", "text": f"Error: {e}"})
            return

        st.session_state.content    = new_content
        st.session_state.docx_bytes = docx_bytes
        st.session_state.form_version += 1

        st.markdown("Done — here's the updated consent form:")
        render_form_card(full_name, new_content, docx_bytes,
                         key=st.session_state.form_version)

        st.session_state.messages.append({
            "role": "assistant",
            "text": "Done — here's the updated consent form:",
            "type": "form",
            "patient_name": full_name,
            "content": new_content,
            "docx_bytes": docx_bytes,
            "form_key": st.session_state.form_version,
        })

# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Consent Form Generator", page_icon="🦷", layout="centered")
st.title("🦷 Consent Form Generator")

_init()

if not TOKEN or not _secret("GEMINI_API_KEY"):
    st.error("API keys not configured. Add DENTALLY_API_TOKEN and GEMINI_API_KEY to Streamlit secrets.")
    st.stop()

# Render existing chat history
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["text"])
        if msg.get("type") == "form":
            render_form_card(
                msg["patient_name"], msg["content"], msg["docx_bytes"],
                key=f"hist_{msg['form_key']}",
            )

# Greeting on first load
if not st.session_state.messages:
    with st.chat_message("assistant"):
        st.markdown("Hi! Which patient would you like to generate a consent form for?")
    st.session_state.messages.append({
        "role": "assistant",
        "text": "Hi! Which patient would you like to generate a consent form for?",
    })

# Chat input
if prompt := st.chat_input("Type a message…"):
    with st.chat_message("user"):
        st.markdown(prompt)
    st.session_state.messages.append({"role": "user", "text": prompt})

    if st.session_state.stage == "init":
        _handle_patient(prompt)
    else:
        _handle_modification(prompt)
