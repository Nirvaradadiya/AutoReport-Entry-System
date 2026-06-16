import streamlit as st
import pandas as pd
import os, re, pdfplumber
from datetime import datetime


st.set_page_config(page_title="AutoReport Entry System", layout="centered")
st.title("Automatic extraction of medical report data into Excel")

REPORT_TYPES = {
    "blood": "Blood - CBC", "sugar": "Blood Sugar", "urine": "Urine Routine",
    "lft": "LFT", "kft": "KFT", "lipid": "Lipid Profile", "xray": "X-Ray"
}

if "locked_type" not in st.session_state:
    st.session_state.locked_type = None

def detect_type(text):
    t = text.lower()
    if "hemoglobin" in t: return "blood"
    if "fasting" in t or "hba1c" in t: return "sugar"
    if "bilirubin" in t: return "lft"
    if "creatinine" in t: return "kft"
    if "cholesterol" in t: return "lipid"
    if "urine" in t: return "urine"
    if "x-ray" in t: return "xray"
    return None

def clean_val(v):
    return "-" if not v or str(v).strip().lower()=="nil" or str(v).strip()=="" else str(v).strip()

# dropdown
locked = st.session_state.locked_type is not None
options = list(REPORT_TYPES.values())
default_idx = list(REPORT_TYPES.keys()).index(st.session_state.locked_type) if locked else 0

choice = st.selectbox("Select Report Type", options, index=default_idx, disabled=locked)

col1, col2 = st.columns([1,4])
with col1:
    if not locked:
        if st.button("LOCK TYPE"):
            st.session_state.locked_type = list(REPORT_TYPES.keys())[options.index(choice)]
            st.rerun()
    else:
        if st.button("Change Type"):
            st.session_state.locked_type = None
            st.rerun()
st.markdown("---")
if locked:
    uploaded_files = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True)

    if uploaded_files:
        loading = st.empty()
        loading.info("Loading...")

        t = st.session_state.locked_type
        file_path = f"hospital_reports/{t}_reports.xlsx"
        os.makedirs("hospital_reports", exist_ok=True)
        os.makedirs("uploads", exist_ok=True)

        df_all = pd.read_excel(file_path) if os.path.exists(file_path) else pd.DataFrame()
        warnings = []

        for uploaded in uploaded_files:
            temp_path = f"uploads/{uploaded.name}"
            with open(temp_path, "wb") as f: f.write(uploaded.getbuffer())
            text = ""
            with pdfplumber.open(temp_path) as pdf:
                for p in pdf.pages: text += (p.extract_text() or "") + "\n"

            detected = detect_type(text)
            if detected and detected!= t:
                warnings.append(f"{uploaded.name} → Wrong report type")
                continue

            def find(pat):
                m = re.search(pat, text, re.I)
                return clean_val(m.group(1) if m else "-")

            data = {"Patient Name": find(r"Patient Name[:\s]*([A-Za-z. ]+)"),
                    "Age/Sex": find(r"Age[\/ ]*Sex[:\s]*([\d\/MF]+)"),
                    "UHID": find(r"(?:UHID|MR No\.?)[:\s]*([A-Z0-9]+)"),
                    "Ref Doctor": find(r"Ref(?:erred)?\s*Dr\.?[:\s]*([A-Za-z. ]+)"),
                    "Sample Date": find(r"Sample[^\d]{0,15}([\d\-/]+)"),
                    "Report Date": find(r"Report[^\d]{0,15}([\d\-/]+)"),
                    "Entry Time": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Source File": uploaded.name}

            # type fields (shortened)
            if t=="blood": data.update({"Hemoglobin":find(r"Hemoglobin[^\d]{0,20}([\d.]+)"),"WBC Count":find(r"WBC[^\d]{0,20}([\d,\.]+)")})
            if t=="sugar": data.update({"Fasting":find(r"Fasting[^\d]{0,30}([\d.]+)"),"PP":find(r"PP[^\d]{0,30}([\d.]+)"),"HbA1c":find(r"HbA1c[^\d]{0,10}([\d.]+)")})
            if t=="urine": data.update({"Color":find(r"Color[:\s]*([A-Za-z ]+)"),"Protein":find(r"Protein[^\w]{0,10}([A-Za-z0-9+-]+)")})

            if not df_all.empty and "Source File" in df_all.columns:
                if (df_all["Source File"] == uploaded.name).any(): continue

            df_all = pd.concat([df_all, pd.DataFrame([data])], ignore_index=True)

        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df_all.to_excel(writer, index=False)
            ws = writer.sheets["Reports"] if "Reports" in writer.sheets else writer.sheets[list(writer.sheets.keys())[0]]
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value)) if c.value else 0 for c in col)+2,45)

        loading.empty()
        for w in warnings: st.warning(w)
        with open(file_path,"rb") as f:
            st.download_button("Download Excel File", f, file_name=os.path.basename(file_path))