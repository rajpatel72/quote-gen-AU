import io
import re
import math
import streamlit as st
from openpyxl import load_workbook

import pdfplumber
from PIL import Image
import pytesseract

# -----------------------------
# Safe value writer
# -----------------------------
def safe_val(v):
    if v is None:
        return ""
    if isinstance(v, (list, dict)):
        return str(v)
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return ""
    try:
        if isinstance(v, (int, float)):
            return v
        s = str(v).strip()
        return float(s) if s.replace(".", "", 1).isdigit() else s
    except Exception:
        return str(v)

# -----------------------------
# Excel writing function
# -----------------------------
def write_usage_to_template(template_bytes, headers, usage_lines_final):
    input_stream = io.BytesIO(template_bytes)
    wb = load_workbook(input_stream)
    ws = wb.active

    # ---- Header section ----
    ws["A6"]  = safe_val(headers.get("Customer Name"))
    ws["B15"] = safe_val(headers.get("Meter Type"))
    ws["B16"] = safe_val(headers.get("Tariff Classification"))
    ws["B17"] = safe_val(headers.get("Distribution Region"))
    ws["B18"] = safe_val(headers.get("Site Address"))
    ws["B19"] = safe_val(headers.get("NMI"))
    ws["B20"] = safe_val(headers.get("Retailer"))

    # ---- Usage table ----
    start_row = 34
    unit_col   = 1  # A
    desc_col   = 3  # C
    before_col = 4  # D
    disc_col   = 5  # E

    for r in range(start_row, 51):
        ws.cell(row=r, column=unit_col,   value="")
        ws.cell(row=r, column=desc_col,   value="")
        ws.cell(row=r, column=before_col, value="")
        ws.cell(row=r, column=disc_col,   value="")

    for i, line in enumerate(usage_lines_final):
        r = start_row + i
        ws.cell(row=r, column=unit_col,   value=safe_val(line.get("Units")))
        ws.cell(row=r, column=desc_col,   value=safe_val(line.get("Description")))
        ws.cell(row=r, column=before_col, value=safe_val(line.get("Rate")))
        ws.cell(row=r, column=disc_col,   value=safe_val(line.get("Discount")))

    output = io.BytesIO()
    wb.save(output)
    return output, "filled_quote.xlsx"

# -----------------------------
# Extraction helpers
# -----------------------------
def find_field(text, label):
    pattern = rf"{label}[:\s]+([^\n]+)"
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""

def extract_usage_table(text):
    """
    Very basic parser: looks for lines like 'Peak 22491.8 0.0632'
    You will need to customize per retailer format!
    """
    usage_lines = []
    for line in text.splitlines():
        if any(k in line.lower() for k in ["peak", "off", "supply", "demand"]):
            parts = line.split()
            if len(parts) >= 3:
                desc = parts[0]
                units = parts[1]
                rate = parts[2]
                discount = parts[3] if len(parts) > 3 else ""
                usage_lines.append({
                    "Description": desc,
                    "Units": units,
                    "Rate": rate,
                    "Discount": discount
                })
    return usage_lines

def extract_from_pdf(path):
    text_data = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text_data.append(page.extract_text() or "")
    full_text = "\n".join(text_data)

    headers = {
        "Customer Name": find_field(full_text, "Customer Name"),
        "Meter Type": find_field(full_text, "Meter Type"),
        "Tariff Classification": find_field(full_text, "Tariff"),
        "Distribution Region": find_field(full_text, "Region"),
        "Site Address": find_field(full_text, "Address"),
        "NMI": find_field(full_text, "NMI"),
        "Retailer": find_field(full_text, "Retailer")
    }

    usage_lines = extract_usage_table(full_text)
    return headers, usage_lines

def extract_from_image(path):
    text = pytesseract.image_to_string(Image.open(path))
    headers = {
        "Customer Name": find_field(text, "Customer Name"),
        "Meter Type": find_field(text, "Meter Type"),
        "Tariff Classification": find_field(text, "Tariff"),
        "Distribution Region": find_field(text, "Region"),
        "Site Address": find_field(text, "Address"),
        "NMI": find_field(text, "NMI"),
        "Retailer": find_field(text, "Retailer")
    }
    usage_lines = extract_usage_table(text)
    return headers, usage_lines

# -----------------------------
# Streamlit UI
# -----------------------------
st.title("Electricity Bill → Quote Generator")

bill_file = st.file_uploader("Upload Bill (PDF/Image)", type=["pdf", "png", "jpg", "jpeg"])
template_file = st.file_uploader("Upload Excel Template", type=["xlsx"])

if bill_file and template_file:
    bill_path = "temp_bill." + bill_file.name.split(".")[-1]
    with open(bill_path, "wb") as f:
        f.write(bill_file.read())

    if bill_file.name.endswith(".pdf"):
        headers, usage_lines_final = extract_from_pdf(bill_path)
    else:
        headers, usage_lines_final = extract_from_image(bill_path)

    st.subheader("Extracted Header Details")
    st.json(headers)

    st.subheader("Extracted Usage Lines")
    st.json(usage_lines_final)

    if st.button("Generate Quote"):
        try:
            template_bytes = template_file.read()
            filled_io, out_name = write_usage_to_template(template_bytes, headers, usage_lines_final)
            st.success("Quote generated successfully!")
            st.download_button("Download Filled Quote", filled_io.getvalue(), file_name=out_name)
        except Exception as e:
            st.error(f"Error: {e}")
