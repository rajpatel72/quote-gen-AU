import io
import math
import streamlit as st
from openpyxl import load_workbook

# -----------------------------
# Helper to clean values for Excel
# -----------------------------
def safe_val(v):
    """Convert any value into something Excel can store"""
    if v is None:
        return ""
    if isinstance(v, (list, dict)):
        return str(v)
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return ""  # handle pandas NaN/inf
    try:
        if isinstance(v, (int, float)):
            return v
        s = str(v).strip()
        return float(s) if s.replace(".", "", 1).isdigit() else s
    except Exception:
        return str(v)

# -----------------------------
# Function to fill Excel template
# -----------------------------
def write_usage_to_template(template_bytes, headers, usage_lines_final):
    """
    Fill the Excel template with header details and dynamic usage lines.
    
    Args:
        template_bytes: Excel template as bytes
        headers: dict of header info (Customer Name, NMI, Retailer, etc.)
        usage_lines_final: list of dicts with usage lines
            e.g. [{"Units": "22491.8", "Description": "Peak", "Rate": "0.0632", "Discount": ""}]
    """

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

    # ---- Current Energy Offer table ----
    start_row = 34
    unit_col   = 1  # A
    desc_col   = 3  # C
    before_col = 4  # D
    disc_col   = 5  # E

    # clear out previous values (rows 34-50 for safety)
    for r in range(start_row, 51):
        ws.cell(row=r, column=unit_col,   value="")
        ws.cell(row=r, column=desc_col,   value="")
        ws.cell(row=r, column=before_col, value="")
        ws.cell(row=r, column=disc_col,   value="")

    # fill new usage lines
    for i, line in enumerate(usage_lines_final):
        r = start_row + i

        if not isinstance(line, dict):
            raise ValueError(f"Usage line must be dict, got: {line}")

        ws.cell(row=r, column=unit_col,   value=safe_val(line.get("Units")))
        ws.cell(row=r, column=desc_col,   value=safe_val(line.get("Description")))
        ws.cell(row=r, column=before_col, value=safe_val(line.get("Rate")))
        ws.cell(row=r, column=disc_col,   value=safe_val(line.get("Discount")))

    output = io.BytesIO()
    wb.save(output)
    return output, "filled_quote.xlsx"

# -----------------------------
# Streamlit app
# -----------------------------
st.title("Electricity Quote Generator")

# Upload template
template_file = st.file_uploader("Upload Excel Template", type=["xlsx"])

# Upload extracted bill data (for now JSON or manual entry)
headers = {
    "Customer Name": "John Doe",
    "Meter Type": "Business",
    "Tariff Classification": "Large Customer",
    "Distribution Region": "NSW",
    "Site Address": "123 George St, Sydney",
    "NMI": "NMI123456",
    "Retailer": "Origin Energy"
}

usage_lines_final = [
    {"Units": "22491.80", "Description": "Peak Usage", "Rate": "0.0632", "Discount": ""},
    {"Units": "25764.91", "Description": "Off-Peak Usage", "Rate": "0.0355", "Discount": "5%"},
    {"Units": "1", "Description": "Supply Charge", "Rate": "1.231", "Discount": ""}
]

if template_file:
    template_bytes = template_file.read()

    if st.button("Generate Quote"):
        try:
            filled_io, out_name = write_usage_to_template(template_bytes, headers, usage_lines_final)
            st.success("Quote generated successfully!")
            st.download_button("Download Filled Quote", filled_io.getvalue(), file_name=out_name)
        except Exception as e:
            st.error(f"Error generating quote: {e}")
