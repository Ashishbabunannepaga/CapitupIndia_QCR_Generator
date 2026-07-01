import streamlit as st
import tempfile
import os
import json
import io
import pandas as pd
from datetime import datetime

# Import the modern Google GenAI SDK
from google import genai
from google.genai import types

# Import styling components for Excel Generation
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set up page configurations
st.set_page_config(page_title="CapitUp QCR Engine", page_icon="🛡️", layout="wide")

# ==========================================
# NATIVE STREAMLIT SECRETS CONFIGURATION
# ==========================================
default_keys = ""
try:
    # Safe fallback if .streamlit/secrets.toml does not exist yet
    if "GEMINI_API_KEYS" in st.secrets:
        default_keys = st.secrets["GEMINI_API_KEYS"]
except Exception:
    pass

# Sidebar configuration inputs
st.sidebar.title("CapitUp Systems")

global_api_keys = st.sidebar.text_input(
    "API Keys Configuration",
    value=default_keys,
    type="password",
    help="Loaded automatically from .streamlit/secrets.toml if configured."
)

# ==========================================
# MODERN AI CLIENT & KEY ROTATION (google-genai)
# ==========================================
def get_gemini_client(api_keys_str):
    if not api_keys_str:
        return None, "Please provide at least one Gemini API key."
    keys = [k.strip() for k in api_keys_str.split(",") if k.strip()]
    if not keys:
        return None, "No valid API keys found."
    
    if "api_key_index" not in st.session_state:
        st.session_state.api_key_index = 0
        
    current_key = keys[st.session_state.api_key_index % len(keys)]
    
    # Initialize the modern Google GenAI Client
    try:
        client = genai.Client(api_key=current_key)
        return client, None
    except Exception as e:
        return None, f"Failed to initialize GenAI client: {str(e)}"

def rotate_key(api_keys_str):
    if api_keys_str:
        keys = [k.strip() for k in api_keys_str.split(",") if k.strip()]
        if len(keys) > 1:
            st.session_state.api_key_index = (st.session_state.api_key_index + 1) % len(keys)
            st.sidebar.warning(f"Switched key to slot index {st.session_state.api_key_index}")

# ==========================================
# EXCEL GENERATOR (Capitup Branded Motor QCR)
# ==========================================
def build_capitup_motor_excel(baseline, quotes):
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison"
    ws.views.sheetView[0].showGridLines = True
    
    num_columns = len(quotes) + 2
    last_col_letter = get_column_letter(num_columns)
    
    font_family = "Calibri"
    thin_side = Side(style='thin', color='000000')
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    def style_range(cell_range, fill=None, font=None, alignment=None, border=cell_border):
        for row in ws[cell_range]:
            for cell in row:
                if fill: cell.fill = fill
                if font: cell.font = font
                if alignment: cell.alignment = alignment
                if border: cell.border = border

    # Header Row 1: Dark Blue Banner
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = "CAPITUP INDIA PRIVATE LIMITED"
    style_range(f"A1:{last_col_letter}1",
                fill=PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid"),
                font=Font(name=font_family, size=15, bold=True, color="FFFFFF"),
                alignment=Alignment(horizontal="center", vertical="center"))
    ws.row_dimensions[1].height = 28
    
    # Header Row 2: Light Green Banner
    make = baseline.get("make", "VEHICLE").upper()
    ws.merge_cells(f"A2:{last_col_letter}2")
    ws["A2"] = f"COMPARISON - {make}"
    style_range(f"A2:{last_col_letter}2",
                fill=PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
                font=Font(name=font_family, size=11, bold=True, color="000000"),
                alignment=Alignment(horizontal="center", vertical="center"))
    ws.row_dimensions[2].height = 22

    # Metadata rows
    metadata = [
        ("Name of the Insured", baseline.get("insured_name", "N/A")),
        ("Address", baseline.get("address", "N/A")),
        ("Renewal date", baseline.get("renewal_date", "N/A")),
        ("Registration Number", baseline.get("registration_number", "N/A")),
        ("Make", baseline.get("make", "N/A")),
        ("Model", baseline.get("model", "N/A")),
        ("Seating Capacity", baseline.get("seating_capacity", "N/A")),
        ("Cubic Capacity", baseline.get("cubic_capacity", "N/A")),
        ("Year of Manufacture", baseline.get("year_of_manufacture", "N/A")),
    ]
    
    current_row = 3
    for label, val in metadata:
        ws.cell(row=current_row, column=1, value=label)
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=num_columns)
        ws.cell(row=current_row, column=2, value=val)
        
        ws.cell(row=current_row, column=1).font = Font(name=font_family, size=11)
        ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
        
        style_range(f"B{current_row}:{last_col_letter}{current_row}",
                    font=Font(name=font_family, size=11, bold=True),
                    alignment=Alignment(horizontal="center", vertical="center"))
        ws.cell(row=current_row, column=1).border = cell_border
        ws.row_dimensions[current_row].height = 20
        current_row += 1

    # Insurers Heading Row
    ws.cell(row=current_row, column=1, value="Insurers").font = Font(name=font_family, size=11, bold=True)
    ws.cell(row=current_row, column=1).border = cell_border
    
    existing_insurer = f"EXISTING ({baseline.get('insurer_name', 'TATA')})"
    ws.cell(row=current_row, column=2, value=existing_insurer).font = Font(name=font_family, size=11, bold=True)
    ws.cell(row=current_row, column=2).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=current_row, column=2).border = cell_border
    
    for idx, q in enumerate(quotes):
        col_idx = idx + 3
        ws.cell(row=current_row, column=col_idx, value=q.get("insurer_name", "UNKNOWN").upper()).font = Font(name=font_family, size=11, bold=True)
        ws.cell(row=current_row, column=col_idx).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=current_row, column=col_idx).border = cell_border
    ws.row_dimensions[current_row].height = 22
    current_row += 1

    # Commercial values mapping
    commercials = [
        ("IDV", "idv", "#,##,##0"),
        ("NCB in %", "ncb_percent", None),
        ("TP", "tp_premium", "#,##,##0"),
        ("OD Premium", "od_premium", "#,##,##0"),
        ("Gst @ 18%", "gst", "#,##,##0"),
        ("Gross Premium", "gross_premium", "#,##,##0"),
    ]
    
    for label, json_key, num_fmt in commercials:
        is_gross = (label == "Gross Premium")
        ws.cell(row=current_row, column=1, value=label).font = Font(name=font_family, size=11, bold=is_gross)
        ws.cell(row=current_row, column=1).border = cell_border
        
        base_val = baseline.get(json_key, 0)
        cell_base = ws.cell(row=current_row, column=2, value=base_val)
        cell_base.font = Font(name=font_family, size=11, bold=is_gross)
        cell_base.alignment = Alignment(horizontal="center", vertical="center")
        cell_base.border = cell_border
        if num_fmt and isinstance(base_val, (int, float)):
            cell_base.number_format = num_fmt
            
        for idx, q in enumerate(quotes):
            col_idx = idx + 3
            q_val = q.get(json_key, 0)
            cell_q = ws.cell(row=current_row, column=col_idx, value=q_val)
            cell_q.font = Font(name=font_family, size=11, bold=is_gross)
            cell_q.alignment = Alignment(horizontal="center", vertical="center")
            cell_q.border = cell_border
            if num_fmt and isinstance(q_val, (int, float)):
                cell_q.number_format = num_fmt
                
        ws.row_dimensions[current_row].height = 20
        current_row += 1

    # Coverages Header row
    ws.merge_cells(f"A{current_row}:{last_col_letter}{current_row}")
    ws.cell(row=current_row, column=1, value="Coverages:").font = Font(name=font_family, size=11, bold=True)
    style_range(f"A{current_row}:{last_col_letter}{current_row}",
                font=Font(name=font_family, size=11, bold=True),
                alignment=Alignment(horizontal="left", vertical="center"))
    ws.row_dimensions[current_row].height = 20
    current_row += 1

    # Dynamic Discovery of Coverage Keys
    all_coverage_keys = set()
    all_coverage_keys.update(baseline.get("coverages", {}).keys())
    for q in quotes:
        all_coverage_keys.update(q.get("coverages", {}).keys())
        
    standard_order = [
        "Zero Depreciation", "Engine Protection", "Return to Invoice", "Tyre Protection",
        "Consumables", "Roadside Assistance", "Key Replacement", "Passenger Assistance",
        "Loss of Personal Belongings", "Basic TP", "PA Cover", "Legal Liability to Paid Driver"
    ]
    
    ordered_keys = [k for k in standard_order if k in all_coverage_keys]
    dynamic_keys = sorted([k for k in all_coverage_keys if k not in standard_order])
    final_coverage_rows = ordered_keys + dynamic_keys
    
    for key in final_coverage_rows:
        ws.cell(row=current_row, column=1, value=key).font = Font(name=font_family, size=11)
        ws.cell(row=current_row, column=1).border = cell_border
        
        base_cov = baseline.get("coverages", {}).get(key, "No")
        cell_base = ws.cell(row=current_row, column=2, value=base_cov)
        cell_base.font = Font(name=font_family, size=11)
        cell_base.alignment = Alignment(horizontal="center", vertical="center")
        cell_base.border = cell_border
        
        for idx, q in enumerate(quotes):
            col_idx = idx + 3
            q_cov = q.get("coverages", {}).get(key, "No")
            cell_q = ws.cell(row=current_row, column=col_idx, value=q_cov)
            cell_q.font = Font(name=font_family, size=11)
            cell_q.alignment = Alignment(horizontal="center", vertical="center")
            cell_q.border = cell_border
            
        ws.row_dimensions[current_row].height = 20
        current_row += 1

    ws.column_dimensions['A'].width = 30
    for col in range(2, num_columns + 1):
        ws.column_dimensions[get_column_letter(col)].width = 20
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue(), final_coverage_rows


# ==========================================
# MASTER LAYOUT DEFINITION & Vertical ROUTER
# ==========================================
def main():
    domain = st.sidebar.selectbox(
        "Select Insurance Domain", 
        ["🚗 Motor QCR Generator", "🏥 Health Insurance (GMC)"]
    )
    
    st.sidebar.write("---")
    
    if domain == "🚗 Motor QCR Generator":
        render_motor_vertical(global_api_keys)
    elif domain == "🏥 Health Insurance (GMC)":
        render_health_vertical(global_api_keys)


# ==========================================
# MOTOR INSURANCE MODULE
# ==========================================
MOTOR_SYSTEM_PROMPT = """
You are an expert insurance processing system analyzing an insurance schedule or quotation.
Analyze the provided PDF document and extract technical specifications, premium details, and coverage options.

Extract and format the output as a clean JSON object matching this schema:
{
  "insurer_name": "Name of insurer, e.g., TATA AIG, ICICI LOMBARD, LIBERTY",
  "insured_name": "Name of Insured / Policyholder",
  "address": "State or Location of Address",
  "renewal_date": "Policy expiration / renewal date in DD-MM-YYYY format",
  "registration_number": "Vehicle Registration Number",
  "make": "Vehicle Make (e.g. TOYOTA)",
  "model": "Vehicle Model & Variant (e.g. INNOVA HYCROSS)",
  "seating_capacity": 5,
  "cubic_capacity": 1987,
  "year_of_manufacture": 2024,
  "idv": 2400000,
  "ncb_percent": "20%",
  "tp_premium": 0.0,
  "od_premium": 29362.0,
  "gst": 5285.0,
  "gross_premium": 34647.0,
  "coverages": {
    "Zero Depreciation": "Yes or No",
    "Engine Protection": "Yes or No",
    "Return to Invoice": "Yes or No",
    "Tyre Protection": "Yes or No",
    "Consumables": "Yes or No",
    "Roadside Assistance": "Yes or No",
    "Key Replacement": "Yes or No",
    "Passenger Assistance": "Yes or No",
    "Loss of Personal Belongings": "Yes or No",
    "Basic TP": "Yes or No",
    "PA Cover": "Yes or No",
    "Legal Liability to Paid Driver": "Yes or No"
  }
}

Use standard normalizations for coverage keys:
- Depreciation Reimbursement/ Nil Depreciation/ Zero Dep -> "Zero Depreciation"
- Engine Protect/ Engine Secure/ Engine Safe -> "Engine Protection"
- Loss of Personal Belongings -> "Loss of Personal Belongings"
- Passenger Assistance/ Emergency Assistance -> "Passenger Assistance"
- Key Replacement / Key Loss -> "Key Replacement"
- Consumable Expenses / Consumables -> "Consumables"
- Road side Assistance / RSA / IL Smart Assist -> "Roadside Assistance"
- Gap Value / Return to Invoice -> "Return to Invoice"
- Tyre Secure / Tyre Protection -> "Tyre Protection"

If any other unique add-on cover (e.g. "Liberty Assure", "Smart Saver", "Emergency Medical Expenses") is explicitly active, add it to the coverages dictionary under its clean name as "Yes".

Ensure you return ONLY valid JSON.
"""

def render_motor_vertical(api_keys):
    st.subheader("🚗 Motor Insurance QCR Generation Module")
    st.markdown("Process previous year schedules and new comparative quotes directly using Layout OCR.")
    
    col_a, col_b = st.columns(2)
    with col_a:
        prev_file = st.file_uploader("1. Previous Year Policy Copy", type=["pdf"])
    with col_b:
        new_files = st.file_uploader("2. Current Quotations", type=["pdf"], accept_multiple_files=True)
        
    if prev_file and new_files:
        client, err = get_gemini_client(api_keys)
        if err:
            st.sidebar.warning(err)
            return
            
        if st.button("Generate Motor QCR Report", width="stretch"):
            baseline = None
            quotes = []
            
            with st.spinner("Analyzing Previous Policy details..."):
                try:
                    prev_bytes = prev_file.read()
                    response = client.models.generate_content(
                        model=selected_model,
                        contents=[
                            types.Part.from_bytes(data=prev_bytes, mime_type='application/pdf'),
                            MOTOR_SYSTEM_PROMPT
                        ],
                        config=types.GenerateContentConfig(response_mime_type="application/json")
                    )
                    baseline = json.loads(response.text.strip())
                except Exception as e:
                    rotate_key(api_keys)
                    st.error(f"Failed parsing previous policy: {e}")
                    return

            for q_file in new_files:
                with st.spinner(f"Analyzing quote: {q_file.name}..."):
                    try:
                        q_bytes = q_file.read()
                        response = client.models.generate_content(
                            model=selected_model,
                            contents=[
                                types.Part.from_bytes(data=q_bytes, mime_type='application/pdf'),
                                MOTOR_SYSTEM_PROMPT
                            ],
                            config=types.GenerateContentConfig(response_mime_type="application/json")
                        )
                        quotes.append(json.loads(response.text.strip()))
                    except Exception as e:
                        rotate_key(api_keys)
                        st.error(f"Failed parsing quote {q_file.name}: {e}")
                        
            if baseline and quotes:
                st.session_state.motor_baseline = baseline
                st.session_state.motor_quotes = quotes
                st.success("Comparison extraction completed successfully.")

    # Render Preview and Excel Generation
    if "motor_baseline" in st.session_state and "motor_quotes" in st.session_state:
        b = st.session_state.motor_baseline
        qs = st.session_state.motor_quotes
        
        excel_bytes, dynamic_covers = build_capitup_motor_excel(b, qs)
        
        # Build UI Dataframe Preview Map
        preview_cols = {f"EXISTING ({b.get('insurer_name', 'TATA')})": []}
        for q in qs:
            preview_cols[q.get("insurer_name", "UNKNOWN").upper()] = []
            
        row_labels = [
            "Name of the Insured", "Address", "Renewal date", "Registration Number", 
            "Make", "Model", "Seating Capacity", "Cubic Capacity", "Year of Manufacture",
            "--- COMMERCIALS ---", "IDV", "NCB in %", "TP", "OD Premium", "Gst @ 18%", "Gross Premium",
            "--- COVERAGES ---"
        ] + dynamic_covers
        
        preview_cols[list(preview_cols.keys())[0]].extend([
            b.get("insured_name"), b.get("address"), b.get("renewal_date"), b.get("registration_number"),
            b.get("make"), b.get("model"), b.get("seating_capacity"), b.get("cubic_capacity"), b.get("year_of_manufacture"),
            "---", b.get("idv"), b.get("ncb_percent"), b.get("tp_premium"), b.get("od_premium"), b.get("gst"), b.get("gross_premium"),
            "---"
        ] + [b.get("coverages", {}).get(key, "No") for key in dynamic_covers])
        
        for q in qs:
            preview_cols[q.get("insurer_name").upper()].extend([
                b.get("insured_name"), b.get("address"), b.get("renewal_date"), b.get("registration_number"),
                b.get("make"), b.get("model"), b.get("seating_capacity"), b.get("cubic_capacity"), b.get("year_of_manufacture"),
                "---", q.get("idv"), q.get("ncb_percent"), q.get("tp_premium"), q.get("od_premium"), q.get("gst"), q.get("gross_premium"),
                "---"
            ] + [q.get("coverages", {}).get(key, "No") for key in dynamic_covers])
            
        # FIX: Enforce string formatting on all preview cells to prevent PyArrow conversion crashes!
        df_preview = pd.DataFrame(preview_cols, index=row_labels).astype(str)
        
        st.markdown("---")
        st.subheader("📋 Capitup Comparison Preview")
        st.dataframe(df_preview, width="stretch")
        
        st.download_button(
            label="📥 Download Capitup Comparative Sheet (Excel)",
            data=excel_bytes,
            file_name=f"Capitup_Motor_QCR_{b.get('registration_number', 'Report')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch"
        )


# ==========================================
# HEALTH INSURANCE MODULE (GMC QCR)
# ==========================================
HEALTH_SYSTEM_PROMPT = """
You are an expert health insurance audit system. Extract all GMC (Group Medical Cover) quote details 
precisely and format the output as a clean JSON matching the following schema.

Extract and format the output as a clean JSON object:
{
  "insurer_name": "Name of insurer, e.g., Star Health, Care Health, Niva Bupa",
  "insured_name": "Name of Corporate Insured / Group",
  "policy_type": "Group Medical Cover (GMC)",
  "total_lives": 250,
  "sum_insured": "Family Floater of 3,00,000",
  "gross_premium": 150000.0,
  "gst": 27000.0,
  "total_premium": 177000.0,
  "coverages": {
    "Room Rent Limits": "1% of Sum Insured or Single Private AC Room",
    "ICU Limit": "2% of Sum Insured or No Limit",
    "Pre-Existing Disease waiting period": "Covered from Day 1 / Waiver",
    "Maternity Limit": "50,000 for Normal, 75,000 for Caesarean",
    "Corporate Buffer": "No / Yes (e.g. 5,00,000)",
    "Co-payment": "No Co-pay / 10% on claims"
  }
}
Respond ONLY with raw JSON. No markdown wrappers.
"""

def render_health_vertical(api_keys):
    st.subheader("🏥 Health Insurance (GMC) Analyzer")
    st.markdown("Upload Health GMC Quotes. The layout parser will extract benefits, waiting periods, and room rent structures side-by-side.")
    
    uploaded_quotes = st.file_uploader("Upload GMC Quotes (PDF)", type=["pdf"], accept_multiple_files=True)
    
    if uploaded_quotes:
        client, err = get_gemini_client(api_keys)
        if err:
            st.sidebar.warning(err)
            return
            
        if st.button("Generate Health GMC QCR Matrix", width="stretch"):
            health_records = []
            
            for q_file in uploaded_quotes:
                with st.spinner(f"Extracting GMC details from {q_file.name}..."):
                    try:
                        q_bytes = q_file.read()
                        response = client.models.generate_content(
                            model=selected_model,
                            contents=[
                                types.Part.from_bytes(data=q_bytes, mime_type='application/pdf'),
                                HEALTH_SYSTEM_PROMPT
                            ],
                            config=types.GenerateContentConfig(response_mime_type="application/json")
                        )
                        health_records.append(json.loads(response.text.strip()))
                    except Exception as e:
                        rotate_key(api_keys)
                        st.error(f"Failed parsing GMC Quote {q_file.name}: {e}")
            
            if health_records:
                st.session_state.health_records = health_records
                st.success("Health GMC analysis complete.")

    if "health_records" in st.session_state and st.session_state.health_records:
        h_records = st.session_state.health_records
        
        # Format a clean side-by-side comparative matrix preview
        columns_map = {}
        for idx, rec in enumerate(h_records):
            insurer = rec.get("insurer_name", f"Insurer {idx+1}").upper()
            columns_map[insurer] = [
                rec.get("insured_name"),
                rec.get("policy_type"),
                rec.get("total_lives"),
                rec.get("sum_insured"),
                rec.get("gross_premium"),
                rec.get("gst"),
                rec.get("total_premium"),
                rec.get("coverages", {}).get("Room Rent Limits", "No"),
                rec.get("coverages", {}).get("ICU Limit", "No"),
                rec.get("coverages", {}).get("Pre-Existing Disease waiting period", "No"),
                rec.get("coverages", {}).get("Maternity Limit", "No"),
                rec.get("coverages", {}).get("Corporate Buffer", "No"),
                rec.get("coverages", {}).get("Co-payment", "No")
            ]
            
        index_labels = [
            "Corporate / Insured Group", "Policy Type", "Total Covered Lives", "Sum Insured Structure",
            "Gross Premium (Basic)", "GST @ 18%", "Total Premium Payable",
            "Room Rent Limits", "ICU Limits", "Pre-Existing Disease Waiver", "Maternity Limits",
            "Corporate Buffer", "Co-payment Clauses"
        ]
        
        # Enforce clean formatting & string cast to completely prevent pyarrow errors
        df_health_preview = pd.DataFrame(columns_map, index=index_labels).astype(str)
        
        st.write("---")
        st.subheader("📋 Health GMC QCR Matrix Preview")
        st.dataframe(df_health_preview, width="stretch")
        
        # Generate basic Excel sheet bytes dynamically for the health download
        wb = Workbook()
        ws = wb.active
        ws.title = "Health GMC Comparison"
        
        ws.cell(row=1, column=1, value="Health GMC Comparison").font = Font(name="Calibri", size=14, bold=True)
        for r_idx, label in enumerate(index_labels, start=3):
            ws.cell(row=r_idx, column=1, value=label).font = Font(name="Calibri", bold=True)
            for c_idx, insurer in enumerate(columns_map.keys(), start=2):
                ws.cell(row=2, column=c_idx, value=insurer).font = Font(name="Calibri", bold=True)
                ws.cell(row=r_idx, column=c_idx, value=columns_map[insurer][r_idx-3])
                
        output_health = io.BytesIO()
        wb.save(output_health)
        
        st.download_button(
            label="📥 Download Health GMC QCR Matrix (Excel)",
            data=output_health.getvalue(),
            file_name="GMC_Comparison_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch"
        )


if __name__ == "__main__":
    # Choose Pro model as default for high-precision table rendering & calculations
    selected_model = "gemini-2.5-flash"
    main()