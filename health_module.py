import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# ENTERPRISE GMC HEALTH SYSTEM PROMPT (34 Fields)
# ==========================================
HEALTH_SYSTEM_PROMPT = """
You are an expert corporate health insurance underwriter and auditor. Your task is to perform an absolute, Layout-based OCR scan of the provided GMC (Group Medical Cover) document (policy schedule, quote, or draft booklet).

Analyze the entire document carefully. Do not assume or hallucinate any values. If a limit, sublimit, or coverage is not mentioned, strictly mark it as "Not Mentioned" or "Excluded".

*** CRITICAL RULES FOR GMC EXTRACTION ***
1. CO-PAYMENT: Check if there is a co-pay. If no co-pay applies, write "Nil". If it applies to dependents or parents only, state the exact percentage and group (e.g., "10% on dependent parents").
2. ROOM RENT LIMIT: Extract the exact limit. Note if it is a percentage (e.g., "1% of Sum Insured") or a category limit (e.g., "Single Private AC Room" or "No Limit").
3. MATERNITY CLAUSE: Look for normal and Caesarean limits separately. Note if there is a maternity waiting period (e.g., "9 months" or "Day 1 / Waived").
4. PREMIUM RECONCILIATION: Verify that: Net Premium + GST = Total Premium.
5. AUDIT TRAIL ANCHORING: For every extracted value, you must populate the "audit_trail" object. State the exact page and paragraph/table context where you verified the value. This prevents any hallucinations.

Extract and format the output as a clean JSON object matching this schema exactly:
{
  "insurer_name": "Name of insurer, e.g., Star Health, Care Health, Niva Bupa",
  "company_name": "Name of Corporate Insured / Employer Group",
  "employee_count": "Total Covered Employees or Lives (e.g., 250 Lives)",
  "net_premium": 150000.0,
  "gst": 27000.0,
  "gross_premium": 177000.0,
  "sum_insured": "Limit structure, e.g., Family Floater of 3,00,000",
  "family_definition": "e.g., 1+3 (Employee, Spouse, 2 Children) or 1+5",
  "family_members": "List of covered relations, e.g., Employee, Spouse, Children, Dependent Parents",
  "age_limits": "Minimum and maximum age limits for members (e.g., Proposer: 18-65 yrs, Children: 90 days - 25 yrs)",
  "coverages": {
    "pre_existing_disease": "PED waiting period / Day 1 Waiver status",
    "room_rent": "Room rent limit or capping terms",
    "ambulance": "Ambulance capping limit per hospitalization",
    "pre_post_hospitalization": "Pre/post period, e.g., 30 Days Pre, 60 Days Post",
    "maternity_limits": "Maternity limits (e.g., Normal: 50k, C-Sec: 75k)",
    "maternity_waiting_period": "Maternity waiting period waiver status",
    "pre_post_natal": "Are pre/post-natal expenses covered within maternity limit?",
    "ayush": "AYUSH treatment coverage status and sublimits",
    "co_payment": "Co-pay percentage and applicability terms",
    "cataract": "Cataract sublimit capping (e.g., 25,000 per eye)",
    "internal_congenital": "Coverage terms for Internal Congenital diseases",
    "fess": "FESS surgery sublimits or capping",
    "psychiatric": "Psychiatric coverage terms and sublimits",
    "new_born_baby": "NBB Day 1 cover status, limits, and inclusion in floater",
    "modern_treatments": "Modern treatment / advanced procedures capping",
    "lasik": "LASIK treatment coverage threshold (e.g., above 7.5 dioptres) or Excluded",
    "mental_illness": "Mental illness treatment capping terms",
    "day_care": "Day care treatments coverage terms",
    "domiciliary": "Domiciliary hospitalization coverage status and sublimits",
    "organ_donor": "Organ donor harvesting cost limits",
    "well_mother": "Well mother baby care coverage terms",
    "external_congenital": "External Congenital Anomaly coverage terms or Excluded",
    "special_conditions": "List any non-standard special terms, exclusions, or endorsements found"
  },
  "audit_trail": {
    "insured_name_verification": "Page X, Table Y proving corporate name",
    "premium_location": "Page X, Table Y proving premium details",
    "maternity_verification": "Page X, Table Y proving maternity limits",
    "copay_verification": "Page X, Table Y proving co-pay terms"
  }
}
Respond ONLY with raw JSON. Do not write any markdown blocks.
"""

# ==========================================
# EXCEL GENERATOR (Capitup Branded GMC Health QCR)
# ==========================================
def build_capitup_health_excel(baseline, quotes):
    wb = Workbook()
    ws = wb.active
    ws.title = "GMC Comparison"
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
    company_name = str(baseline.get("company_name") or "CORPORATE CLIENT").upper()
    ws.merge_cells(f"A2:{last_col_letter}2")
    ws["A2"] = f"GMC COMPARISON REPORT - {company_name}"
    style_range(f"A2:{last_col_letter}2",
                fill=PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
                font=Font(name=font_family, size=11, bold=True, color="000000"),
                alignment=Alignment(horizontal="center", vertical="center"))
    ws.row_dimensions[2].height = 22

    # Insurers Heading Row
    ws.cell(row=3, column=1, value="Insurers").font = Font(name=font_family, size=11, bold=True)
    ws.cell(row=3, column=1).border = cell_border
    
    base_insurer = str(baseline.get('insurer_name') or 'TATA')
    existing_insurer = f"EXISTING ({base_insurer})"
    ws.cell(row=3, column=2, value=existing_insurer).font = Font(name=font_family, size=11, bold=True)
    ws.cell(row=3, column=2).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=3, column=2).border = cell_border
    
    excel_cols_written = {existing_insurer: True}
    for idx, q in enumerate(quotes):
        col_idx = idx + 3
        q_insurer = str(q.get("insurer_name") or "UNKNOWN").upper()
        col_name = q_insurer
        counter = 1
        while col_name in excel_cols_written:
            col_name = f"{q_insurer} ({counter})"
            counter += 1
            
        excel_cols_written[col_name] = True
        ws.cell(row=3, column=col_idx, value=col_name).font = Font(name=font_family, size=11, bold=True)
        ws.cell(row=3, column=col_idx).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=3, column=col_idx).border = cell_border
    ws.row_dimensions[3].height = 22

    # Map GMC comparison fields
    gmc_fields = [
        ("Corporate Client Name", "company_name"),
        ("Employee Count / Lives", "employee_count"),
        ("Sum Insured Structure", "sum_insured"),
        ("Family Definition", "family_definition"),
        ("Family Members Covered", "family_members"),
        ("Age Limits Structure", "age_limits"),
        ("Net Premium (Basic)", "net_premium"),
        ("GST @ 18%", "gst"),
        ("Gross Premium Payable", "gross_premium"),
        # Coverages
        ("Pre-Existing Disease (PED)", "coverages.pre_existing_disease"),
        ("Room Rent Limits", "coverages.room_rent"),
        ("Ambulance Limits", "coverages.ambulance"),
        ("Pre/Post Hospitalization", "coverages.pre_post_hospitalization"),
        ("Maternity Normal/C-Section", "coverages.maternity_limits"),
        ("Maternity Waiting Period", "coverages.maternity_waiting_period"),
        ("Pre/Post Natal Cover", "coverages.pre_post_natal"),
        ("AYUSH Treatment Limit", "coverages.ayush"),
        ("Co-payment Clauses", "coverages.co_payment"),
        ("Cataract Sublimit", "coverages.cataract"),
        ("Internal Congenital Disease", "coverages.internal_congenital"),
        ("FESS Cover Capping", "coverages.fess"),
        ("Psychiatric Cover", "coverages.psychiatric"),
        ("New Born Baby Cover", "coverages.new_born_baby"),
        ("Modern Treatment Limit", "coverages.modern_treatments"),
        ("LASIK Surgery", "coverages.lasik"),
        ("Mental Illness Cover", "coverages.mental_illness"),
        ("Day Care Treatment", "coverages.day_care"),
        ("Domiciliary Hospitalization", "coverages.domiciliary"),
        ("Organ Donor Cover", "coverages.organ_donor"),
        ("Well Mother Cover", "coverages.well_mother"),
        ("External Congenital Anomaly", "coverages.external_congenital"),
        ("Special Conditions / Terms", "coverages.special_conditions"),
    ]

    current_row = 4
    for label, path in gmc_fields:
        ws.cell(row=current_row, column=1, value=label).font = Font(name=font_family, size=11, bold=True)
        ws.cell(row=current_row, column=1).border = cell_border
        
        # Helper to fetch nested JSON keys securely
        def get_value(record, dotted_path):
            parts = dotted_path.split('.')
            val = record
            for p in parts:
                if isinstance(val, dict):
                    val = val.get(p, "No")
                else:
                    return "No"
            return val

        # Write baseline (existing)
        b_val = get_value(baseline, path)
        cell_base = ws.cell(row=current_row, column=2, value=b_val)
        cell_base.font = Font(name=font_family, size=11)
        cell_base.alignment = Alignment(horizontal="center", vertical="center")
        cell_base.border = cell_border
        
        # Write quotes
        for idx, q in enumerate(quotes):
            col_idx = idx + 3
            q_val = get_value(q, path)
            cell_q = ws.cell(row=current_row, column=col_idx, value=q_val)
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
    return output.getvalue()


# ==========================================
# HEALTH MODULE WORKSPACE RENDERER (Dependency Injected)
# ==========================================
def render_health_vertical(api_keys, process_pdf_document_func, api_key_manager_class):
    st.subheader("🏥 Health Insurance (GMC) Analyzer")
    st.markdown("Upload Health GMC Quotes. The layout parser will extract benefits, waiting periods, and room rent structures side-by-side.")
    
    col_a, col_b = st.columns(2)
    with col_a:
        prev_file = st.file_uploader("1. Previous Year Master Policy", type=["pdf"])
    with col_b:
        new_files = st.file_uploader("2. Current Renewal Quotes", type=["pdf"], accept_multiple_files=True)
    
    if prev_file and new_files:
        if not api_keys:
            st.error("Please configure API keys in the sidebar or `.streamlit/secrets.toml`.")
            return
            
        if st.button("Generate Health GMC QCR Matrix", type="primary", width="stretch"):
            baseline = None
            quotes = []
            
            try:
                key_manager = api_key_manager_class(api_keys)
                
                with st.spinner("Analyzing Previous Policy details..."):
                    prev_bytes = prev_file.read()
                    baseline = process_pdf_document_func(prev_bytes, prev_file.name, HEALTH_SYSTEM_PROMPT, key_manager)

                for q_file in new_files:
                    with st.spinner(f"Extracting GMC details from {q_file.name}..."):
                        q_bytes = q_file.read()
                        record = process_pdf_document_func(q_bytes, q_file.name, HEALTH_SYSTEM_PROMPT, key_manager)
                        quotes.append(record)
                
                if baseline and quotes:
                    st.session_state.health_baseline = baseline
                    st.session_state.health_quotes = quotes
                    st.success("Health GMC analysis complete.")
            
            except Exception as e:
                st.error(f"Processing halted: {e}")

    if "health_baseline" in st.session_state and "health_quotes" in st.session_state:
        b = st.session_state.health_baseline
        qs = st.session_state.health_quotes
        
        # Build UI Dataframe Preview Map with DEDUPLICATION
        preview_cols = {}
        
        # Mapping values cleanly
        gmc_fields_map = [
            ("Corporate Client Name", "company_name"),
            ("Employee Count / Lives", "employee_count"),
            ("Sum Insured Structure", "sum_insured"),
            ("Family Definition", "family_definition"),
            ("Family Members Covered", "family_members"),
            ("Age Limits Structure", "age_limits"),
            ("Net Premium (Basic)", "net_premium"),
            ("GST @ 18%", "gst"),
            ("Gross Premium Payable", "gross_premium"),
            ("Pre-Existing Disease (PED)", "coverages.pre_existing_disease"),
            ("Room Rent Limits", "coverages.room_rent"),
            ("Ambulance Limits", "coverages.ambulance"),
            ("Pre/Post Hospitalization", "coverages.pre_post_hospitalization"),
            ("Maternity Normal/C-Section", "coverages.maternity_limits"),
            ("Maternity Waiting Period", "coverages.maternity_waiting_period"),
            ("Pre/Post Natal Cover", "coverages.pre_post_natal"),
            ("AYUSH Treatment Limit", "coverages.ayush"),
            ("Co-payment Clauses", "coverages.co_payment"),
            ("Cataract Sublimit", "coverages.cataract"),
            ("Internal Congenital Disease", "coverages.internal_congenital"),
            ("FESS Cover Capping", "coverages.fess"),
            ("Psychiatric Cover", "coverages.psychiatric"),
            ("New Born Baby Cover", "coverages.new_born_baby"),
            ("Modern Treatment Limit", "coverages.modern_treatments"),
            ("LASIK Surgery", "coverages.lasik"),
            ("Mental Illness Cover", "coverages.mental_illness"),
            ("Day Care Treatment", "coverages.day_care"),
            ("Domiciliary Hospitalization", "coverages.domiciliary"),
            ("Organ Donor Cover", "coverages.organ_donor"),
            ("Well Mother Cover", "coverages.well_mother"),
            ("External Congenital Anomaly", "coverages.external_congenital"),
            ("Special Conditions / Terms", "coverages.special_conditions"),
        ]
        
        def get_value(record, dotted_path):
            parts = dotted_path.split('.')
            val = record
            for p in parts:
                if isinstance(val, dict):
                    val = val.get(p, "No")
                else:
                    return "No"
            return val

        base_insurer = str(b.get('insurer_name') or 'TATA').upper()
        base_key = f"EXISTING ({base_insurer})"
        preview_cols[base_key] = [get_value(b, path) for label, path in gmc_fields_map]
        
        # Populate Quotes dynamically with dynamic key deduplication
        for q in qs:
            q_insurer = str(q.get("insurer_name") or "UNKNOWN").upper()
            col_name = q_insurer
            counter = 1
            while col_name in preview_cols:
                col_name = f"{q_insurer} ({counter})"
                counter += 1
                
            preview_cols[col_name] = [get_value(q, path) for label, path in gmc_fields_map]
            
        index_labels = [label for label, path in gmc_fields_map]
        df_health_preview = pd.DataFrame(preview_cols, index=index_labels).astype(str)
        
        st.write("---")
        st.subheader("📋 Health GMC QCR Matrix Preview")
        st.dataframe(df_health_preview, width="stretch")
        
        # Generate basic Excel sheet bytes dynamically for the health download
        excel_bytes = build_capitup_health_excel(b, qs)
        
        st.download_button(
            label="📥 Download Health GMC QCR Matrix (Excel)",
            data=excel_bytes,
            file_name=f"Capitup_GMC_QCR_{b.get('company_name', 'GMC_Report')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch"
        )
