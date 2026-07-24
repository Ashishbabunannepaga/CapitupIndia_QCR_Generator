import streamlit as st
import json
import time
from google import genai
from google.genai import types
from google.genai.errors import APIError

# ==========================================
# ENTERPRISE API KEY MANAGER
# ==========================================
class APIKeyManager:
    def __init__(self, keys_input):
        if isinstance(keys_input, list):
            self.keys = [str(k).strip() for k in keys_input if str(k).strip()]
        elif isinstance(keys_input, str):
            self.keys = [k.strip() for k in keys_input.split(',') if k.strip()]
        else:
            self.keys = []
        self.current_index = 0
        self.dead_keys = set() 

    def get_next_key(self):
        if not self.keys:
            raise ValueError("No API keys provided.")
        if len(self.dead_keys) == len(self.keys):
            raise Exception("All provided API keys are permanently invalid (401/403).")
        start_index = self.current_index
        while True:
            key = self.keys[self.current_index]
            self.current_index = (self.current_index + 1) % len(self.keys)
            if key not in self.dead_keys:
                return key
            if self.current_index == start_index:
                raise Exception("All keys are permanently dead.")

    def mark_key_dead(self, key, reason):
        masked_key = f"{key[:6]}...{key[-4:]}" if len(key) > 10 else "UNKNOWN_KEY"
        st.sidebar.error(f"💀 Key ({masked_key}) permanently invalid: {reason}.")
        self.dead_keys.add(key)

# ==========================================
# CORE AI PROCESSING & MODEL WATERFALL
# ==========================================
def process_pdf_document(file_bytes, file_name, prompt, key_manager, retries=6):
    model_waterfall = [
        'gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-3.5-flash', 
        'gemini-3.1-pro-preview', 'gemini-3-pro-preview', 
        'gemini-2.5-flash', 'gemini-2.5-flash-lite'
    ]
    json_config = types.GenerateContentConfig(response_mime_type="application/json", temperature=0.0)

    for attempt in range(retries):
        current_model = model_waterfall[attempt % len(model_waterfall)]
        try:
            api_key = key_manager.get_next_key()
            client = genai.Client(api_key=api_key)
            response = client.models.generate_content(
                model=current_model,
                contents=[types.Part.from_bytes(data=file_bytes, mime_type='application/pdf'), prompt], 
                config=json_config
            )
            return json.loads(response.text.strip())
        except APIError as e:
            error_code = getattr(e, "code", None) or getattr(e, "status_code", None)
            if error_code == 429:
                time.sleep(15) 
            elif error_code in [400, 401, 403]:
                key_manager.mark_key_dead(api_key, f"Auth Error ({error_code})")
        except Exception:
            time.sleep(2)
    raise Exception(f"Failed to process {file_name} after retries.")

# ==========================================
# UNDERWRITER MATHEMATICAL AUDIT ENGINE
# ==========================================
class CommercialReconciliationEngine:
    def __init__(self, data_record):
        self.data = data_record
        self.insurer = str(data_record.get("insurer_name") or "UNKNOWN").upper()
        self.gross_premium = float(str(data_record.get("gross_premium") or data_record.get("total_premium") or 0).replace(",", ""))
        self.od_premium = float(str(data_record.get("od_premium") or data_record.get("gross_premium") or 0).replace(",", ""))
        self.tp_premium = float(str(data_record.get("tp_premium") or 0).replace(",", ""))
        self.gst = float(str(data_record.get("gst") or 0).replace(",", ""))
        
    def run_audit(self):
        calc_gross = self.od_premium + self.tp_premium + self.gst
        delta = abs(self.gross_premium - calc_gross)
        logs = [f"ℹ️ Components: OD {self.od_premium:,.2f} | TP {self.tp_premium:,.2f} | GST {self.gst:,.2f}"]
        if delta <= 2.0:
            logs.append(f"✅ Verified: Matches Gross {self.gross_premium:,.2f} perfectly.")
        else:
            logs.append(f"⚠️ Warning: Deviation of ₹{delta:,.2f}. Doc lists {self.gross_premium:,.2f}")
        return logs

def render_underwriter_audit_panel(baseline, quotes):
    st.markdown("---")
    st.subheader("🛡️ Mathematical Reconciliation Logs")
    with st.expander("🔍 EXISTING Policy Audit", expanded=False):
        for log in CommercialReconciliationEngine(baseline).run_audit(): st.markdown(log)
    for q in quotes:
        with st.expander(f"🔍 Quote Audit: {q.get('insurer_name','').upper()}", expanded=False):
            for log in CommercialReconciliationEngine(q).run_audit(): st.markdown(log)
