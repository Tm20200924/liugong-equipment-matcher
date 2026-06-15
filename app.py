# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import os, sys, json, io, tempfile
from datetime import datetime

st.set_page_config(page_title="LiuGong Matcher v19", page_icon="🔧", layout="wide")

# ==================== DATA LOADING (Google Drive + Local Fallback) ====================

@st.cache_resource
def load_engine():
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), "engine"))
    import engine_v19 as eng
    return eng

@st.cache_data(ttl=3600)
def load_data():
    base = os.path.dirname(__file__)
    eng = load_engine()
    os.makedirs(os.path.join(base, "engine"), exist_ok=True)

    # Try Google Drive download (from Streamlit secrets)
    gdrive = st.secrets.get("gdrive", {})
    
    config_path = os.path.join(base, "engine", "config.json")
    db_path = os.path.join(base, "engine", "product_db_full.json")
    comp_path = os.path.join(base, "engine", "competitor_db.json")

    # Download from Google Drive if URLs provided and files missing
    for key, path in [("product_db_url", db_path), ("competitor_db_url", comp_path)]:
        url = gdrive.get(key)
        if url and not os.path.exists(path):
            try:
                import gdown
                gdown.download(url, path, quiet=True)
            except Exception:
                pass

    # Set paths for engine
    eng.CONFIG_PATH = config_path
    eng.DB_PATH = db_path
    eng.COMP_DB_PATH = comp_path

    config = eng.load_config()
    db = eng.load_db()

    # Override config from Streamlit secrets if available
    if "config" in st.secrets:
        for k, v in st.secrets["config"].items():
            config[k] = v

    return config, db

try:
    eng = load_engine()
    config, db = load_data()
    data_loaded = True
except Exception as e:
    data_loaded = False
    load_error = str(e)

# ==================== SIDEBAR ====================

with st.sidebar:
    st.header("⚙️ Settings")
    
    if not data_loaded:
        st.error(f"Data loading failed: {load_error}")
        st.info("Configure Google Drive URLs in .streamlit/secrets.toml")
        st.stop()
    
    st.subheader("Exchange Rate & Tax")
    rate = st.number_input("Rate (CNY→RUB)", value=config.get("exchange_rate_rub_cny", 11.5), step=0.1, format="%.1f")
    duty = st.number_input("Customs Duty (%)", value=config.get("customs_duty_rate_pct", 5.0), step=0.1, format="%.1f")
    proc = st.number_input("Processing Fee (%)", value=config.get("customs_processing_rate_pct", 0.3), step=0.1, format="%.1f")
    vat = st.number_input("VAT (%)", value=config.get("vat_rate_pct", 22.0), step=0.1, format="%.1f")
    warehouse = st.number_input("Warehouse (RUB)", value=config.get("customs_warehouse_fee_rub", 20000), step=1000)
    customs_fee = st.number_input("Customs Fee (RUB)", value=config.get("customs_fee_rub", 6000), step=1000)
    agent_fee = st.number_input("Agent Fee (RUB)", value=config.get("customs_agent_fee_rub", 30000), step=1000)
    
    st.divider()
    st.caption(f"Products: {len(db)} models | v19")

# Apply sidebar values
config["exchange_rate_rub_cny"] = rate
config["customs_duty_rate_pct"] = duty
config["customs_processing_rate_pct"] = proc
config["vat_rate_pct"] = vat
config["customs_warehouse_fee_rub"] = warehouse
config["customs_fee_rub"] = customs_fee
config["customs_agent_fee_rub"] = agent_fee

# ==================== HEADER ====================

st.title("🔧 LiuGong Equipment Matcher v19")
st.caption("Upload inquiry → Auto-match → DAP Price → Verify")

# ==================== TAB 1: INPUT ====================

tab1, tab2, tab3 = st.tabs(["📥 Input", "📊 Results", "✅ Verify"])

with tab1:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Upload File")
        uploaded = st.file_uploader("Drop inquiry Excel/CSV", type=["xlsx", "xls", "csv"])
    
    with col2:
        st.subheader("Or Paste Text")
        manual_text = st.text_area("One inquiry per line", height=150,
                                   placeholder="Excavator 20t x2
Roller 3t x6
Bulldozer TY230 x4")
    
    if st.button("🚀 Match Now", type="primary", use_container_width=True):
        inquiries = []
        
        if uploaded:
            ext = uploaded.name.rsplit(".", 1)[-1].lower()
            if ext in ("xlsx", "xls"):
                inquiries = eng.parse_inquiry_excel(uploaded)
            elif ext == "csv":
                import io as _io
                df = pd.read_csv(_io.BytesIO(uploaded.getvalue()))
                for i, row in df.iterrows():
                    name = str(row.iloc[1]) if len(row) > 1 else str(row.iloc[0])
                    qty = int(row.iloc[2]) if len(row) > 2 and str(row.iloc[2]).isdigit() else 1
                    inquiries.append({"seq": str(i+1), "name": name, "qty": qty})
            st.success(f"Parsed {len(inquiries)} inquiries from file")
        
        if manual_text.strip():
            lines = [l.strip() for l in manual_text.strip().split("
") if l.strip()]
            for i, line in enumerate(lines):
                import re
                qty = 1
                qm = re.search(r"x\s*(\d+)", line, re.IGNORECASE)
                if qm: qty = int(qm.group(1))
                inquiries.append({"seq": str(i+1), "name": line, "qty": qty})
            st.success(f"Parsed {len(inquiries)} inquiries from text")
        
        if inquiries:
            st.session_state["inquiries"] = inquiries
            st.session_state["results"] = []
            with st.spinner("Matching..."):
                for item in inquiries:
                    cands, ton, bucket, ver = eng.match_products(item["name"], db, config)
                    st.session_state["results"].append({
                        "item": item, "candidates": cands, "ton": ton,
                        "bucket": bucket, "verification": ver
                    })
        else:
            st.warning("Please upload a file or paste inquiry text")

# ==================== TAB 2: RESULTS ====================

with tab2:
    if "results" not in st.session_state or not st.session_state["results"]:
        st.info("👈 Go to Input tab first")
    else:
        results = st.session_state["results"]
        matched = sum(1 for r in results if r["candidates"])
        st.metric("Match Rate", f"{matched}/{len(results)}")
        
        rows = []
        for r in results:
            item = r["item"]
            cands = r["candidates"]
            cat = eng.classify_inquiry(item["name"]) or ""
            comp, _ = eng.detect_competitor(item["name"])
            if comp: cat = f"Competitor({comp})" if not cat else f"{cat}/{comp}"
            
            if not cands:
                rows.append({
                    "#": item["seq"], "Inquiry": item["name"], "Qty": item["qty"],
                    "Category": cat or "N/A", "Model": "—", "Score": "—",
                    "DAP Total (RUB)": "—", "Reason": "Non-LiuGong" if eng.is_non_liugong(item["name"]) else "No match",
                    "Verify": "—"
                })
            else:
                for ci, cand in enumerate(cands[:2]):
                    p = cand["product"]
                    cost = eng.calc_dap(p.get("dap_price_cny", 0) or 0, p.get("scrap_tax_rub", 0) or 0, config)
                    reasons = "; ".join(cand["reasons"]) if cand["reasons"] else "—"
                    vl = ""
                    if r["verification"] and r["verification"].get("top_verification"):
                        vl = r["verification"]["top_verification"].get("level", "")
                    rows.append({
                        "#": item["seq"] if ci == 0 else "",
                        "Inquiry": item["name"] if ci == 0 else "",
                        "Qty": item["qty"] if ci == 0 else "",
                        "Category": cat if ci == 0 else "",
                        "Model": p["model"],
                        "Score": f'{cand["score"]}pts' + (" ★" if ci == 0 else ""),
                        "DAP Total (RUB)": f'{cost["total_rub"]:,.0f}',
                        "Reason": reasons,
                        "Verify": vl
                    })
        
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, height=500)
        
        # Download
        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("📥 Download CSV", csv, f"Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", "text/csv")

# ==================== TAB 3: VERIFICATION ====================

with tab3:
    if "results" not in st.session_state:
        st.info("👈 Complete matching first")
    else:
        ver_rows = []
        for r in st.session_state["results"]:
            ver = r.get("verification", {})
            tv = (ver or {}).get("top_verification", {}) or {}
            ver_rows.append({
                "Inquiry": r["item"]["name"][:40],
                "Category": ver.get("classified_cat", ""),
                "Req. Ton": ver.get("extracted_ton", 0),
                "Local Ton": tv.get("local_ton", 0),
                "Online Ton": tv.get("online_ton", 0),
                "Local HP": tv.get("local_hp", 0),
                "Online HP": tv.get("online_hp", 0),
                "Level": tv.get("level", ""),
                "Status": tv.get("status", ""),
            })
        if ver_rows:
            st.subheader("Cross-Verification Report")
            vdf = pd.DataFrame(ver_rows)
            st.dataframe(vdf, use_container_width=True)
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Dual Verified", sum(1 for v in ver_rows if v["Level"] == "dual"))
            c2.metric("Single Source", sum(1 for v in ver_rows if v["Level"] == "single"))
            c3.metric("Conflict", sum(1 for v in ver_rows if v["Level"] == "conflict"))

st.divider()
st.caption("LiuGong Equipment Matcher v19 | github.com/Tm20200924/liugong-equipment-matcher")
