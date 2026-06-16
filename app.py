# -*- coding: utf-8 -*-
import streamlit as st
st.set_page_config(page_title="LiuGong Matcher v19", page_icon="🔧", layout="wide")

# ========== SAFE IMPORTS ==========
import os, sys, json, io, traceback, re
from datetime import datetime

# Ensure engine/ is importable
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "engine"))
sys.path.insert(0, os.path.dirname(__file__))

# Try imports, collect errors
errors = []

try:
    import pandas as pd
except Exception as e:
    pd = None
    errors.append(f"pandas: {e}")

HAS_OPENPYXL = False
try:
    import openpyxl
    HAS_OPENPYXL = True
except Exception:
    pass

# Load engine
eng = None
try:
    import engine_v19 as eng
except Exception as e:
    errors.append(f"engine_v19: {e}")

# ========== DATA LOADING ==========
@st.cache_resource
def init_data():
    config = {}
    db = []
    if eng:
        try:
            eng.CONFIG_PATH = os.path.join(os.path.dirname(__file__), "engine", "config.json")
            eng.DB_PATH = os.path.join(os.path.dirname(__file__), "engine", "product_db_full.json")
            eng.COMP_DB_PATH = os.path.join(os.path.dirname(__file__), "engine", "competitor_db.json")
            config = eng.load_config()
            db = eng.load_db()
        except Exception as e:
            errors.append(f"load_data: {e}")

        # Demo fallback
        if not db:
            try:
                from engine.demo_data import to_product_db, to_competitor_db
                with open(eng.DB_PATH, "w", encoding="utf-8") as f:
                    json.dump(to_product_db(), f, ensure_ascii=False)
                with open(eng.COMP_DB_PATH, "w", encoding="utf-8") as f:
                    json.dump(to_competitor_db(), f, ensure_ascii=False)
                db = eng.load_db()
            except Exception as e:
                errors.append(f"demo_fallback: {e}")

    if "config" in st.secrets:
        for k, v in st.secrets["config"].items():
            config[k] = v

    return config, db, errors

config, db, init_errors = init_data()

# ========== SIDEBAR ==========
with st.sidebar:
    st.header("Settings")
    if init_errors:
        with st.expander("Warnings"):
            for e in init_errors:
                st.warning(e)

    if not eng or not db:
        st.error("Engine not loaded. Check warnings above.")
        st.stop()

    rate = st.number_input("Rate CNY->RUB", value=float(config.get("exchange_rate_rub_cny", 11.5)), step=0.1)
    duty = st.number_input("Customs Duty %", value=float(config.get("customs_duty_rate_pct", 5.0)), step=0.1)
    proc = st.number_input("Processing Fee %", value=float(config.get("customs_processing_rate_pct", 0.3)), step=0.1)
    vat = st.number_input("VAT %", value=float(config.get("vat_rate_pct", 22.0)), step=0.1)
    warehouse = st.number_input("Warehouse RUB", value=int(config.get("customs_warehouse_fee_rub", 20000)), step=1000)
    customs_fee = st.number_input("Customs Fee RUB", value=int(config.get("customs_fee_rub", 6000)), step=1000)
    agent_fee = st.number_input("Agent Fee RUB", value=int(config.get("customs_agent_fee_rub", 30000)), step=1000)
    st.divider()
    st.caption(f"Products: {len(db)} | v19")

config.update({
    "exchange_rate_rub_cny": rate, "customs_duty_rate_pct": duty,
    "customs_processing_rate_pct": proc, "vat_rate_pct": vat,
    "customs_warehouse_fee_rub": warehouse, "customs_fee_rub": customs_fee,
    "customs_agent_fee_rub": agent_fee
})

# ========== MAIN ==========
st.title("LiuGong Equipment Matcher v19")
st.caption("Upload inquiry -> Match -> DAP RUB -> Verify")

tab1, tab2, tab3 = st.tabs(["Input", "Results", "Verify"])

with tab1:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Upload File")
        uploaded = st.file_uploader("Excel/CSV", type=["xlsx", "xls", "csv"])
    with c2:
        st.subheader("Or Paste Text")
        txt = st.text_area("One per line", height=150,
            placeholder="Excavator 20t x2\nRoller 3t x6\nBulldozer TY230")

    if st.button("Match Now", type="primary", use_container_width=True):
        inquiries = []
        if uploaded:
            ext = uploaded.name.rsplit(".", 1)[-1].lower()
            if ext in ("xlsx", "xls"):
                if HAS_OPENPYXL:
                    inquiries = eng.parse_inquiry_excel(uploaded)
                else:
                    st.warning("Excel needs openpyxl. Use CSV or text.")
            elif ext == "csv":
                df = pd.read_csv(io.BytesIO(uploaded.getvalue())) if pd else None
                if df is not None:
                    for i, row in df.iterrows():
                        name = str(row.iloc[0])
                        qty = int(str(row.iloc[1])) if len(row) > 1 and str(row.iloc[1]).isdigit() else 1
                        inquiries.append({"seq": str(i+1), "name": name, "qty": qty})
            if inquiries:
                st.success(f"Parsed {len(inquiries)} from file")

        if txt.strip():
            for i, line in enumerate(txt.strip().split("\n")):
                line = line.strip()
                if not line: continue
                qty = 1
                qm = re.search(r"x\s*(\d+)", line, re.IGNORECASE)
                if qm: qty = int(qm.group(1))
                inquiries.append({"seq": str(i+1), "name": line, "qty": qty})
            st.success(f"Parsed {len(inquiries)} from text")

        if inquiries:
            st.session_state["inquiries"] = inquiries
            st.session_state["results"] = []
            with st.spinner("Matching..."):
                for item in inquiries:
                    cands, ton, bucket, ver = eng.match_products(item["name"], db, config)
                    st.session_state["results"].append({
                        "item": item, "candidates": cands,
                        "ton": ton, "bucket": bucket, "verification": ver
                    })
        elif not uploaded:
            st.warning("Upload a file or paste inquiries")

with tab2:
    if "results" not in st.session_state:
        st.info("Complete matching in Input tab first")
    else:
        results = st.session_state["results"]
        matched = sum(1 for r in results if r["candidates"])
        st.metric("Matched", f"{matched}/{len(results)}")
        rows = []
        for r in results:
            item = r["item"]; cands = r["candidates"]
            cat = eng.classify_inquiry(item["name"]) or ""
            comp, _ = eng.detect_competitor(item["name"])
            if comp: cat = f"Competitor({comp})"
            if not cands:
                rows.append({"#": item["seq"], "Inquiry": item["name"], "Qty": item["qty"],
                    "Category": cat or "N/A", "Model": "--", "Score": "--",
                    "DAP RUB": "--", "Reason": "Non-LiuGong" if eng.is_non_liugong(item["name"]) else "No match", "Verify": "--"})
            else:
                for ci, cand in enumerate(cands[:2]):
                    p = cand["product"]
                    cost = eng.calc_dap(p.get("dap_price_cny", 0) or 0, p.get("scrap_tax_rub", 0) or 0, config)
                    reasons = "; ".join(cand["reasons"]) if cand["reasons"] else "--"
                    vl = r.get("verification", {}).get("top_verification", {}) or {}
                    rows.append({"#": item["seq"] if ci==0 else "",
                        "Inquiry": item["name"] if ci==0 else "",
                        "Qty": item["qty"] if ci==0 else "",
                        "Category": cat if ci==0 else "",
                        "Model": p["model"],
                        "Score": f'{cand["score"]}pts' + (" *" if ci==0 else ""),
                        "DAP RUB": f'{cost["total_rub"]:,.0f}',
                        "Reason": reasons,
                        "Verify": vl.get("level", "")})
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, height=500)
        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Download CSV", csv, f"Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

with tab3:
    if "results" not in st.session_state:
        st.info("Complete matching first")
    else:
        vr = []
        for r in st.session_state["results"]:
            v = r.get("verification", {}) or {}
            tv = v.get("top_verification", {}) or {}
            vr.append({"Inquiry": r["item"]["name"][:40], "Cat": v.get("classified_cat",""),
                "Req Ton": v.get("extracted_ton",0), "Local Ton": tv.get("local_ton",0),
                "Web Ton": tv.get("online_ton",0), "Level": tv.get("level","")})
        if vr:
            st.subheader("Cross-Verification")
            st.dataframe(pd.DataFrame(vr), use_container_width=True)
            c1,c2,c3 = st.columns(3)
            c1.metric("Dual", sum(1 for v in vr if v["Level"]=="dual"))
            c2.metric("Single", sum(1 for v in vr if v["Level"]=="single"))
            c3.metric("Conflict", sum(1 for v in vr if v["Level"]=="conflict"))

st.divider()
st.caption("LiuGong Matcher v19 | github.com/Tm20200924/liugong-equipment-matcher")
