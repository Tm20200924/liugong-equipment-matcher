import streamlit as st
st.title("LiuGong Matcher Test")
st.write("If you see this, deployment works!")

# Test imports
try:
    import openpyxl; st.success("openpyxl OK")
except: st.error("openpyxl MISSING")
try:
    import pandas; st.success("pandas OK")
except: st.error("pandas MISSING")
try:
    import requests; st.success("requests OK")
except: st.error("requests MISSING")
try:
    sys.path.insert(0, "engine")
    import engine_v19; st.success("engine_v19 OK")
except Exception as e: st.error(f"engine_v19: {e}")
