import streamlit as st
st.title("LiuGong Matcher v19")
st.success("Deployment works!")
st.write("If you see this, the environment is correct.")

# Minimal import test
try:
    from engine import engine_v19
    st.success(f"engine_v19 loaded: {len(engine_v19.load_db())} products")
except Exception as e:
    st.error(f"engine_v19: {e}")
    import traceback
    st.code(traceback.format_exc())
