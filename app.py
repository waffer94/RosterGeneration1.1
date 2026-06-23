"""
Streamlit web UI for the course roster builder.

Run locally:   streamlit run app.py
Deploy free:   push this repo to GitHub, then deploy at https://share.streamlit.io

Repo must contain:
    app.py
    build_rosters.py
    requirements.txt
    templates/    <- blank template .docx files
    assets/logo.png
    .streamlit/config.toml
"""

import base64
import io
import os
import tempfile
import zipfile

import streamlit as st
import build_rosters as br

TEMPLATE_DIR = "templates"
LOGO_PATH = "assets/logo.png"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

st.set_page_config(page_title="F.A.S.T. Rescue — Roster Builder",
                   page_icon="🚑", layout="centered")

# --------------------------------------------------------------------------- #
# Styling
# --------------------------------------------------------------------------- #
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', system-ui, sans-serif; }
.block-container { padding-top: 2rem; max-width: 820px; }
#MainMenu, footer { visibility: hidden; }

.app-header {
    display: flex; align-items: center; gap: 18px;
    background: linear-gradient(135deg, #2b2b2b 0%, #4a4a4a 100%);
    border-radius: 16px; padding: 22px 26px; margin-bottom: 8px;
    box-shadow: 0 6px 20px rgba(0,0,0,0.12);
}
.app-header img { width: 64px; height: 64px; border-radius: 50%; background:#fff; padding:3px; }
.app-header .title { color: #fff; font-size: 1.55rem; font-weight: 700; line-height: 1.1; }
.app-header .subtitle { color: #f2b8b4; font-size: 0.95rem; font-weight: 500; margin-top: 2px; }
.app-header .accent { color: #E03127; }

.hint { color:#6b7280; font-size:0.9rem; margin: 2px 0 18px; }

div[data-testid="stFileUploader"] {
    border: 2px dashed #d6431f55; border-radius: 14px;
    padding: 10px 14px; background: #fff7f6;
}
.stButton > button, .stDownloadButton > button {
    border-radius: 10px; font-weight: 600; padding: 0.5rem 1.1rem;
}
.stDownloadButton > button { border: 1px solid #E0312733; }

.result-card {
    border: 1px solid #ececec; border-left: 5px solid #E03127;
    border-radius: 12px; padding: 14px 18px; margin-bottom: 18px;
    background: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}
.foot { color:#9aa0a6; font-size:0.8rem; text-align:center; margin-top:28px; }
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------------------------------- #
# Optional password gate
#   Add to Streamlit Cloud Secrets to enable:  password = "your-password"
# --------------------------------------------------------------------------- #
_pw = st.secrets.get("password", None) if hasattr(st, "secrets") else None
if _pw and st.session_state.get("authed") is not True:
    st.markdown("#### 🔒 This tool is password protected")
    entered = st.text_input("Password", type="password", label_visibility="collapsed",
                            placeholder="Enter password")
    if entered == _pw:
        st.session_state["authed"] = True
        st.rerun()
    elif entered:
        st.error("Incorrect password.")
    st.stop()

# --------------------------------------------------------------------------- #
# Header
# --------------------------------------------------------------------------- #
logo_uri = ""
if os.path.exists(LOGO_PATH):
    logo_uri = "data:image/png;base64," + base64.b64encode(open(LOGO_PATH, "rb").read()).decode()

st.markdown(f"""
<div class="app-header">
  {'<img src="' + logo_uri + '"/>' if logo_uri else ''}
  <div>
    <div class="title">Course Roster <span class="accent">Builder</span></div>
    <div class="subtitle">F.A.S.T. Rescue Incorporated</div>
  </div>
</div>
""", unsafe_allow_html=True)
st.markdown('<div class="hint">Upload the Excel export and download print-ready Word rosters.</div>',
            unsafe_allow_html=True)

with st.expander("How it works"):
    st.markdown(
        "1. Export your courses to Excel.\n"
        "2. Upload the `.xlsx` below.\n"
        "3. Click **Generate rosters** and download the Word files.\n\n"
        "The right template is chosen automatically from the course name "
        "(First Aid In-Class / Blended / Recertification, Working at Heights, or a "
        "General template for anything else). Cancelled courses are skipped unless you opt in."
    )

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
uploaded = st.file_uploader("Excel export (.xlsx)", type=["xlsx"])
c1, c2 = st.columns(2)
keep_order = c1.toggle("Keep export order", help="Otherwise names are sorted A–Z.")
include_cancelled = c2.toggle("Include cancelled courses")

go = st.button("Generate rosters", type="primary", use_container_width=True,
               disabled=uploaded is None)

# --------------------------------------------------------------------------- #
# Generate
# --------------------------------------------------------------------------- #
if uploaded and go:
    template_map = br.load_template_map(TEMPLATE_DIR)
    if not template_map:
        st.error(f"No templates found in '{TEMPLATE_DIR}/'. Add the blank template "
                 ".docx files there and redeploy.")
        st.stop()

    log, files = [], []
    with st.spinner("Building rosters…"):
        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = os.path.join(tmp, "export.xlsx")
            with open(xlsx_path, "wb") as f:
                f.write(uploaded.getbuffer())
            out_dir = os.path.join(tmp, "out")
            os.makedirs(out_dir, exist_ok=True)

            made = 0
            for rec in br.read_export(xlsx_path):
                status, info, ctype, n = br.build_one(
                    rec, template_map, out_dir,
                    sort_alpha=not keep_order,
                    include_cancelled=include_cancelled)
                if status == "ok":
                    made += 1
                    with open(info, "rb") as fh:
                        files.append((os.path.basename(info), fh.read()))
                    log.append(("ok", os.path.basename(info), f"{ctype} · {n} participants"))
                else:
                    log.append(("skip", rec.get("customer") or "Unknown", info))

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in files:
                z.writestr(name, data)

    st.session_state["results"] = {
        "files": files, "zip": zip_buf.getvalue(), "log": log,
        "made": made, "skipped": sum(1 for s, *_ in log if s == "skip"),
    }

# --------------------------------------------------------------------------- #
# Results (persisted so download clicks don't clear them)
# --------------------------------------------------------------------------- #
res = st.session_state.get("results")
if res:
    st.divider()
    m1, m2 = st.columns(2)
    m1.metric("Rosters created", res["made"])
    m2.metric("Skipped", res["skipped"])

    if res["files"]:
        st.download_button("⬇  Download all rosters (.zip)", res["zip"],
                           file_name="rosters.zip", mime="application/zip",
                           type="primary", use_container_width=True)

    for status, title, detail in res["log"]:
        if status == "ok":
            st.markdown(f'<div class="result-card">✅ <b>{title}</b><br>'
                        f'<span style="color:#6b7280">{detail}</span></div>',
                        unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="result-card" style="border-left-color:#d9a300">'
                        f'⏭️ <b>Skipped — {title}</b><br>'
                        f'<span style="color:#6b7280">{detail}</span></div>',
                        unsafe_allow_html=True)

    if res["files"]:
        with st.expander("Download individual files"):
            for i, (name, data) in enumerate(res["files"]):
                st.download_button(name, data, file_name=name,
                                   mime=DOCX_MIME, key=f"dl_{i}",
                                   use_container_width=True)

st.markdown('<div class="foot">F.A.S.T. Rescue Incorporated · Course Roster Builder</div>',
            unsafe_allow_html=True)
