import streamlit as st, tempfile
from pathlib import Path
from generator_core import run_generator

st.set_page_config(page_title="Service-Offering Generator", layout="wide")
st.title("Service-Offering Generator")

templates = st.file_uploader(
    "Upload one or more template XLSX files", type="xlsx", accept_multiple_files=True
)

with st.form("inputs"):
    kw      = st.text_input("Keywords", "permissions, granting")
    apps    = st.text_input("New apps", "SDMS")
    days    = st.text_input("Day blocks", "Mon-Fri")
    hours   = st.text_input("Hours blocks", "9-17")
    dm      = st.text_input("Delivery Manager", "f")
    gprod   = st.checkbox("Global Prod", False)
    rsp     = st.text_input("RSP duration", "2h")
    rsl     = st.text_input("RSL duration", "5d")
    sr_im   = st.selectbox("SR or IM", ["SR", "IM"])
    corp    = st.checkbox("Filter CORP rows", True)
    deliv   = st.text_input("Delivering tag", "HS PL")
    mgr     = st.text_input("Support group / Managed by Group", "f")
    aliases = st.checkbox("Add aliases", False)
    submitted = st.form_submit_button("Generate")

if submitted:
    if not templates:
        st.error("Please upload at least one XLSX template.")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = []
            for f in templates:
                p = Path(tmpdir) / f.name
                p.write_bytes(f.getbuffer())
                paths.append(p)

            try:
                outfile = run_generator(
                    keywords=[k.strip().lower() for k in kw.split(",") if k.strip()],
                    new_apps=[a.strip() for a in apps.split(",") if a.strip()],
                    days=[d.strip() for d in days.split(",") if d.strip()],
                    hours=[h.strip() for h in hours.split(",") if h.strip()],
                    delivery_manager=dm,
                    global_prod=gprod,
                    rsp_duration=rsp,
                    rsl_duration=rsl,
                    sr_or_im=sr_im,
                    require_corp=corp,
                    delivering_tag=deliv,
                    support_group=mgr.split("/",1)[0],
                    managed_by_group=mgr.split("/",1)[-1],
                    aliases_on=aliases,
                    src_dir=Path(tmpdir),
                    out_dir=Path(tmpdir)
                )
                st.success("Ready!")
                st.download_button("Download XLSX", outfile.read_bytes(), file_name=outfile.name)
            except Exception as e:
                st.error(str(e))
