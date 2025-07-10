import streamlit as st
import tempfile
from pathlib import Path
from generator_core import run_generator

st.set_page_config(page_title="Service-Offering Generator", layout="wide")
st.title("Service-Offering Generator")

templates = st.file_uploader(
    "Upload one or more template XLSX files",
    type="xlsx",
    accept_multiple_files=True
)

with st.form("inputs"):
    keywords = st.text_input(
        "Keywords (comma-separated)", "permissions, granting"
    )
    new_apps = st.text_input("New apps (comma-separated)", "SDMS")
    days = st.text_input("Days (e.g. Mon-Fri; comma-separated)", "Mon-Fri")
    hours = st.text_input("Hours (e.g. 9-17; comma-separated)", "9-17")
    delivery_manager = st.text_input("Delivery manager name", "f")
    global_prod = st.checkbox("Global Prod", False)
    rsp = st.text_input("RSP duration (e.g. 2h)", "2h")
    rsl = st.text_input("RSL duration (e.g. 5d)", "5d")
    sr_or_im = st.selectbox("SR or IM", ["SR", "IM"])
    corp = st.checkbox("CORP", True)
    if corp:
        delivering_tag = st.text_input("Who delivers the service", "HS PL")
    else:
        delivering_tag = ""
    support = st.text_input(
        "Support group / Managed by group (split with slash)", "f"
    )
    aliases = st.text_input("Aliases (comma-separated)", "")
    generate = st.form_submit_button("Generate")

if generate:
    if not templates:
        st.error("Please upload at least one XLSX template.")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = []
            for uploaded in templates:
                out_path = Path(tmpdir) / uploaded.name
                out_path.write_bytes(uploaded.getbuffer())
                paths.append(out_path)

            try:
                output_file = run_generator(
                    keywords=[k.strip().lower()
                              for k in keywords.split(",") if k.strip()],
                    new_apps=[a.strip()
                              for a in new_apps.split(",") if a.strip()],
                    days=[d.strip() for d in days.split(",") if d.strip()],
                    hours=[h.strip() for h in hours.split(",") if h.strip()],
                    delivery_manager=delivery_manager,
                    global_prod=global_prod,
                    rsp_duration=rsp,
                    rsl_duration=rsl,
                    sr_or_im=sr_or_im,
                    require_corp=corp,
                    delivering_tag=delivering_tag.upper(),
                    support_group=support.split("/", 1)[0],
                    managed_by_group=support.split("/", 1)[-1],
                    aliases_on=bool(aliases.strip()),
                    src_dir=Path(tmpdir),
                    out_dir=Path(tmpdir),
                )
                st.success("Done – your file is ready to download.")
                st.download_button(
                    "Download XLSX",
                    output_file.read_bytes(),
                    file_name=output_file.name
                )
            except Exception as e:
                st.error(f"Error: {e}")
