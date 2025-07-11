import streamlit as st
import tempfile
from pathlib import Path
from generator_core import run_generator

st.set_page_config(page_title="Service-Offering Generator", layout="wide")
st.title("Service-Offering Generator")

# Always-live checkboxes + inputs for instant enable/disable
add_aliases = st.checkbox("Add Aliases")
aliases_input = st.text_input("Aliases (comma-separated)", disabled=not add_aliases)

is_corp = st.checkbox("CORP")
service_deliverer = st.text_input(
    "Who delivers the service (e.g. HS PL)",
    disabled=not is_corp
)

# Now the rest of the inputs in one form:
uploaded_templates = st.file_uploader(
    "Upload one or more offerings files (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

with st.form("main_form"):
    keywords_input      = st.text_input("Keywords (comma-separated)")
    apps_input          = st.text_input("New apps (comma-separated)")
    day_blocks_input    = st.text_input("Days (e.g. Mon-Fri; comma-separated)")
    hour_blocks_input   = st.text_input("Hours (e.g. 9-17; comma-separated)")
    manager_name        = st.text_input("Delivery manager name")
    is_global_prod      = st.checkbox("Global Prod", key="__gp__")
    rsp_input           = st.text_input("RSP duration (e.g. 2h)")
    rsl_input           = st.text_input("RSL duration (e.g. 5d)")
    sr_or_im_choice     = st.selectbox("Select SR or IM", ["Select…", "SR", "IM"], index=0)
    if sr_or_im_choice == "Select…":
        sr_or_im_choice = ""
    support_input       = st.text_input("Support group / Managed-by group")
    generate_clicked    = st.form_submit_button("Generate")

if generate_clicked:
    if not uploaded_templates:
        st.error("Please upload at least one XLSX file.")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = []
            for file in uploaded_templates:
                p = Path(tmpdir) / file.name
                p.write_bytes(file.getbuffer())
                paths.append(p)

            try:
                result = run_generator(
                    keywords=[k.strip().lower() for k in keywords_input.split(",") if k.strip()],
                    new_apps=[a.strip() for a in apps_input.split(",")  if a.strip()],
                    days=[d.strip()       for d in day_blocks_input.split(",") if d.strip()],
                    hours=[h.strip()      for h in hour_blocks_input.split(",") if h.strip()],
                    delivery_manager=manager_name,
                    global_prod=is_global_prod,
                    rsp_duration=rsp_input,
                    rsl_duration=rsl_input,
                    sr_or_im=sr_or_im_choice,
                    require_corp=is_corp,
                    delivering_tag=service_deliverer.upper() if is_corp else "",
                    support_group=support_input.split("/",1)[0],
                    managed_by_group=support_input.split("/",1)[-1],
                    aliases_on=add_aliases,
                    src_dir=Path(tmpdir),
                    out_dir=Path(tmpdir),
                )
                st.success("Done – download your file below.")
                st.download_button("Download XLSX", result.read_bytes(), file_name=result.name)
            except Exception as e:
                st.error(f"Something went wrong: {e}")
