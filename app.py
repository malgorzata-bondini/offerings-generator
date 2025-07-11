import streamlit as st
import tempfile
from pathlib import Path
from generator_core import run_generator

st.set_page_config(page_title="ServiceNow Offerings Generator", layout="wide")
st.title("ServiceNow Offerings Generator")

templates = st.file_uploader(
    "Upload one or more offering files (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

with st.form("input_form"):
    keywords_input = st.text_input("Keywords (comma separated)")
    apps_input = st.text_input("New apps (comma separated)")
    days_input = st.text_input("Days (e.g. Mon-Fri; comma separated)")
    hours_input = st.text_input("Hours (e.g. 9-17; comma separated)")
    manager_name = st.text_input("Delivery manager")
    is_global_prod = st.checkbox("Global Prod in Service Offerings?")
    rsp_input = st.text_input("RSP (e.g. 2h)")
    rsl_input = st.text_input("RSL (e.g. 5d)")
    sr_or_im_choice = st.selectbox("Select SR or IM", ["Select…", "SR", "IM"])
    if sr_or_im_choice == "Select…":
        sr_or_im_choice = ""
    include_aliases = st.checkbox("Add Aliases")
    aliases_input = st.text_input("Aliases (comma separated)")
    is_corp = st.checkbox("CORP in Child Service Offerings?")
    service_deliverer = ""
    if is_corp:
        service_deliverer = st.selectbox(
            "Who delivers the service?",
            ["HS DE","DS DE","HS PL","DS PL","HS CY","DS CY","HS UA","DS UA","HS MD","DS MD"]
        )
    support_input = st.text_input("Support group / Managed by group")
    generate = st.form_submit_button("Generate")

if generate:
    if not templates:
        st.error("Please upload at least one xlsx file.")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:
        paths = []
        for f in templates:
            p = Path(tmpdir) / f.name
            p.write_bytes(f.getbuffer())
            paths.append(p)

        try:
            outfile = run_generator(
                keywords=[k.strip().lower() for k in keywords_input.split(",") if k.strip()],
                new_apps=[a.strip() for a in apps_input.split(",") if a.strip()],
                days=[d.strip() for d in days_input.split(",") if d.strip()],
                hours=[h.strip() for h in hours_input.split(",") if h.strip()],
                delivery_manager=manager_name,
                global_prod=is_global_prod,
                rsp_duration=rsp_input,
                rsl_duration=rsl_input,
                sr_or_im=sr_or_im_choice,
                require_corp=is_corp,
                delivering_tag=service_deliverer,
                support_group=support_input.split("/",1)[0],
                managed_by_group=support_input.split("/",1)[-1],
                aliases_on=include_aliases,
                src_dir=Path(tmpdir),
                out_dir=Path(tmpdir)
            )
        except ValueError as e:
            st.error(str(e))
            st.stop()

        st.success("Done, download your file below.")
        st.download_button("Download xlsx", outfile.read_bytes(), file_name=outfile.name)
