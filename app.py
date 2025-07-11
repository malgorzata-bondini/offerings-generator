import streamlit as st
import tempfile
from pathlib import Path
from generator_core import run_generator

st.set_page_config(page_title="ServiceNow Offerings Generator", layout="wide")
st.title("ServiceNow Offerings Generator")

uploaded_templates = st.file_uploader(
    "Upload one or more offering files (.xlsx) (Child lvl1)",
    type="xlsx",
    accept_multiple_files=True
)

with st.form("input_form"):
    keywords_input      = st.text_input("Keywords (comma separated)")
    apps_input          = st.text_input("New apps (comma separated)")
    day_blocks_input    = st.text_input("Days (e.g. Mon-Fri; comma separated)")
    hour_blocks_input   = st.text_input("Hours (e.g. 9-17; comma separated)")
    manager_name        = st.text_input("Delivery manager")
    is_global_prod      = st.checkbox("Global Prod in Service Offerings?")
    rsp_input           = st.text_input("RSP (e.g. 2h)")
    rsl_input           = st.text_input("RSL (e.g. 5d)")

    sr_or_im_choice     = st.selectbox("Select SR or IM", ["Select…", "SR", "IM"], index=0)
    if sr_or_im_choice == "Select…":
        sr_or_im_choice = ""

    support_group_input = st.text_input("Support group / Managed by group")

    generate_clicked    = st.form_submit_button("Generate")

# live fields below the form
add_aliases       = st.checkbox("Add Aliases")
aliases_input     = st.text_input("Aliases if needed (comma separated)", disabled=not add_aliases)

is_corp           = st.checkbox("CORP in Child Service Offerings?")
service_deliverer = st.text_input(
    "If CORP, who delivers the service (e.g. HS PL)",
    disabled=not is_corp
)

if generate_clicked:
    if not uploaded_templates:
        st.error("Please upload at least one xlsx file.")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:
        paths = []
        for f in uploaded_templates:
            p = Path(tmpdir) / f.name
            p.write_bytes(f.getbuffer())
            paths.append(p)

        # *** here we catch your ValueError and show it directly ***
        try:
            output_path = run_generator(
                keywords           = [k.strip().lower() for k in keywords_input.split(",") if k.strip()],
                new_apps           = [a.strip() for a in apps_input.split(",") if a.strip()],
                days               = [d.strip() for d in day_blocks_input.split(",") if d.strip()],
                hours              = [h.strip() for h in hour_blocks_input.split(",") if h.strip()],
                delivery_manager   = manager_name,
                global_prod        = is_global_prod,
                rsp_duration       = rsp_input,
                rsl_duration       = rsl_input,
                sr_or_im           = sr_or_im_choice,
                require_corp       = is_corp,
                delivering_tag     = service_deliverer.upper() if is_corp else "",
                support_group      = support_group_input.split("/", 1)[0],
                managed_by_group   = support_group_input.split("/", 1)[-1],
                aliases_on         = add_aliases,
                src_dir            = Path(tmpdir),
                out_dir            = Path(tmpdir),
            )
        except ValueError as e:
            st.error(str(e))
            st.stop()
        except Exception as e:
            st.error(f"Unexpected error:\n{e}")
            st.stop()

        st.success("Done, download the file below.")
        st.download_button(
            "Download your file",
            output_path.read_bytes(),
            file_name=output_path.name
        )
