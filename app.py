import streamlit as st
import tempfile
from pathlib import Path
from generator_core import run_generator

st.set_page_config(page_title="Service-Offering Generator", layout="wide")
st.title("Service-Offering Generator")

uploaded_templates = st.file_uploader(
    "Upload one or more files with offerings",
    type="xlsx",
    accept_multiple_files=True
)

with st.form("input_form"):
    keywords_input = st.text_input("Keywords (comma-separated)", "")
    apps_input = st.text_input("New apps (comma-separated)", "")
    day_blocks_input = st.text_input("Days (e.g. Mon–Fri; comma-separated)", "")
    hour_blocks_input = st.text_input("Hours (e.g. 9–17; comma-separated)", "")
    manager_name = st.text_input("Delivery manager name", "")
    is_global_prod = st.checkbox("Global Prod", False)
    rsp_duration_input = st.text_input("RSP duration (e.g. 2h)", "")
    rsl_duration_input = st.text_input("RSL duration (e.g. 5d)", "")

    sr_or_im_choice = st.selectbox(
        "Select SR or IM",
        options=["Select…", "SR", "IM"],
        index=0
    )
    if sr_or_im_choice == "Select…":
        sr_or_im_choice = ""

    is_corp = st.checkbox("CORP", False)
    service_deliverer = st.text_input(
        "Who delivers the service",
        "",
        disabled=not is_corp
    )

    include_aliases = st.checkbox("Add Aliases", False)
    aliases_input = st.text_input(
        "Aliases",
        "",
        disabled=not include_aliases
    )

    support_group_input = st.text_input(
        "Support group / Managed by group", ""
    )

    generate_button_clicked = st.form_submit_button("Generate")

if generate_button_clicked:
    if not uploaded_templates:
        st.error("Please upload at least one .xlsx file")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            template_paths = []
            for template_file in uploaded_templates:
                temp_path = Path(tmpdir) / template_file.name
                temp_path.write_bytes(template_file.getbuffer())
                template_paths.append(temp_path)

            try:
                # if the check-box is off, ignore the typed value
                delivering_tag_final = service_deliverer.upper() if is_corp else ""
                aliases_flag = include_aliases  # run_generator only needs the flag

                result_path = run_generator(
                    keywords=[k.strip().lower() for k in keywords_input.split(",") if k.strip()],
                    new_apps=[a.strip() for a in apps_input.split(",") if a.strip()],
                    days=[d.strip() for d in day_blocks_input.split(",") if d.strip()],
                    hours=[h.strip() for h in hour_blocks_input.split(",") if h.strip()],
                    delivery_manager=manager_name,
                    global_prod=is_global_prod,
                    rsp_duration=rsp_duration_input,
                    rsl_duration=rsl_duration_input,
                    sr_or_im=sr_or_im_choice,
                    require_corp=is_corp,
                    delivering_tag=delivering_tag_final,
                    support_group=support_group_input.split("/", 1)[0],
                    managed_by_group=support_group_input.split("/", 1)[-1],
                    aliases_on=aliases_flag,
                    src_dir=Path(tmpdir),
                    out_dir=Path(tmpdir),
                )
                st.success("Download your file below")
                st.download_button(
                    "Download",
                    result_path.read_bytes(),
                    file_name=result_path.name
                )
            except Exception as e:
                st.error(f"Oops, something went wrong: {e}")
