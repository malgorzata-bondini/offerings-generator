import streamlit as st
import pandas as pd
from pathlib import Path
import shutil
import tempfile
from generator_core import run_generator, NAMING_CONVENTIONS

st.set_page_config(page_title="Service Offerings Generator", layout="wide")

st.title("🔧 Service Offerings Generator")
st.markdown("---")

col1, col2 = st.columns([1, 1])

with col1:
    st.header("📁 Input Files")
    uploaded_files = st.file_uploader(
        "Upload ALL_Service_Offering Excel files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Upload one or more Excel files with naming pattern: ALL_Service_Offering_*.xlsx"
    )
    
    if uploaded_files:
        st.success(f"✅ Uploaded {len(uploaded_files)} file(s)")
        for file in uploaded_files:
            st.text(f"• {file.name}")

with col2:
    st.header("⚙️ Configuration")
    
    tab1, tab2, tab3, tab4 = st.tabs(["Basic", "Schedule", "Advanced", "Naming"])
    
    with tab1:
        st.subheader("Basic Settings")
        
        keywords = st.text_area(
            "Keywords (one per line)",
            value="software\ninstallation",
            help="Enter keywords to filter service offerings"
        ).strip().split('\n')
        keywords = [k.strip() for k in keywords if k.strip()]
        
        new_apps = st.text_area(
            "Applications (one per line)",
            value="ServiceNow\nOutlook\nTeams",
            help="Enter application names to generate offerings for"
        ).strip().split('\n')
        new_apps = [a.strip() for a in new_apps if a.strip()]
        
        sr_or_im = st.radio("Service Type", ["SR", "IM"], horizontal=True)
        
        delivery_manager = st.text_input("Delivery Manager", value="John Doe")
    
    with tab2:
        st.subheader("Schedule Settings")
        
        col_day, col_hour = st.columns(2)
        with col_day:
            days = st.multiselect(
                "Days",
                ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
                default=["Mon", "Tue", "Wed", "Thu", "Fri"]
            )
        
        with col_hour:
            hours = []
            for day in days:
                hour = st.text_input(f"Hours for {day}", value="8-17", key=f"hour_{day}")
                hours.append(hour)
        
        col_rsp, col_rsl = st.columns(2)
        with col_rsp:
            rsp_duration = st.text_input("RSP Duration", value="12h")
        with col_rsl:
            rsl_duration = st.text_input("RSL Duration", value="24h")
    
    with tab3:
        st.subheader("Advanced Settings")
        
        require_corp = st.checkbox("CORP Service", value=False)
        
        if require_corp:
            delivering_tag = st.text_input("Delivering Tag", value="IT")
        else:
            delivering_tag = ""
        
        global_prod = st.checkbox("Global Prod", value=False)
        
        support_group = st.text_input("Support Group", value="IT Support")
        managed_by_group = st.text_input("Managed by Group", value="IT Operations")
        
        aliases_on = st.checkbox("Enable Aliases", value=True)
    
    with tab4:
        st.subheader("Naming Convention")
        
        naming_convention = st.selectbox(
            "Select Naming Convention:",
            options=list(NAMING_CONVENTIONS.keys()),
            format_func=lambda x: NAMING_CONVENTIONS[x],
            index=3  # Default to child_lvl1_software
        )
        
        st.info(f"Selected: {NAMING_CONVENTIONS[naming_convention]}")
        
        st.markdown("### Example outputs:")
        examples = {
            "child_lvl1_software": "[SR HS PL IT] Software assistance Outlook Prod Mon-Fri 8-17",
            "child_lvl1_hardware": "[SR DS PL IT] Hardware configuration Mon-Fri 7-16:30",
            "child_lvl1_other": "[SR HS PL IT] Mailbox management Mon-Fri 8-16",
            "parent_sam": "[SR HS PL IT] Parent Microsoft Software Request Office 365 Mon-Fri 6-21"
        }
        
        if naming_convention in examples:
            st.code(examples[naming_convention])

st.markdown("---")

if st.button("🚀 Generate Service Offerings", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("⚠️ Please upload at least one Excel file")
    elif not keywords:
        st.error("⚠️ Please enter at least one keyword")
    elif not new_apps:
        st.error("⚠️ Please enter at least one application")
    elif not days or not hours:
        st.error("⚠️ Please configure schedule")
    else:
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                src_dir = Path(temp_dir) / "input"
                out_dir = Path(temp_dir) / "output"
                src_dir.mkdir(exist_ok=True)
                
                for uploaded_file in uploaded_files:
                    file_path = src_dir / uploaded_file.name
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                
                with st.spinner("🔄 Generating service offerings..."):
                    result_file = run_generator(
                        keywords=keywords,
                        new_apps=new_apps,
                        days=days,
                        hours=hours,
                        delivery_manager=delivery_manager,
                        global_prod=global_prod,
                        rsp_duration=rsp_duration,
                        rsl_duration=rsl_duration,
                        sr_or_im=sr_or_im,
                        require_corp=require_corp,
                        delivering_tag=delivering_tag,
                        support_group=support_group,
                        managed_by_group=managed_by_group,
                        aliases_on=aliases_on,
                        src_dir=src_dir,
                        out_dir=out_dir,
                        naming_convention=naming_convention
                    )
                
                st.success("✅ Service offerings generated successfully!")
                
                with open(result_file, "rb") as f:
                    st.download_button(
                        label="📥 Download Generated File",
                        data=f.read(),
                        file_name=result_file.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                st.info(f"Generated file: {result_file.name}")
                
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.exception(e)

st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
        Service Offerings Generator v2.0 | With Naming Conventions Support
    </div>
    """,
    unsafe_allow_html=True
)