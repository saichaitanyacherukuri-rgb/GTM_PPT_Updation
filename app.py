"""
Streamlit Web UI for the PPT Auto-Updater.
Run with: py -m streamlit run app.py

Cloud-compatible: all file I/O happens inside a per-session temp directory
so the app works on Streamlit Community Cloud without any local filesystem
write access to the source directory.
"""

import json
import io
import re
import importlib
import shutil
import tempfile
from pathlib import Path

import streamlit as st

import update_ppt
importlib.reload(update_ppt)

from update_ppt import (
    SCRIPT_DIR, CONFIG_FILE, DEFAULTS, MASTER_FILE_NAME,
    find_data_files, run_update, handle_insert, handle_remove,
    _load_full_config, _save_full_config,
)

st.set_page_config(page_title="GTM Deck Updater", layout="wide")
st.title("GTM Deck Auto-Updater")


# ---------------------------------------------------------------------------
# Session-scoped temp directory
# Every user session gets its own isolated working folder in /tmp (or Windows
# equivalent). Uploaded files are written there; nothing ever touches the
# source-code directory, making this cloud-safe.
# ---------------------------------------------------------------------------

def _init_session():
    """Create the per-session temp dir and seed it with the bundled config."""
    if "work_dir" not in st.session_state:
        work_dir = Path(tempfile.mkdtemp(prefix="gtm_deck_"))
        st.session_state["work_dir"] = work_dir
        # Copy the bundled slide_config.json into the temp dir if it exists
        if CONFIG_FILE.exists():
            shutil.copy(CONFIG_FILE, work_dir / "slide_config.json")

_init_session()


def get_work_dir() -> Path:
    return st.session_state["work_dir"]


def get_config_file() -> Path:
    return get_work_dir() / "slide_config.json"


# ---------------------------------------------------------------------------
# Helper functions for saving uploaded files into the session temp dir
# ---------------------------------------------------------------------------

def save_uploaded_ppt(ppt_file) -> Path:
    dest = get_work_dir() / ppt_file.name
    dest.write_bytes(ppt_file.getvalue())
    return dest


def save_uploaded_data(data_files_up: list) -> list:
    saved = []
    for f in data_files_up:
        dest = get_work_dir() / f.name
        try:
            dest.write_bytes(f.getvalue())
            saved.append(dest)
        except Exception as e:
            st.warning(f"Cannot write {f.name}: {e}")
    return saved


def save_single_file(uploaded_file) -> Path:
    dest = get_work_dir() / uploaded_file.name
    dest.write_bytes(uploaded_file.getvalue())
    return dest


# ---------------------------------------------------------------------------
# Sidebar: Mode selection + File uploads
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("1. Data Source Mode")
    data_mode = st.radio(
        "How do you want to provide data?",
        options=["individual", "master"],
        format_func=lambda x: (
            "Individual slide files  (Slide 10.xlsx, Slide 13.xlsx …)"
            if x == "individual"
            else f"Single master file  ({MASTER_FILE_NAME})"
        ),
        key="data_mode",
    )

    st.markdown("---")
    st.header("2. Upload Files")

    ppt_file = st.file_uploader(
        "PowerPoint template (.pptx)",
        type=["pptx"],
        help="Upload the PPT deck you want to update.",
    )

    master_file_up = None
    data_files_up  = []

    if data_mode == "master":
        master_file_up = st.file_uploader(
            f"Master Excel file ({MASTER_FILE_NAME})",
            type=["xlsx"],
            help=f"Upload {MASTER_FILE_NAME} with sheets named 'slide 10', 'slide 15.1', etc.",
            key="master_uploader",
        )
        if master_file_up:
            st.info(f"Master file: {master_file_up.name}")
    else:
        data_files_up = st.file_uploader(
            "Data files (Slide N.csv / .xlsx)",
            type=["csv", "xlsx"],
            accept_multiple_files=True,
            help="Upload files named like 'Slide 14.csv', 'Slide 15.xlsx'.",
            key="individual_uploader",
        ) or []
        if data_files_up:
            st.info(f"{len(data_files_up)} data file(s) uploaded")
            for f in data_files_up:
                st.caption(f"  - {f.name}")

    if ppt_file:
        st.success(f"PPT: {ppt_file.name}")

    st.markdown("---")
    st.header("3. Update from Uploads")

    ready = ppt_file and (
        (data_mode == "master" and master_file_up) or
        (data_mode == "individual" and data_files_up)
    )

    if ready:
        if st.button("Update PPT (Uploaded Files)", type="primary", use_container_width=True, key="sidebar_update"):
            save_uploaded_ppt(ppt_file)

            if data_mode == "master":
                save_single_file(master_file_up)
                only_slides_arg = None
            else:
                save_uploaded_data(data_files_up)
                pattern = re.compile(r"^slide\s+(\d+)\.(csv|xlsx)$", re.IGNORECASE)
                only_slides_arg = set()
                for f in data_files_up:
                    m = pattern.match(f.name)
                    if m:
                        only_slides_arg.add(int(m.group(1)))

            with st.spinner("Updating presentation..."):
                import io as _io
                import contextlib

                log_buffer = _io.StringIO()
                with contextlib.redirect_stdout(log_buffer):
                    try:
                        output_path = run_update(
                            get_work_dir(), ppt_file.name,
                            only_slides=only_slides_arg,
                            mode=data_mode,
                            config_file=get_config_file(),
                        )
                        success = True
                    except SystemExit as e:
                        success = False
                        error_msg = str(e)

            st.code(log_buffer.getvalue(), language="text")

            if success:
                st.success(f"Done: {output_path.name}")
                with open(output_path, "rb") as fp:
                    st.download_button(
                        "Download Updated PPT",
                        data=fp.read(),
                        file_name=output_path.name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True,
                        key="sidebar_download",
                    )
            else:
                st.error(f"Failed: {error_msg}")
    else:
        if data_mode == "master":
            st.caption(f"Upload a PPT and {MASTER_FILE_NAME} above to enable update.")
        else:
            st.caption("Upload a PPT and at least one data file above to enable update.")


# ---------------------------------------------------------------------------
# Main area tabs
# ---------------------------------------------------------------------------
tab_update, tab_config, tab_slides = st.tabs([
    "Update PPT", "Slide Config", "Insert / Remove Slides"
])


# ========================== TAB 1: Update PPT =============================
with tab_update:
    st.subheader("Run Update")

    existing_data = find_data_files(get_work_dir())
    pptx_files = [
        f for f in get_work_dir().glob("*.pptx")
        if not f.name.startswith("~$") and "_updated" not in f.name
    ]

    if pptx_files or ppt_file:
        options = [f.name for f in pptx_files]
        if ppt_file and ppt_file.name not in options:
            options.insert(0, ppt_file.name)

        selected_ppt = st.selectbox("Select PPT to update", options)

        if data_mode == "master":
            master_exists = (get_work_dir() / MASTER_FILE_NAME).exists()
            if master_file_up or master_exists:
                label = master_file_up.name if master_file_up else f"{MASTER_FILE_NAME} (on disk)"
                st.write(f"Master file: **{label}**")
            else:
                st.warning(f"Master file not found. Upload {MASTER_FILE_NAME} in the sidebar.")
        else:
            combined_slides = set(existing_data.keys())
            if data_files_up:
                pattern = re.compile(r"^slide\s+(\d+)\.(csv|xlsx)$", re.IGNORECASE)
                for f in data_files_up:
                    m = pattern.match(f.name)
                    if m:
                        combined_slides.add(int(m.group(1)))
            if combined_slides:
                st.write(f"Slides with data: **{sorted(combined_slides)}**")
            else:
                st.warning("No data files found. Upload files named like 'Slide 14.csv'.")

        if st.button("Update PPT", type="primary", use_container_width=True):
            if ppt_file:
                save_uploaded_ppt(ppt_file)
            if data_mode == "master" and master_file_up:
                save_single_file(master_file_up)
            elif data_mode == "individual" and data_files_up:
                save_uploaded_data(data_files_up)

            with st.spinner("Updating presentation..."):
                import io as _io
                import contextlib

                log_buffer = _io.StringIO()
                with contextlib.redirect_stdout(log_buffer):
                    try:
                        output_path = run_update(
                            get_work_dir(), selected_ppt,
                            mode=data_mode,
                            config_file=get_config_file(),
                        )
                        success = True
                    except SystemExit as e:
                        success = False
                        error_msg = str(e)

            log_text = log_buffer.getvalue()
            st.code(log_text, language="text")

            if success:
                st.success(f"Updated PPT saved: {output_path.name}")
                with open(output_path, "rb") as fp:
                    st.download_button(
                        "Download Updated PPT",
                        data=fp.read(),
                        file_name=output_path.name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True,
                    )
            else:
                st.error(f"Update failed: {error_msg}")
    else:
        st.info("Upload a PPT file in the sidebar to get started.")


# ========================== TAB 2: Config Editor ==========================
with tab_config:
    st.subheader("Slide Configuration")
    st.caption("Edit the JSON config below. Changes are saved to your session when you click 'Save Config'.")

    raw_config = _load_full_config(config_file=get_config_file())
    config_text = json.dumps(raw_config, indent=4, ensure_ascii=False)

    edited_config = st.text_area(
        "slide_config.json",
        value=config_text,
        height=500,
        label_visibility="collapsed",
    )

    col_save, col_dl, col_reset = st.columns(3)
    with col_save:
        if st.button("Save Config", type="primary", use_container_width=True):
            try:
                parsed = json.loads(edited_config)
                _save_full_config(parsed, config_file=get_config_file())
                st.success("Config saved to session.")
                st.rerun()
            except json.JSONDecodeError as e:
                st.error(f"Invalid JSON: {e}")
    with col_dl:
        st.download_button(
            "Download Config",
            data=edited_config.encode("utf-8"),
            file_name="slide_config.json",
            mime="application/json",
            use_container_width=True,
            help="Download the config file to keep a local copy.",
        )
    with col_reset:
        if st.button("Reload from session", use_container_width=True):
            st.rerun()


# ========================== TAB 3: Insert / Remove Slides =================
with tab_slides:
    st.subheader("Insert / Remove Slides")
    st.caption(
        "When you add or remove a slide in the PPT, use these tools to "
        "automatically renumber all configs and data files."
    )

    current_data = find_data_files(get_work_dir())
    raw_cfg = _load_full_config(config_file=get_config_file())
    config_slides = sorted(int(k) for k in raw_cfg if not k.startswith("_") and k.isdigit())
    file_slides = sorted(current_data.keys())

    st.write(f"**Config entries:** {config_slides}")
    st.write(f"**Data files:** {file_slides}")

    st.markdown("---")

    col_ins, col_rem = st.columns(2)

    with col_ins:
        st.markdown("**Insert a slide**")
        insert_pos = st.number_input(
            "Inserted at position", min_value=1, value=1, step=1, key="ins_pos"
        )
        ins_dry = st.checkbox("Dry run (preview only)", value=True, key="ins_dry")

        if st.button("Insert Slide", use_container_width=True):
            import io as _io, contextlib
            buf = _io.StringIO()
            with contextlib.redirect_stdout(buf):
                handle_insert(
                    int(insert_pos), dry_run=ins_dry,
                    folder=get_work_dir(), config_file=get_config_file(),
                )
            st.code(buf.getvalue(), language="text")
            if not ins_dry:
                st.rerun()

    with col_rem:
        st.markdown("**Remove a slide**")
        remove_pos = st.number_input(
            "Removed at position", min_value=1, value=1, step=1, key="rem_pos"
        )
        rem_dry = st.checkbox("Dry run (preview only)", value=True, key="rem_dry")

        if st.button("Remove Slide", use_container_width=True):
            import io as _io, contextlib
            buf = _io.StringIO()
            with contextlib.redirect_stdout(buf):
                handle_remove(
                    int(remove_pos), dry_run=rem_dry,
                    folder=get_work_dir(), config_file=get_config_file(),
                )
            st.code(buf.getvalue(), language="text")
            if not rem_dry:
                st.rerun()
