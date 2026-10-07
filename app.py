"""
Test Bank Audit Pipeline - Streamlit Web Interface

No bypass logic, no fake delays: progress is driven by real chunk
completion events from the pipeline.
"""

import os
import shutil
import sys
import tempfile
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

load_dotenv(override=True)  # override=True so edits to .env take effect without a full process restart

PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from testbank_pipeline import TestBankPipeline  # noqa: E402


st.set_page_config(
    page_title="Test Bank Audit Pipeline",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
        .main-header { font-size: 2.5rem; color: #1f77b4; text-align: center; margin-bottom: 1rem; }
        .sub-header { font-size: 1.1rem; color: #666; text-align: center; margin-bottom: 2rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<h1 class="main-header">Test Bank Audit Pipeline</h1>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-header">Upload a textbook chapter and a previous-edition test bank to generate a fully audited test bank.</p>',
    unsafe_allow_html=True,
)


def _gemini_cache_key() -> str:
    """Encodes every env var that affects the Gemini client. Passed (without
    a leading underscore) to get_pipeline so Streamlit hashes it as part of
    the cache key — any change to model, auth mode, project, or location
    creates a fresh cached client instead of silently reusing a stale one."""
    return "|".join([
        os.getenv("GEMINI_MODEL_ID", "gemini-3.5-flash"),
        os.getenv("GOOGLE_GENAI_USE_VERTEXAI", ""),
        os.getenv("GOOGLE_CLOUD_PROJECT", ""),
        os.getenv("GOOGLE_CLOUD_LOCATION", ""),
        os.getenv("GEMINI_API_KEY", "")[:8],  # just enough to detect a change, not the full secret
    ])


@st.cache_resource(show_spinner="Initializing pipeline (Gemini client)...")
def get_pipeline(cache_key: str) -> TestBankPipeline:
    """Initialize the pipeline once per unique Gemini config (see
    _gemini_cache_key). cache_key is unused inside the function body —
    its only job is to be part of what Streamlit hashes to decide whether
    to reuse or rebuild the cached TestBankPipeline."""
    return TestBankPipeline(base_path=str(PROJECT_ROOT))


st.subheader("📄 Upload Files")

textbook_file = st.file_uploader(
    "Current Edition Textbook Chapter (.doc, .docx, or .pdf)",
    type=["doc", "docx", "pdf"],
    help="Upload the new edition textbook chapter — this is the authoritative source of truth.",
    key="textbook_chapter",
)

test_bank_file = st.file_uploader(
    "Previous Edition Test Bank (.docx or .doc)",
    type=["docx", "doc"],
    help="Upload the previous edition test bank to be audited and updated.",
    key="previous_test_bank",
)

status_placeholder = st.empty()
progress_bar = st.progress(0)

with status_placeholder.container():
    st.info("Ready to process. Upload both files and click 'Start Audit'.")

st.divider()
_, center_col, _ = st.columns([1, 1, 1])
with center_col:
    start_clicked = st.button(
        "🚀 Start Audit",
        type="primary",
        use_container_width=True,
        disabled=(textbook_file is None or test_bank_file is None),
    )

if start_clicked:
    try:
        pipeline = get_pipeline(_gemini_cache_key())

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)

            textbook_path = temp_path / textbook_file.name
            textbook_path.write_bytes(textbook_file.getvalue())

            test_bank_path = temp_path / test_bank_file.name
            test_bank_path.write_bytes(test_bank_file.getvalue())

            default_prompt = PROJECT_ROOT / "prompts" / "testbank.txt"
            if not default_prompt.exists():
                status_placeholder.error("Prompt file not found at prompts/testbank.txt")
                st.stop()

            def on_progress(message: str, chunks_done: int, total_chunks: int):
                status_placeholder.info(message)
                if total_chunks > 0:
                    progress_bar.progress(min(1.0, chunks_done / total_chunks))

            result = pipeline.process_test_bank(
                textbook_chapter_file=str(textbook_path),
                previous_test_bank_file=str(test_bank_path),
                prompt_file="prompts/testbank.txt",
                output_dir=str(temp_path / "output"),
                on_progress=on_progress,
            )

            progress_bar.progress(1.0)
            status_placeholder.success("Processing complete!")

            output_bytes = Path(result.output_docx_path).read_bytes()
            output_filename = f"{Path(test_bank_file.name).stem}_AUDITED.docx"

            st.session_state["output_data"] = output_bytes
            st.session_state["output_filename"] = output_filename
            st.session_state["pipeline_result"] = result

        st.divider()
        st.success("🎉 Audit completed successfully!")

        summary_cols = st.columns(4)
        summary_cols[0].metric("Total items", result.total_items)
        summary_cols[1].metric("Revised", result.revise_count)
        summary_cols[2].metric("Removed", result.remove_count)
        summary_cols[3].metric("Parse errors", result.parse_error_count)

        if result.parse_error_count:
            st.warning(
                f"{result.parse_error_count} item(s) could not be parsed from Gemini's response and "
                "were carried over unchanged for human review (marked PARSE_ERROR)."
            )

        # ---- Cost summary ------------------------------------------------
        if result.cost_log_path and Path(result.cost_log_path).exists():
            cost_log_bytes = Path(result.cost_log_path).read_bytes()
            st.info(
                f"💰 Cost log updated: `{Path(result.cost_log_path).name}` — "
                "includes a **Summary** sheet with totals, averages, and per-model breakdown."
            )
            st.download_button(
                label="📊 Download Cost Log with Summary (.xlsx)",
                data=cost_log_bytes,
                file_name=Path(result.cost_log_path).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # ---- Output log --------------------------------------------------
        if result.output_log_path and Path(result.output_log_path).exists():
            output_log_bytes = Path(result.output_log_path).read_bytes()
            st.info(
                f"📋 Output log updated: `{Path(result.output_log_path).name}` — "
                "download to see per-item decisions (KEEP / REVISE / REMOVE)."
            )
            st.download_button(
                label="📋 Download Output Log (.xlsx)",
                data=output_log_bytes,
                file_name=Path(result.output_log_path).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.download_button(
            label="📥 Download Audited Test Bank (.docx)",
            data=st.session_state["output_data"],
            file_name=st.session_state["output_filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )

    except Exception as exc:  # noqa: BLE001 - surface full traceback to the user
        status_placeholder.error(f"❌ Error: {exc}")
        progress_bar.progress(0)
        st.exception(exc)

st.divider()
st.markdown(
    "<div style='text-align: center; color: #666; padding: 1rem;'>"
    "<p>Test Bank Audit Pipeline</p></div>",
    unsafe_allow_html=True,
)