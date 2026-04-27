"""
Excel Engine — Cross-platform Streamlit GUI.

Works on macOS, Windows & Linux by defaulting to openpyxl-only mode.
Run with:  streamlit run gui/app.py
"""

from __future__ import annotations

import sys
import traceback
from pathlib import Path

import streamlit as st

# Ensure the project root is on sys.path so `excel_engine` is importable
_project_root = Path(__file__).resolve().parent.parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

from excel_engine import (  # noqa: E402
    ExcelEngine,
    EngineConfig,
    EngineResult,
    Layer,
    __version__,
)

# ────────────────────────────────────────────────────────────────
# Page config & header
# ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel Engine", page_icon="📊", layout="wide")
st.title("📊 Excel Engine")
st.caption("Autonomous Excel Automation — Works on macOS, Windows & Linux")

# ────────────────────────────────────────────────────────────────
# Session-state helpers (keep processed file alive for download)
# ────────────────────────────────────────────────────────────────
if "processed_bytes" not in st.session_state:
    st.session_state.processed_bytes = None
    st.session_state.processed_name = None
    st.session_state.result = None

# ────────────────────────────────────────────────────────────────
# Sidebar — configuration
# ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")

    phase = st.radio(
        "Execution Mode",
        ["Standard Mode (works everywhere)", "Advanced Mode (Mac only — uses Excel app)"],
        index=0,
    )
    dry_run = st.checkbox("Preview Only (show plan without making changes)", value=False)
    max_retries = st.slider("Retry attempts if something fails", 0, 5, 3)
    verify = st.checkbox("Double-check results after each step", value=True)

    st.divider()
    with st.expander("🚀 Performance Options", expanded=False):
        parallel_exec = st.checkbox(
            "Parallel execution (multi-sheet workbooks)",
            value=False,
            help="Execute tasks on different sheets in parallel for faster processing.",
        )
        max_workers = st.slider(
            "Max parallel workers", 2, 8, 4,
            disabled=not parallel_exec,
            help="Number of threads for parallel execution.",
        )
        circuit_breaker = st.checkbox(
            "Circuit breaker (skip failing layers)",
            value=True,
            help="Automatically skip layers that fail repeatedly instead of retrying endlessly.",
        )

    st.divider()
    with st.expander("ℹ️ Engine Info", expanded=False):
        st.write(f"**Version:** {__version__}")

        available_layers = [l.name for l in Layer]
        if phase == "Standard Mode (works everywhere)":
            st.write("**Active layers:** OPENPYXL")
        else:
            st.write(f"**Active layers:** {', '.join(available_layers)}")

        with st.expander("All layers"):
            for layer in Layer:
                st.write(f"Layer {layer.value}: {layer.name}")

# ────────────────────────────────────────────────────────────────
# Main area — file upload
# ────────────────────────────────────────────────────────────────
st.info(
    "👋 **How it works:** Upload your Excel file on the left, then provide your "
    "assignment instructions on the right (upload a file or paste the text). "
    "Click **Run Excel Engine** when ready!"
)

col1, col2 = st.columns(2, gap="large")

with col1:
    st.subheader("📁 Excel Workbook")
    workbook_file = st.file_uploader(
        "Upload your .xlsx or .xlsm file",
        type=["xlsx", "xlsm"],
        help="Legacy .xls (BIFF) format is not supported — please convert to .xlsx first.",
    )

with col2:
    st.subheader("📝 Instructions")
    instruction_file = st.file_uploader(
        "Upload instructions",
        type=["txt", "docx", "pdf", "rtf"],
        help="RTFD (macOS bundle) cannot be uploaded — paste the text instead, "
        "or export to PDF/DOCX first.",
    )
    instruction_text = st.text_area(
        "Or paste instructions here",
        height=150,
        help="Example: 'In cell B2, enter the formula =SUM(A1:A10). "
        "Format column C as currency. Add a header row with bold text.'",
    )

# ────────────────────────────────────────────────────────────────
# Additional resource files (optional)
# ────────────────────────────────────────────────────────────────
with st.expander("📎 Additional Resource / Data Files (optional)", expanded=False):
    st.caption(
        "Upload any companion files the assignment references — "
        "data workbooks, PDFs, images, etc."
    )
    resource_files_upload = st.file_uploader(
        "Upload resource files",
        type=[
            "xlsx", "xls", "xlsm", "csv",
            "pdf", "docx", "doc", "txt", "rtf",
            "zip",
            "png", "jpg", "jpeg", "gif", "bmp", "tiff",
        ],
        accept_multiple_files=True,
        help="Supports .xlsx, .csv, .pdf, .docx, .txt, .zip, and image files.",
    )


# ────────────────────────────────────────────────────────────────
# Helper: display EngineResult
# ────────────────────────────────────────────────────────────────
def _show_result(result: EngineResult) -> None:
    """Render metrics, per-section results, and errors."""
    tasks_failed = result.tasks_total - result.tasks_completed

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Tasks", result.tasks_total)
    m2.metric("Passed", result.tasks_completed)
    m3.metric("Failed", tasks_failed)
    m4.metric("Time", f"{result.elapsed_seconds:.1f}s")

    if result.success:
        st.success(
            f"✅ All done — {result.tasks_completed}/{result.tasks_total} tasks "
            f"across {result.sections_completed}/{result.sections_total} sections"
        )
    else:
        st.warning(
            f"⚠️ Completed {result.tasks_completed}/{result.tasks_total} tasks "
            f"({tasks_failed} failed)"
        )

    # Per-section verification results
    if result.verifications:
        with st.expander("📋 Section Verification Details", expanded=False):
            for sv in result.verifications:
                passed = sv.pass_count
                total = len(sv.results)
                icon = "✅" if sv.all_passed else "❌"
                st.markdown(f"**{icon} {sv.section_id}** — {passed}/{total} passed")
                for vr in sv.results:
                    color = "green" if vr.passed else "red"
                    st.markdown(
                        f"&nbsp;&nbsp;&nbsp;:{color}[{'PASS ✓' if vr.passed else 'FAIL ✗'}] "
                        f"`{vr.task_id}` ({vr.task_type.value}) — {vr.message}"
                    )

    # Errors
    if result.errors:
        with st.expander(f"❌ Errors ({len(result.errors)})", expanded=True):
            for err in result.errors:
                st.error(err)

    if result.failed_tasks:
        with st.expander(f"⛔ Failed Tasks ({len(result.failed_tasks)})"):
            for ft in result.failed_tasks:
                st.write(f"- {ft}")

    # Formula recalculation results
    if result.formula_errors is not None:
        fe = result.formula_errors
        if fe.skipped:
            st.info(f"ℹ️ Formula recalculation skipped — {fe.warning}")
        elif fe.total_errors > 0:
            with st.expander(
                f"⚠️ Formula Errors ({fe.total_errors} errors in "
                f"{fe.total_formulas} formulas)",
                expanded=True,
            ):
                for err_type, info in fe.error_summary.items():
                    st.markdown(f"**{err_type}** — {info['count']} occurrences")
                    for loc in info["locations"][:10]:
                        st.write(f"&nbsp;&nbsp;• {loc}")
        else:
            st.success(
                f"✅ Formulas OK — {fe.total_formulas} formulas, 0 errors"
            )


# ────────────────────────────────────────────────────────────────
# Run button + progress
# ────────────────────────────────────────────────────────────────
if st.button("🚀 Run Excel Engine", type="primary", use_container_width=True):
    # Validation
    if not workbook_file:
        st.error("Please upload an Excel workbook first!")
    elif not instruction_file and not instruction_text.strip():
        st.error("Please upload instructions or paste them above!")
    else:
        # HIGH-3: Warn when both instruction sources are provided
        if instruction_file and instruction_text.strip():
            st.warning(
                "⚠️ Both an instruction file and pasted text were provided. "
                "The uploaded file will be used; pasted text is ignored."
            )

        # Build a persistent working directory inside session_state so the
        # download button can read the bytes after the run block exits.
        import tempfile, shutil  # noqa: E401

        work_dir = Path(tempfile.mkdtemp(prefix="excel_engine_"))

        try:
            # Save uploaded workbook
            wb_path = work_dir / workbook_file.name
            wb_path.write_bytes(workbook_file.read())

            # Save uploaded instructions (if any)
            instr_path = None
            if instruction_file:
                instr_path = work_dir / instruction_file.name
                instr_path.write_bytes(instruction_file.read())

            # Save uploaded resource files (if any)
            res_paths = []
            if resource_files_upload:
                for rf in resource_files_upload:
                    rf_path = work_dir / rf.name
                    rf_path.write_bytes(rf.read())
                    res_paths.append(rf_path)

            # Configure engine
            config = EngineConfig()
            config.max_retries = max_retries
            config.verify_after_each_section = verify
            if phase == "Standard Mode (works everywhere)":
                config.layer_order = [Layer.OPENPYXL]

            # v1.1.0 features
            if hasattr(config, "parallel_execution"):
                config.parallel_execution = parallel_exec
                config.max_workers = max_workers
            if hasattr(config, "circuit_breaker_threshold"):
                config.circuit_breaker_threshold = 5 if circuit_breaker else 0

            engine = ExcelEngine(config)

            # ── Unified scan → plan → (dry-run check) → execute path ──
            with st.spinner("Running Excel Engine..."):
                progress_bar = st.progress(0)
                status_text = st.empty()

                # Step 1: Scan / extract tasks
                status_text.text("📖 Scanning instructions...")
                progress_bar.progress(10)

                if instr_path:
                    tasks = engine.scan(instr_path)
                else:
                    tasks = engine.extractor.extract(instruction_text.strip())

                if len(tasks) == 0:
                    progress_bar.progress(100)
                    status_text.text("")
                    st.warning(
                        "⚠️ No tasks found in your instructions. "
                        "Check that your file contains specific Excel tasks."
                    )
                    st.stop()

                # Step 2: Plan
                status_text.text(f"📋 Planning {len(tasks)} tasks...")
                progress_bar.progress(25)
                plan = engine.plan(tasks)

                # Show plan
                with st.expander(
                    f"📋 Execution Plan — {plan.section_count} sections, "
                    f"{plan.total_tasks} tasks",
                    expanded=False,
                ):
                    for i, section in enumerate(plan.sections):
                        st.write(
                            f"**Section {i + 1}: {section.id}** "
                            f"({section.name}) — {section.task_count} tasks"
                        )
                        for task in section.tasks:
                            st.write(
                                f"&nbsp;&nbsp;- {task.task_type.value}: "
                                f"{task.description}"
                            )

                if dry_run:
                    status_text.text("✅ Dry run complete — no changes made")
                    progress_bar.progress(100)
                    st.success(
                        f"Dry run: {plan.total_tasks} tasks planned across "
                        f"{plan.section_count} sections"
                    )
                    st.session_state.result = None
                    st.session_state.processed_bytes = None
                else:
                    # Step 3: Execute with real-time progress callback
                    status_text.text("⚡ Executing tasks...")
                    progress_bar.progress(50)

                    task_log = st.empty()

                    def _gui_progress(event: dict) -> None:
                        """Update Streamlit UI with per-task progress."""
                        task_id = event.get("task", "")
                        status = event.get("status", "")
                        if status == "executing":
                            status_text.text(f"⚡ Executing: {task_id}...")
                        elif status == "completed":
                            passed = event.get("passed", None)
                            icon = "✅" if passed else ("❌" if passed is False else "⏩")
                            status_text.text(f"{icon} {task_id} done")

                    result = engine.execute(
                        plan, wb_path,
                        progress_callback=_gui_progress,
                    )
                    progress_bar.progress(90)

                    # Step 4: Done
                    status_text.text("✅ Complete!")
                    progress_bar.progress(100)

                    st.session_state.result = result
                    st.session_state.processed_bytes = wb_path.read_bytes()
                    st.session_state.processed_name = (
                        f"{workbook_file.name.rsplit('.', 1)[0]}_completed.xlsx"
                    )

        except Exception as exc:
            exc_str = str(exc).lower()
            if ".xls" in exc_str and "xlsx" not in exc_str or "invalid file" in exc_str:
                st.error(
                    "❌ This file uses an older format. "
                    "Please re-save it as **.xlsx** in Excel and try again."
                )
            elif "permission" in exc_str or "locked" in exc_str:
                st.error(
                    "❌ The file appears to be locked. "
                    "Close it in Excel and try again."
                )
            else:
                st.error(
                    "❌ Something went wrong. "
                    "Please check your files and try again."
                )
            with st.expander("🔧 Technical Details"):
                st.code(traceback.format_exc())
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)

# ────────────────────────────────────────────────────────────────
# Results section (persists across reruns via session_state)
# ────────────────────────────────────────────────────────────────
if st.session_state.result is not None:
    st.divider()
    st.subheader("📊 Results")
    _show_result(st.session_state.result)

if st.session_state.processed_bytes is not None:
    st.download_button(
        "📥 Download Processed Workbook",
        data=st.session_state.processed_bytes,
        file_name=st.session_state.processed_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

if st.session_state.result is not None or st.session_state.processed_bytes is not None:
    if st.button("🔄 Start Over", type="secondary"):
        for key in ["processed_bytes", "processed_name", "result"]:
            st.session_state.pop(key, None)
        st.rerun()

# ────────────────────────────────────────────────────────────────
# Footer
# ────────────────────────────────────────────────────────────────
st.divider()
st.caption("💡 Tip: Use Tab to navigate between fields, Enter to activate buttons.")
st.caption(
    f"Excel Engine v{__version__} · Works on macOS, Windows & Linux · "
    "[GitHub](https://github.com/mmanthe37/excel-engine)"
)
