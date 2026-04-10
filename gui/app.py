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
        ["Offline (openpyxl)", "Full (all layers — macOS only)"],
        index=0,
    )
    dry_run = st.checkbox("Dry Run (plan only, don't modify)", value=False)
    max_retries = st.slider("Max Retries", 0, 5, 2)
    verify = st.checkbox("Verify after each section", value=True)

    st.divider()
    st.header("ℹ️ Engine Info")
    st.write(f"**Version:** {__version__}")

    available_layers = [l.name for l in Layer]
    if phase == "Offline (openpyxl)":
        st.write("**Active layers:** OPENPYXL")
    else:
        st.write(f"**Active layers:** {', '.join(available_layers)}")

    with st.expander("All layers"):
        for layer in Layer:
            st.write(f"Layer {layer.value}: {layer.name}")

# ────────────────────────────────────────────────────────────────
# Main area — file upload
# ────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Excel Workbook")
    workbook_file = st.file_uploader("Upload your .xlsx file", type=["xlsx", "xls"])

with col2:
    st.subheader("📝 Instructions")
    instruction_file = st.file_uploader(
        "Upload instructions", type=["txt", "docx", "pdf", "rtfd"]
    )
    instruction_text = st.text_area("Or paste instructions here", height=150)


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
                        f"&nbsp;&nbsp;&nbsp;:{color}[{'✓' if vr.passed else '✗'}] "
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

            # Configure engine
            config = EngineConfig()
            config.max_retries = max_retries
            config.verify_after_each_section = verify
            if phase == "Offline (openpyxl)":
                config.layer_order = [Layer.OPENPYXL]

            engine = ExcelEngine(config)

            # ── Text-only path (no file to scan) ──
            if not instr_path and instruction_text.strip():
                with st.spinner("Running Excel Engine..."):
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    status_text.text("⚡ Running with pasted instructions...")
                    progress_bar.progress(20)

                    result = engine.run(
                        workbook=wb_path,
                        instruction_text=instruction_text.strip(),
                    )

                    progress_bar.progress(100)
                    status_text.text("✅ Complete!")

                st.session_state.result = result
                st.session_state.processed_bytes = wb_path.read_bytes()
                st.session_state.processed_name = f"processed_{workbook_file.name}"

            # ── File-based path: scan → plan → execute ──
            else:
                with st.spinner("Running Excel Engine..."):
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    # Step 1: Scan
                    status_text.text("📖 Scanning instructions...")
                    progress_bar.progress(10)
                    tasks = engine.scan(instr_path)

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
                        # Step 3: Execute
                        status_text.text("⚡ Executing tasks...")
                        progress_bar.progress(50)
                        result = engine.execute(plan, wb_path)
                        progress_bar.progress(90)

                        # Step 4: Done
                        status_text.text("✅ Complete!")
                        progress_bar.progress(100)

                        st.session_state.result = result
                        st.session_state.processed_bytes = wb_path.read_bytes()
                        st.session_state.processed_name = (
                            f"processed_{workbook_file.name}"
                        )

        except Exception as exc:
            st.error(f"❌ Error: {exc}")
            with st.expander("Error Details"):
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

# ────────────────────────────────────────────────────────────────
# Footer
# ────────────────────────────────────────────────────────────────
st.divider()
st.caption(
    f"Excel Engine v{__version__} · Works on macOS, Windows & Linux · "
    "[GitHub](https://github.com/mmanthe37/excel-engine)"
)
