"""
Excel Engine — Autonomous Excel automation engine for macOS.

A 6-layer architecture for completing Microsoft Excel worksheet assignments:
  Layer 1: openpyxl (offline file manipulation)
  Layer 2: xlwings (live Excel API bridge)
  Layer 3: AppleScript (Excel-specific commands)
  Layer 4: System Events (ribbon/dialog UI automation)
  Layer 5: VBA via VBE (clipboard injection)
  Layer 6: PyAutoGUI (last-resort desktop control)
"""

__version__ = "1.0.0"
__author__ = "Excel Engine Contributors"

from excel_engine.engine import ExcelEngine, EngineResult
from excel_engine.config import EngineConfig, TaskType, Layer
from excel_engine.recalc import recalculate, scan_formula_errors, RecalcResult
from excel_engine.parsers.task_extractor import Task
from excel_engine.planner.task_planner import ExecutionPlan, Section
from excel_engine.verifier.workbook_verifier import WorkbookVerifier, SectionVerification, VerificationResult
from excel_engine.presets.financial import FinancialPreset, apply_financial_preset, apply_number_formats

__all__ = [
    "ExcelEngine", "EngineResult", "EngineConfig",
    "TaskType", "Layer", "Task",
    "ExecutionPlan", "Section",
    "WorkbookVerifier", "SectionVerification", "VerificationResult",
    "recalculate", "scan_formula_errors", "RecalcResult",
    "FinancialPreset", "apply_financial_preset", "apply_number_formats",
    "__version__",
]
