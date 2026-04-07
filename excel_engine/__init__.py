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

from excel_engine.engine import ExcelEngine
from excel_engine.config import EngineConfig

__all__ = ["ExcelEngine", "EngineConfig", "__version__"]
