"""
Excel enums, style names, AppleScript constants, and other lookup tables.

These values come from 114+ checkpoints of real-world SAM assignment automation.
"""

from __future__ import annotations


class ExcelConstants:
    """Central reference for Excel constants used across layers."""

    # ── AppleScript sort/subtotal enum values ──
    XL_ASCENDING = 1
    XL_DESCENDING = 2

    # Subtotal function enums (AppleScript / VBA)
    XL_SUM = -4157
    XL_AVERAGE = -4106
    XL_COUNT = -4112
    XL_COUNT_NUMS = 2
    XL_MAX = -4136
    XL_MIN = -4139
    XL_PRODUCT = -4149
    XL_STDEV = -4155
    XL_VAR = -4164

    SUBTOTAL_FUNCTIONS = {
        "sum": XL_SUM,
        "average": XL_AVERAGE,
        "count": XL_COUNT,
        "countnums": XL_COUNT_NUMS,
        "max": XL_MAX,
        "min": XL_MIN,
        "product": XL_PRODUCT,
        "stdev": XL_STDEV,
        "var": XL_VAR,
    }

    # ── AppleScript save format ──
    SAVE_FORMAT_XLSX = "Excel XML file format"
    SAVE_FORMAT_XLSM = "Excel macro-enabled XML file format"

    # ── Table style names (exact, case-sensitive) ──
    TABLE_STYLES = [
        "TableStyleLight1", "TableStyleLight2", "TableStyleLight3",
        "TableStyleLight4", "TableStyleLight5", "TableStyleLight6",
        "TableStyleLight7", "TableStyleLight8", "TableStyleLight9",
        "TableStyleLight10", "TableStyleLight11", "TableStyleLight12",
        "TableStyleLight13", "TableStyleLight14",
        "TableStyleMedium1", "TableStyleMedium2", "TableStyleMedium3",
        "TableStyleMedium4", "TableStyleMedium5", "TableStyleMedium6",
        "TableStyleMedium7", "TableStyleMedium8", "TableStyleMedium9",
        "TableStyleMedium10", "TableStyleMedium11", "TableStyleMedium12",
        "TableStyleMedium13", "TableStyleMedium14",
        "TableStyleDark1", "TableStyleDark2", "TableStyleDark3",
        "TableStyleDark4", "TableStyleDark5", "TableStyleDark6",
        "TableStyleDark7", "TableStyleDark8", "TableStyleDark9",
        "TableStyleDark10", "TableStyleDark11",
    ]

    # ── Chart types (openpyxl) ──
    CHART_TYPES = {
        "bar": "BarChart",
        "bar3d": "BarChart3D",
        "line": "LineChart",
        "pie": "PieChart",
        "area": "AreaChart",
        "scatter": "ScatterChart",
        "doughnut": "DoughnutChart",
        "radar": "RadarChart",
        "bubble": "BubbleChart",
    }

    # ── Number format codes ──
    NUMBER_FORMATS = {
        "general": "General",
        "number": "#,##0.00",
        "number_no_dec": "#,##0",
        "currency": "$#,##0.00",
        "currency_no_dec": "$#,##0",
        "accounting": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
        "percentage": "0.00%",
        "percentage_no_dec": "0%",
        "date_short": "m/d/yyyy",
        "date_long": "mmmm d, yyyy",
        "date_medium": "mmm d, yyyy",
        "time": "h:mm AM/PM",
        "text": "@",
        "phone": "[<=9999999]###-####;(###) ###-####",
        "zip": "00000",
    }

    # ── Conditional formatting operators ──
    CF_OPERATORS = {
        "greater_than": "greaterThan",
        "less_than": "lessThan",
        "equal": "equal",
        "not_equal": "notEqual",
        "greater_equal": "greaterThanOrEqual",
        "less_equal": "lessThanOrEqual",
        "between": "between",
        "not_between": "notBetween",
    }

    # ── Data validation types ──
    DV_TYPES = {
        "list": "list",
        "whole": "whole",
        "decimal": "decimal",
        "date": "date",
        "time": "time",
        "text_length": "textLength",
        "custom": "custom",
    }

    # ── Border styles ──
    BORDER_STYLES = [
        "thin", "medium", "thick", "double", "hair",
        "dotted", "dashed", "dashDot", "dashDotDot",
        "mediumDashed", "mediumDashDot", "mediumDashDotDot",
        "slantDashDot",
    ]

    # ── Alignment options ──
    HORIZONTAL_ALIGNMENTS = [
        "general", "left", "center", "right", "fill",
        "justify", "centerContinuous", "distributed",
    ]
    VERTICAL_ALIGNMENTS = [
        "top", "center", "bottom", "justify", "distributed",
    ]

    # ── System Events UI references ──
    RIBBON_TABS = {
        "home": "Home",
        "insert": "Insert",
        "page_layout": "Page Layout",
        "formulas": "Formulas",
        "data": "Data",
        "review": "Review",
        "view": "View",
        "table_design": "Table Design",
        "chart_design": "Chart Design",
        "slicer": "Slicer",
    }
