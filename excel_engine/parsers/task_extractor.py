"""
Task Extractor — Extract structured tasks from parsed instruction text.

Identifies specific Excel operations (formulas, tables, formatting, charts, etc.)
from natural language instructions and maps them to TaskType enums.  Trained against
114+ SAM textbook checkpoints covering Modules 3–7.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Optional

from excel_engine.config import TaskType
from excel_engine.parsers.instruction_parser import InstructionStep

logger = logging.getLogger(__name__)


@dataclass
class Task:
    """A single extracted task to be executed."""
    id: str
    task_type: TaskType
    description: str
    sheet: Optional[str] = None
    cell: Optional[str] = None
    range: Optional[str] = None
    value: Optional[str] = None
    formula: Optional[str] = None
    style: Optional[str] = None
    params: dict = field(default_factory=dict)
    depends_on: list[str] = field(default_factory=list)
    completed: bool = False

    @property
    def key(self) -> str:
        """Unique key for dependency resolution."""
        return f"{self.sheet or 'default'}:{self.id}"


# ---------------------------------------------------------------------------
# Helper: build a case-insensitive pattern from a raw string
# ---------------------------------------------------------------------------
def _p(pattern: str) -> re.Pattern:
    return re.compile(pattern, re.I)


# ---------------------------------------------------------------------------
# Regex patterns to identify task types from instruction text.
# Order matters only within a single type; first-match wins per type per line.
# ---------------------------------------------------------------------------
_PATTERNS: dict[TaskType, list[re.Pattern]] = {

    # ── Lookup / Dynamic-Array Functions (must precede generic FORMULA) ──
    TaskType.LOOKUP_FUNCTION: [
        _p(r"\bXLOOKUP\b"),
        _p(r"\bVLOOKUP\b"),
        _p(r"\bHLOOKUP\b"),
        _p(r"\bINDEX\s*\(.*MATCH"),
        _p(r"INDEX\s*/\s*MATCH"),  # "INDEX/MATCH" phrasing
        _p(r"(?:enter|use|create|type|add)\s+(?:the\s+|a\s+|an\s+)?(?:XLOOKUP|VLOOKUP|HLOOKUP)\b"),
    ],
    TaskType.FILTER_FUNCTION: [
        _p(r"=\s*FILTER\s*\("),
        _p(r"\bFILTER\s+function\b"),
        _p(r"(?:enter|use|create)\s+(?:the\s+|a\s+)?FILTER\b"),
    ],
    TaskType.SORT_FUNCTION: [
        _p(r"=\s*SORT\s*\("),
        _p(r"\bSORT\s+function\b"),
        _p(r"(?:enter|use|create)\s+(?:the\s+|a\s+)?SORT\b(?!\s+(?:the|data|range|table|by|ascending|descending))"),
    ],
    TaskType.UNIQUE_FUNCTION: [
        _p(r"=\s*UNIQUE\s*\("),
        _p(r"\bUNIQUE\s+function\b"),
        _p(r"(?:enter|use|create)\s+(?:the\s+|a\s+)?UNIQUE\b"),
    ],
    TaskType.TEXT_FUNCTION: [
        _p(r"=\s*(?:CONCAT|CONCATENATE|TEXTJOIN|LEFT|RIGHT|MID|UPPER|LOWER|PROPER|TRIM|LEN|SUBSTITUTE|TEXT)\s*\("),
        _p(r"\b(?:CONCAT|CONCATENATE|TEXTJOIN|LEFT|RIGHT|MID|UPPER|LOWER|PROPER|TRIM|LEN|SUBSTITUTE)\s+function\b"),
    ],

    # ── 3-D & External References ──
    TaskType.THREE_D_REFERENCE: [
        _p(r"3-?D\s+(?:reference|formula|cell)"),
        _p(r"(?:across|through|multiple)\s+(?:work)?sheets?\b"),
        _p(r"\w+:\w+!"),  # Sheet1:Sheet3!A1 pattern
    ],
    TaskType.EXTERNAL_REFERENCE: [
        _p(r"external\s+(?:reference|link|data)"),
        _p(r"\[.+?\.xls[xm]?\]"),  # [OtherWorkbook.xlsx]
        _p(r"(?:link|reference)\s+(?:to|from)\s+(?:another|external|different)\s+workbook"),
        _p(r"(?:link|connect|pull|import)\s+(?:to|from|in)\s+(?:the\s+)?\w+\s+workbook"),
    ],

    # ── Cell Value Entry ──
    TaskType.CELL_VALUE: [
        _p(r"(?:enter|type|input)\s+(?:the\s+)?(?:value\s+)?(?:\d[\d,.]*|\"[^\"]+\")\s+(?:in|into)\s+(?:the\s+)?cell"),
        _p(r"(?:in|into)\s+(?:the\s+)?cell\s+[A-Z]{1,3}\d{1,7},?\s*(?:enter|type|input)\s+(?:the\s+)?(?:value\s+)?(?:\d|\"|\w)"),
        _p(r"(?:enter|type|input|put|place|write)\s+(?:\d[\d,.]*)\s+in\s+(?:cell\s+)?[A-Z]{1,3}\d{1,7}"),
    ],

    # ── Generic Formulas & Functions ──
    TaskType.FORMULA: [
        _p(r"(?:enter|type|input|create|add|use)\s+(?:the\s+)?(?:a\s+)?formula\b"),
        _p(r"(?:enter|type|input|create|add|use)\s+(?:the\s+)?(?:a\s+)?function\b"),
        _p(r"(?:enter|type|input|create|add|use)\s+(?:the\s+)?(?:a\s+)?\w+\s+function\b"),
        _p(r"(?:enter|type|create|add|use)\s+(?:the\s+)?(?:an?\s+)?\w+\s+formula\b"),  # P2-13: "enter a SUM formula"
        _p(r"using\s+(?:the\s+)?\w+\s+function\b"),  # P0-3: "using the X function"
        _p(r"\b(?:TODAY|NOW|DATE|YEAR|MONTH|HOUR|MINUTE|SECOND)\s+function\b"),  # P0-3: date/time function names
        _p(r"(?:in\s+cell\s+[A-Z]{1,3}\d{1,7}),?\s*(?:enter|type)\s+="),
        _p(r"=(?:SUM|AVERAGE|COUNT|COUNTA|COUNTIF|COUNTIFS|SUMIF|SUMIFS|AVERAGEIF|AVERAGEIFS)\s*\("),
        _p(r"=(?:IF|IFS|AND|OR|NOT|IFERROR|IFNA)\s*\("),
        _p(r"=(?:MIN|MAX|MEDIAN|MODE|STDEV|VAR|LARGE|SMALL|RANK|PERCENTILE)\s*\("),
        _p(r"=(?:ROUND|ROUNDUP|ROUNDDOWN|INT|ABS|MOD|POWER|SQRT)\s*\("),
        _p(r"=(?:TODAY|NOW|DATE|YEAR|MONTH|DAY|HOUR|MINUTE|SECOND|DATEDIF|EDATE|EOMONTH|NETWORKDAYS|WORKDAY)\s*\("),
        _p(r"=(?:PMT|FV|PV|NPER|RATE|NPV|IRR)\s*\("),
        _p(r"=(?:SUBTOTAL)\s*\("),
        _p(r"=\w+\("),  # any =FUNC( pattern
        _p(r"=\s*[A-Z]{1,3}\d+\s*[+\-*/]"),  # =A1+B1 arithmetic
    ],

    # ── Tables ──
    TaskType.TABLE_CREATE: [
        _p(r"(?:create|format\s+as|convert\s+to|make\s+into)\s+(?:a\s+)?(?:an?\s+)?(?:Excel\s+)?table\b"),
        _p(r"format\s+(?:the\s+)?(?:range|data|cells)\s+as\s+(?:a\s+)?table\b"),
    ],
    TaskType.TABLE_STYLE: [
        _p(r"(?:apply|use|change|set)\s+(?:the\s+)?(?:table\s+)?style\s+\w+"),
        _p(r"TableStyle\w+"),
    ],
    TaskType.TABLE_TOTAL_ROW: [
        _p(r"total\s*row"),
        _p(r"(?:add|show|enable|turn\s+on|display)\s+(?:the\s+)?totals?\s+row"),
    ],
    TaskType.CALCULATED_COLUMN: [
        _p(r"calculated\s+column"),
        _p(r"\[@\w+\]"),        # structural reference [@Column]
        _p(r"\[@\[.+?\]\]"),    # structural reference [@[Column Name]]
        _p(r"structural\s+reference"),
    ],

    # ── Conditional Formatting ──
    TaskType.CONDITIONAL_FORMAT: [
        _p(r"conditional\s+format"),
        _p(r"highlight\s+cells?\s+(?:rules?|that|greater|less|between|equal|containing)"),
        _p(r"(?:color|colour)\s+scale"),
        _p(r"data\s+bars?"),
        _p(r"icon\s+sets?"),
        _p(r"top\s*/?bottom\s+(?:\d+|rules?)"),
        _p(r"(?:above|below)\s+average\s+(?:rule|formatting|highlight)"),
        _p(r"duplicate\s+values?\s+(?:rule|formatting|highlight)"),
        _p(r"new\s+(?:formatting\s+)?rule"),
        # Conditional instruction patterns (if/when value conditions)
        _p(r"if\s+(?:the\s+)?value\s+in\s+cell\s+[A-Z]{1,3}\d{1,7}\s+is\s+(?:greater|less|equal|not)"),
        _p(r"when\s+(?:the\s+)?(?:total|value|sum|count)\s+(?:exceeds?|is\s+(?:greater|less|over|under))"),
        _p(r"if\s+(?:the\s+)?(?:total|value|sum|number|count).+?format\s+(?:it\s+)?(?:in|as|with)\s+\w+"),  # P2-12: "if total exceeds 1000, format in red"
    ],

    # ── Number Format ──
    TaskType.NUMBER_FORMAT: [
        _p(r"(?:format|change)\s+(?:the\s+)?(?:cells?\s+)?(?:as|to|with)\s+(?:currency|accounting|percentage|percent|number|date|time|text|comma|scientific|fraction|general)\b"),
        _p(r"number\s+format"),
        _p(r"(?:apply|use)\s+(?:the\s+)?(?:Accounting|Currency|Percentage|Percent|Number|Comma|Scientific)\s+(?:format|style)"),
        _p(r"(?:increase|decrease)\s+(?:decimal|indent)"),
        _p(r"\b\d+\s+decimal\s+places?\b"),
    ],

    # ── Alignment ──
    TaskType.ALIGNMENT: [
        _p(r"(?:center|left|right)\s+align"),
        _p(r"(?:align|alignment)\s+(?:to\s+)?(?:center|left|right|top|bottom|middle)"),
        _p(r"wrap\s+text"),
        _p(r"text\s+wrapp(?:ing)?\b"),  # P1-6: "text wrapping"
        _p(r"(?:enable|disable|turn\s+on|turn\s+off)\s+(?:text\s+)?wrapp?(?:ing)?\b"),  # P1-6: "enable wrapping"
        _p(r"merge\s+(?:and\s+)?center"),
        _p(r"(?:horizontal|vertical)\s+(?:alignment|centering)"),
        _p(r"(?:indent|orientation|text\s+direction|shrink\s+to\s+fit)"),
        _p(r"rotate\s+(?:text|cell)"),
    ],

    # ── Column Width / Row Height ──
    TaskType.COLUMN_WIDTH: [
        _p(r"(?:change|set|adjust|resize)\s+(?:the\s+)?column\s+width"),
        _p(r"autofit\s+(?:column|width)"),
        _p(r"column\s+width\s+(?:to\s+)?\d+"),
        _p(r"(?:widen|narrow)\s+(?:the\s+)?column"),
        _p(r"column\s+[A-Z]{1,3}\s+(?:to\s+)?(?:a\s+)?width\s+(?:of\s+)?\d+"),
        _p(r"(?:change|set|adjust)\s+(?:the\s+)?width\s+of\s+(?:the\s+)?column"),
        _p(r"width\s+of\s+column\s+[A-Z]{1,3}\s+to\s+\d+"),
        _p(r"(?:hide|unhide|show)\s+(?:the\s+)?columns?\b"),  # P0-1: hide/unhide columns
        _p(r"auto\s*fit\s+(?:all\s+)?(?:the\s+)?columns?\b"),  # P0-5: "autofit all columns"
        _p(r"auto\s*fit\b"),  # P0-5: bare "autofit"
    ],
    TaskType.ROW_HEIGHT: [
        _p(r"(?:change|set|adjust|resize)\s+(?:the\s+)?row\s+height"),
        _p(r"autofit\s+(?:row|height)"),
        _p(r"row\s+height\s+(?:to\s+)?\d+"),
        _p(r"(?:hide|unhide|show)\s+(?:the\s+)?rows?\b"),  # P0-1: hide/unhide rows
    ],

    # ── View & Layout ──
    TaskType.FREEZE_PANES: [
        _p(r"freeze\s+(?:the\s+)?(?:top\s+)?(?:row|panes?|column)"),
        _p(r"freeze\s+(?:at|from)\s+(?:cell\s+)?[A-Z]{1,3}\d+"),
    ],
    TaskType.SPLIT_PANES: [
        _p(r"split\s+(?:the\s+)?(?:window|panes?)"),
    ],
    TaskType.PAGE_BREAK: [
        _p(r"page\s+break"),
        _p(r"(?:insert|add|remove)\s+(?:a\s+)?(?:horizontal|vertical)?\s*page\s*break"),
    ],

    # ── Data Tools ──
    TaskType.AUTOFILTER: [
        _p(r"auto\s*filter"),
        _p(r"(?:apply|add|enable|turn\s+on)\s+(?:a\s+)?filter(?:s|\b)"),
        _p(r"filter\s+(?:the\s+)?(?:data|range|table|records?)\b"),  # P0-4: "filter the data"
    ],
    TaskType.ADVANCED_FILTER: [
        _p(r"advanced\s+filter"),
        _p(r"(?:filter|extract)\s+(?:to|into)\s+(?:a\s+)?(?:different|another|separate)\s+(?:location|range)"),
    ],
    TaskType.SORT: [
        _p(r"sort\s+(?:the\s+)?(?:data|range|table|column|rows?)"),
        _p(r"sort\s+(?:ascending|descending|by)\b"),
        _p(r"(?:ascending|descending)\s+(?:order|sort)\b"),
        _p(r"(?:sort|arrange|order)\s+(?:by|on)\s+(?:the\s+)?(?:column|field)\b"),
        _p(r"custom\s+sort"),
    ],
    TaskType.SUBTOTAL: [
        _p(r"\bsubtotal\b(?!\s*\()"),  # the word subtotal but not =SUBTOTAL(
        _p(r"(?:add|insert|create)\s+(?:a\s+)?subtotals?\b"),
        _p(r"(?:group|outline)\s+(?:and\s+)?subtotals?\b"),
    ],
    TaskType.DATA_VALIDATION: [
        _p(r"data\s+validation"),
        _p(r"(?:drop\s*-?\s*down|dropdown)\s+list"),
        _p(r"(?:restrict|limit|validate)\s+(?:the\s+)?(?:input|entry|data|values?)"),
        _p(r"input\s+message"),
        _p(r"error\s+alert"),
        _p(r"(?:remove|delete|eliminate)\s+(?:the\s+)?duplicates?\b"),  # P1-7: "remove duplicates"
    ],
    TaskType.GOAL_SEEK: [
        _p(r"goal\s+seek"),
        _p(r"what.?if\s+analysis"),
        _p(r"scenario\s+manager"),
    ],

    # ── Charts ──
    TaskType.CHART_BAR: [
        _p(r"(?:create|insert|add)\s+(?:a\s+)?(?:bar|column)\s+chart"),
        _p(r"(?:clustered|stacked|100%?\s+stacked)\s+(?:bar|column)"),
        _p(r"(?:2-?D|3-?D)\s+(?:bar|column)\s+chart"),
        _p(r"(?:change|switch)\s+(?:the\s+)?chart\s+(?:type\s+)?to\s+(?:a\s+)?(?:bar|column)\b"),  # P2-11
        _p(r"(?:add|change|format|remove)\s+(?:a\s+|the\s+)?(?:chart\s+)?(?:title|trendline|data\s+labels?|legend|axis)"),  # P1-8: chart modifications
    ],
    TaskType.CHART_LINE: [
        _p(r"(?:create|insert|add)\s+(?:a\s+)?line\s+chart"),
        _p(r"(?:line\s+with\s+markers?|stacked\s+line)"),
        _p(r"(?:change|switch)\s+(?:the\s+)?chart\s+(?:type\s+)?to\s+(?:a\s+)?line\b"),  # P2-11
    ],
    TaskType.CHART_PIE: [
        _p(r"(?:create|insert|add)\s+(?:a\s+)?(?:pie|doughnut)\s+chart"),
        _p(r"(?:3-?D\s+pie|exploded\s+pie|pie\s+of\s+pie)"),
        _p(r"(?:change|switch)\s+(?:the\s+)?chart\s+(?:type\s+)?to\s+(?:a\s+)?(?:pie|doughnut)\b"),  # P2-11
    ],
    TaskType.CHART_SCATTER: [
        _p(r"(?:create|insert|add)\s+(?:a\s+)?(?:scatter|XY)\s+chart"),
        _p(r"(?:scatter|XY)\s+(?:chart|plot|with\s+)"),
        _p(r"(?:scatter\s+with\s+)?(?:smooth\s+lines?|straight\s+lines?)"),
        _p(r"(?:change|switch)\s+(?:the\s+)?chart\s+(?:type\s+)?to\s+(?:a\s+)?(?:scatter|XY)\b"),  # P2-11
    ],
    TaskType.CHART_AREA: [
        _p(r"(?:create|insert|add)\s+(?:a\s+|an\s+)?area\s+chart"),
        _p(r"(?:stacked\s+area|100%?\s+stacked\s+area)"),
        _p(r"area\s+chart"),
        _p(r"(?:change|switch)\s+(?:the\s+)?chart\s+(?:type\s+)?to\s+(?:a\s+|an\s+)?area\b"),  # P2-11
    ],
    TaskType.CHART_COMBO: [
        _p(r"combo\s+chart"),
        _p(r"(?:combination|mixed)\s+chart"),
        _p(r"secondary\s+(?:axis|y-?axis)"),
        _p(r"(?:change|switch)\s+(?:the\s+)?chart\s+(?:type\s+)?to\s+(?:a\s+)?combo\b"),  # P2-11
    ],
    TaskType.CHART_HISTOGRAM: [
        _p(r"histogram"),
        _p(r"(?:frequency|distribution)\s+chart"),
        _p(r"(?:bin\s+width|overflow\s+bin|underflow\s+bin)"),
    ],
    TaskType.SPARKLINE: [
        _p(r"sparkline"),
        _p(r"(?:insert|add|create)\s+(?:a\s+)?(?:line|column|win/loss)\s+sparkline"),
    ],

    # ── Ranges & References ──
    TaskType.NAMED_RANGE: [
        _p(r"(?:create|define|name|add)\s+(?:a\s+)?(?:named\s+)?range"),
        _p(r"name\s+(?:manager|box)"),
        _p(r"(?:assign|give)\s+(?:the\s+)?name\b"),
    ],
    TaskType.HYPERLINK: [
        _p(r"hyperlink"),
        _p(r"(?:insert|add|create)\s+(?:a\s+)?(?:hyper)?link"),
        _p(r"link\s+to\s+(?:a\s+)?(?:web|email|file|sheet|cell)"),
    ],

    # ── Advanced Features ──
    TaskType.SLICER: [
        _p(r"(?:insert|create|add)\s+(?:a\s+)?slicers?\b"),
        _p(r"\bslicer\b"),
    ],
    TaskType.PIVOT_TABLE: [
        _p(r"pivot\s*table"),
        _p(r"(?:create|insert|add)\s+(?:a\s+)?pivot"),
    ],
    TaskType.PIVOT_CHART: [
        _p(r"pivot\s*chart"),
    ],

    # ── Sheet Operations ──
    TaskType.SHEET_CREATE: [
        _p(r"(?:insert|add|create)\s+(?:a\s+)?(?:new\s+)?(?:work)?sheet"),
    ],
    TaskType.SHEET_RENAME: [
        _p(r"rename\s+(?:the\s+)?(?:work)?sheet"),
        _p(r"(?:change|set)\s+(?:the\s+)?(?:sheet|tab)\s+name"),
    ],
    TaskType.SHEET_MOVE: [
        _p(r"(?:move|reorder)\s+(?:the\s+)?(?:work)?sheet"),
    ],
    TaskType.SHEET_COPY: [
        _p(r"(?:copy|duplicate)\s+(?:the\s+)?(?:work)?sheet"),
        _p(r"(?:copy|duplicate)\s+(?:the\s+)?[\w\s]+?\s+(?:sheet|worksheet|tab)"),
    ],
    TaskType.TAB_COLOR: [
        _p(r"(?:change|set|apply)\s+(?:the\s+)?(?:sheet\s+)?tab\s+colo[u]?r"),
    ],

    # ── Formatting ──
    TaskType.FONT: [
        _p(r"(?:change|set|apply)\s+(?:the\s+)?font"),
        _p(r"(?:make|format)\s+(?:the\s+)?(?:text|cell|range|selection)\s+(?:bold|italic|underline)"),
        _p(r"(?:format|make)\s+.*?\bas\s+(?:bold|italic|underline)\b"),  # P0-2: "format X as bold"
        _p(r"\bas\s+(?:bold|italic|underline|strikethrough)\b"),  # P0-2: bare "as bold"
        _p(r"font\s+(?:size|color|colour|name|face)"),
        _p(r"\b(?:bold|italic|underline|strikethrough)\s+(?:the|to)\b"),
        _p(r"(?:apply|use)\s+(?:the\s+)?[\w\s]+?\s+style\s+(?:to|for)\b"),  # P1-10: cell style names
    ],
    TaskType.FILL: [
        _p(r"(?:fill|background|cell)\s+colo[u]?r"),
        _p(r"(?:shade|highlight)\s+(?:the\s+)?cell"),
        _p(r"(?:apply|add|set)\s+(?:a\s+)?(?:fill|shading)\b"),
    ],
    TaskType.BORDER: [
        _p(r"(?:add|apply|draw|set)\s+(?:a\s+)?borders?\b"),
        _p(r"(?:outside|inside|bottom|top|left|right|all|thick|thin|double)\s+borders?\b"),
    ],
    TaskType.MERGE_CELLS: [
        _p(r"merge\s+(?:and\s+center\s+)?cells?"),
        _p(r"(?:unmerge|split)\s+cells?"),
    ],

    # ── File Operations ──
    TaskType.SAVE: [
        _p(r"save\s+(?:the\s+)?(?:work)?book\b"),
        _p(r"(?:press|use)\s+(?:Ctrl|Cmd)\s*\+?\s*S\b"),
    ],
    TaskType.SAVE_AS: [
        _p(r"save\s+(?:the\s+)?(?:work)?book\s+as\b"),
    ],
    TaskType.PRINT_SETTINGS: [
        _p(r"(?:set|change|adjust)\s+(?:the\s+)?(?:print|page)\s+(?:area|orientation|margins?|header|footer|setup|layout|scaling|title)"),
        _p(r"(?:landscape|portrait)\s+orientation"),
        _p(r"(?:print\s+)?(?:page\s+)?(?:header|footer)\s+(?:and\s+)?(?:header|footer)"),  # P1-9: fixed false positive
        _p(r"(?:print\s+titles?|repeat\s+(?:rows?|columns?)\s+(?:at|on))"),
        _p(r"(?:fit\s+to|scale\s+to)\s+(?:\d+\s+)?page"),
        _p(r"(?:add|insert|create)\s+(?:a\s+)?(?:header|footer)s?\b"),  # P1-9: headers/footers for printing
    ],
}


# ---------------------------------------------------------------------------
# Reference extraction patterns
# ---------------------------------------------------------------------------

# Cell: A1, AB123, XFD1048576
_CELL_REF = re.compile(r"\b([A-Z]{1,3}\d{1,7})\b")
# Range: A1:B10
_RANGE_REF = re.compile(r"\b([A-Z]{1,3}\d{1,7}:[A-Z]{1,3}\d{1,7})\b")

# Sheet name extraction — handles many SAM phrasings:
#   "on the Sales sheet", "in worksheet 'Q1 Data'", "Go to the Dashboard sheet",
#   "switch to Employees", "the Wages worksheet"
_SHEET_REF = re.compile(
    r"(?:"
    # Pattern 1: "on/in/go to the sheet Sales", "go to sheet Summary"
    r"(?:on|in|of|to|from|go\s+to|switch\s+to|navigate\s+to|select)\s+"
    r"(?:the\s+)?"
    r"(?:sheet|worksheet|tab)\s+"
    r"[\"']?([A-Za-z][\w\s\-]*?)[\"']?"
    r"|"
    # Pattern 2: "on the 'Q1 Sales' worksheet" (quoted name before sheet word)
    r"(?:on|in|of|to|from|go\s+to|switch\s+to)\s+"
    r"(?:the\s+)?"
    r"[\"']([A-Za-z][\w\s\-]+?)[\"']"
    r"\s+(?:sheet|worksheet|tab)"
    r"|"
    # Pattern 3: "the 'Revenue' tab" (quoted name before sheet word, no preposition)
    r"(?:the\s+)?"
    r"[\"']([A-Za-z][\w\s\-]+?)[\"']"
    r"\s+(?:sheet|worksheet|tab)"
    r"|"
    # Pattern 4: "Go to the Revenue worksheet" (unquoted name before sheet word)
    r"(?:on|in|of|to|from|go\s+to|switch\s+to|navigate\s+to)\s+"
    r"(?:the\s+)?"
    r"([A-Z][A-Za-z0-9]+)"
    r"\s+(?:sheet|worksheet|tab)"
    r")"
    r"(?:\s|,|\.|;|$)",
    re.I,
)

# Formula extraction — handles nested parens, arithmetic, no-paren formulas
_FORMULA_REF = re.compile(
    r"(="
    r"(?:"
    r"[A-Z]+\([^)]*(?:\([^)]*\)[^)]*)*\)"  # =FUNC(... possibly nested ...)
    r"|"
    r"[A-Z]{1,3}\d+\s*[+\-*/].+?(?=\s+(?:in|into|on|then|and|for)\b|\s*[,;]|$)"  # =A1+B1*C1
    r"|"
    r"[A-Z]+\(.+?\)"                          # =FUNC(simple args)
    r")"
    r")",
    re.I,
)

# Value extraction: "enter the value 42" / "type Michael Manthe"
_VALUE_REF = re.compile(
    r"(?:enter|type|input)\s+(?:the\s+)?(?:value\s+)?[\"'](.+?)[\"']",
    re.I,
)

# Numeric value: "width to 20" / "height of 15.5"
_NUMERIC_REF = re.compile(
    r"(?:to|of|=)\s+(\d+(?:\.\d+)?)\b",
)


class TaskExtractor:
    """Extract structured Task objects from instruction text."""

    def __init__(self) -> None:
        self._counter = 0

    def extract(self, text: str) -> list[Task]:
        """
        Extract all tasks from instruction text.
        Splits text into instruction lines/paragraphs and identifies tasks.
        """
        self._counter = 0
        tasks: list[Task] = []

        lines = self._split_instructions(text)

        for line in lines:
            line_tasks = self._extract_from_line(line)
            tasks.extend(line_tasks)

        self._resolve_dependencies(tasks)

        logger.info("Extracted %d tasks from instructions", len(tasks))
        return tasks

    def extract_from_steps(self, steps: list[InstructionStep]) -> list[Task]:
        """
        Extract tasks from pre-parsed InstructionStep objects.

        Uses each step's ``sheet_context`` as the default sheet when the
        step text does not contain an explicit sheet reference.  Processes
        steps in order, maintaining context, and returns tasks with proper
        sheet assignments.  Complements (does not replace) ``extract()``.
        """
        self._counter = 0
        tasks: list[Task] = []

        for step in steps:
            line_tasks = self._extract_from_line(step.text)
            for task in line_tasks:
                # Inherit sheet from step context when not already set
                if not task.sheet and step.sheet_context:
                    task.sheet = step.sheet_context
                # Stash parent_step reference for traceability
                if step.parent_step is not None:
                    task.params["parent_step"] = step.parent_step
                task.params["step_number"] = step.step_number
            tasks.extend(line_tasks)

        self._resolve_dependencies(tasks)
        logger.info("Extracted %d tasks from %d steps", len(tasks), len(steps))
        return tasks

    def _next_id(self) -> str:
        """Generate a sequential task ID."""
        self._counter += 1
        return f"task_{self._counter:03d}"

    def _split_instructions(self, text: str) -> list[str]:
        """Split instruction text into individual instruction lines."""
        # Pre-strip section header lines so they don't merge with steps
        text = re.sub(
            r"\n\s*(Section\s+\d+[:.][^\n]*)",
            r"\n\n\1",
            text,
            flags=re.I,
        )

        # Try numbered steps first:  1. / Step 1: / a. / a) / • / -
        numbered = re.split(
            r"\n\s*(?:\d+[.)]\s+|Step\s+\d+[:.]\s+|Section\s+\d+[:.]\s*|[a-z][.)]\s+|[•●]\s+|-\s+)",
            text,
            flags=re.I,
        )
        if len(numbered) > 2:
            return [s.strip() for s in numbered if s.strip()]

        # Fall back to line-by-line with minimum length
        paragraphs = text.split("\n")
        result = []
        for p in paragraphs:
            p = p.strip()
            if p and len(p) > 4:
                result.append(p)
        return result

    def _extract_from_line(self, line: str) -> list[Task]:
        """Extract tasks from a single instruction line."""
        tasks = []
        matched_types: set[TaskType] = set()

        for task_type, patterns in _PATTERNS.items():
            if task_type in matched_types:
                continue
            for pattern in patterns:
                if pattern.search(line):
                    task = self._build_task(task_type, line)
                    tasks.append(task)
                    matched_types.add(task_type)
                    break  # one match per task type per line

        # Fallback: line contains a formula literal but no pattern matched
        if not tasks:
            formula_match = _FORMULA_REF.search(line)
            if formula_match:
                task = self._build_task(TaskType.FORMULA, line)
                task.formula = formula_match.group(1)
                tasks.append(task)

        # Suppress spurious CELL_VALUE when a formula-type task already exists
        _FORMULA_TYPES = {
            TaskType.FORMULA,
            TaskType.LOOKUP_FUNCTION,
            TaskType.FILTER_FUNCTION,
            TaskType.SORT_FUNCTION,
            TaskType.UNIQUE_FUNCTION,
            TaskType.TEXT_FUNCTION,
            TaskType.THREE_D_REFERENCE,
        }
        has_formula_task = any(t.task_type in _FORMULA_TYPES for t in tasks)
        if has_formula_task:
            tasks = [t for t in tasks if t.task_type != TaskType.CELL_VALUE]

        return tasks

    def _build_task(self, task_type: TaskType, line: str) -> Task:
        """Build a Task object from a matched line."""
        task = Task(
            id=self._next_id(),
            task_type=task_type,
            description=line[:300],
        )

        # ── Extract sheet reference ──
        sheet_match = _SHEET_REF.search(line)
        if sheet_match:
            # Take the first non-None capture group
            task.sheet = next(
                (g.strip() for g in sheet_match.groups() if g), None
            )

        # ── Extract target cell from "in cell X" / "into cell X" phrasing ──
        target_match = re.search(
            r"(?:in|into)\s+(?:the\s+)?(?:cell\s+)?([A-Z]{1,3}\d{1,7})\b",
            line,
            re.I,
        )

        # ── Extract cell / range reference ──
        range_match = _RANGE_REF.search(line)
        if range_match:
            task.range = range_match.group(1)
        else:
            cell_match = _CELL_REF.search(line)
            if cell_match:
                task.cell = cell_match.group(1)

        # If a target cell was explicitly stated, prefer it over the range-
        # derived fallback (which may have come from inside a formula).
        if target_match:
            task.cell = target_match.group(1).upper()

        # ── Extract formula ──
        formula_match = _FORMULA_REF.search(line)
        if formula_match:
            task.formula = formula_match.group(1)

        # ── Extract table style ──
        style_match = re.search(r"(TableStyle\w+\d+)", line)
        if style_match:
            task.style = style_match.group(1)

        # ── Extract numeric params (width, height, etc.) ──
        if task_type in (TaskType.COLUMN_WIDTH, TaskType.ROW_HEIGHT):
            num_match = _NUMERIC_REF.search(line)
            if num_match:
                task.params["size"] = float(num_match.group(1))

        # ── Extract value ──
        val_match = _VALUE_REF.search(line)
        if val_match:
            task.value = val_match.group(1)

        # ── Extract bare numeric value for CELL_VALUE tasks ──  (P2-14)
        if task_type == TaskType.CELL_VALUE and not task.value:
            bare_num = re.search(
                r"(?:enter|type|input|put)\s+(?:the\s+)?(?:number\s+)?(\d[\d,.]*)",
                line, re.I,
            )
            if bare_num:
                task.value = bare_num.group(1)

        return task

    def _resolve_dependencies(self, tasks: list[Task]) -> None:
        """
        Auto-detect task dependencies based on ordering rules:
          - Table operations depend on table creation on the same sheet
          - Calculated columns depend on table creation
          - Charts depend on their data being present
          - Slicers depend on tables or pivot tables
          - PivotChart depends on PivotTable
          - Dynamic-array functions may depend on named ranges
          - Save depends on all other tasks
        """
        table_creates = [t for t in tasks if t.task_type == TaskType.TABLE_CREATE]
        pivot_creates = [t for t in tasks if t.task_type == TaskType.PIVOT_TABLE]
        named_ranges = [t for t in tasks if t.task_type == TaskType.NAMED_RANGE]

        _TABLE_DEPENDENT = {
            TaskType.TABLE_STYLE, TaskType.TABLE_TOTAL_ROW,
            TaskType.CALCULATED_COLUMN,
        }
        _CHART_TYPES = {
            TaskType.CHART_BAR, TaskType.CHART_LINE, TaskType.CHART_PIE,
            TaskType.CHART_SCATTER, TaskType.CHART_AREA, TaskType.CHART_COMBO,
            TaskType.CHART_HISTOGRAM,
        }
        _DYNAMIC_ARRAY = {
            TaskType.FILTER_FUNCTION, TaskType.SORT_FUNCTION,
            TaskType.UNIQUE_FUNCTION,
        }

        for task in tasks:
            # Table-dependent operations
            if task.task_type in _TABLE_DEPENDENT:
                for tc in table_creates:
                    if tc.sheet == task.sheet or task.sheet is None:
                        task.depends_on.append(tc.id)

            # Slicer depends on table or pivot
            if task.task_type == TaskType.SLICER:
                for tc in table_creates + pivot_creates:
                    task.depends_on.append(tc.id)

            # PivotChart depends on PivotTable
            if task.task_type == TaskType.PIVOT_CHART:
                for pc in pivot_creates:
                    task.depends_on.append(pc.id)

            # Sparklines depend on data being present (formulas on same sheet)
            if task.task_type == TaskType.SPARKLINE:
                for other in tasks:
                    if (other.task_type == TaskType.FORMULA
                            and other.sheet == task.sheet
                            and other.id != task.id):
                        task.depends_on.append(other.id)

            # Dynamic-array functions referencing table columns depend on tables
            if task.task_type in _DYNAMIC_ARRAY:
                for tc in table_creates:
                    if tc.sheet == task.sheet or task.sheet is None:
                        task.depends_on.append(tc.id)

            # XLOOKUP/VLOOKUP may reference named ranges
            if task.task_type == TaskType.LOOKUP_FUNCTION:
                for nr in named_ranges:
                    task.depends_on.append(nr.id)

            # Sort/subtotal operations depend on table or data being present
            if task.task_type in (TaskType.SORT, TaskType.SUBTOTAL):
                for tc in table_creates:
                    if tc.sheet == task.sheet or task.sheet is None:
                        task.depends_on.append(tc.id)

            # Save depends on all other tasks
            if task.task_type == TaskType.SAVE:
                for other in tasks:
                    if other.id != task.id and other.task_type != TaskType.SAVE:
                        task.depends_on.append(other.id)
