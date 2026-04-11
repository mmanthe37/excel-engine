package tools

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/mmanthe37/excel-engine/go-mcp-server/config"
)

var cfg *config.Config

// SetConfig stores the global config for all tool handlers.
func SetConfig(c *config.Config) {
	cfg = c
}

// --- Tool 1: CompleteAssignment ---

type CompleteAssignmentInput struct {
	WorkbookPath    string         `json:"workbook_path" jsonschema:"Path to the .xlsx workbook file"`
	InstructionPath string         `json:"instruction_path" jsonschema:"Path to the instruction file (.docx .pdf .txt)"`
	Options         map[string]any `json:"options,omitempty" jsonschema:"Optional engine config overrides (max_retries etc.)"`
}

func CompleteAssignmentHandler(ctx context.Context, input CompleteAssignmentInput) (string, error) {
	if input.WorkbookPath == "" || input.InstructionPath == "" {
		return "", fmt.Errorf("workbook_path and instruction_path are required")
	}
	if ctx.Err() != nil {
		return "", ctx.Err()
	}

	optionsJSON := "{}"
	if input.Options != nil {
		b, _ := json.Marshal(input.Options)
		optionsJSON = string(b)
	}

	script := fmt.Sprintf(`
import json
from pathlib import Path
from excel_engine import ExcelEngine, EngineConfig

wb = Path(%q).expanduser().resolve()
inst = Path(%q).expanduser().resolve()
options = json.loads(%q)

config = EngineConfig()
safe_keys = {"max_retries", "verify_after_each_section", "retina_display", "scan_timeout"}
for k, v in options.items():
    if k in safe_keys and hasattr(config, k):
        setattr(config, k, v)

engine = ExcelEngine(config=config)
result = engine.run(workbook=wb, instructions=inst)
print(json.dumps({
    "success": result.success,
    "workbook": str(result.workbook_path),
    "sections_completed": result.sections_completed,
    "sections_total": result.sections_total,
    "tasks_completed": result.tasks_completed,
    "tasks_total": result.tasks_total,
    "elapsed_seconds": round(result.elapsed_seconds, 2),
    "errors": result.errors,
    "summary": result.summary(),
}))
`, input.WorkbookPath, input.InstructionPath, optionsJSON)

	raw, err := pythonBridge(ctx, cfg, script)
	if err != nil {
		return "", err
	}
	return string(raw), nil
}

// --- Tool 2: ParseInstructions ---

type ParseInstructionsInput struct {
	InstructionPath string `json:"instruction_path" jsonschema:"Path to the instruction file"`
}

func ParseInstructionsHandler(ctx context.Context, input ParseInstructionsInput) (string, error) {
	if input.InstructionPath == "" {
		return "", fmt.Errorf("instruction_path is required")
	}
	if ctx.Err() != nil {
		return "", ctx.Err()
	}

	script := fmt.Sprintf(`
import json
from pathlib import Path
from excel_engine.parsers.instruction_parser import InstructionParser
from excel_engine.parsers.task_extractor import TaskExtractor

inst = Path(%q).expanduser().resolve()
parser = InstructionParser()
text = parser.parse(inst)

extractor = TaskExtractor()
tasks = extractor.extract(text)

print(json.dumps({
    "instruction_file": str(inst),
    "raw_text_preview": text[:500] + ("..." if len(text) > 500 else ""),
    "raw_text_length": len(text),
    "task_count": len(tasks),
    "tasks": [{"id": t.id, "task_type": t.task_type.value, "description": t.description,
               "sheet": t.sheet, "cell": t.cell, "range": t.range, "value": t.value,
               "formula": t.formula} for t in tasks],
}))
`, input.InstructionPath)

	raw, err := pythonBridge(ctx, cfg, script)
	if err != nil {
		return "", err
	}
	return string(raw), nil
}

// --- Tool 3: ExecuteOpenpyxl ---

type ExecuteOpenpyxlInput struct {
	WorkbookPath string   `json:"workbook_path" jsonschema:"Path to the .xlsx workbook"`
	Tasks        []string `json:"tasks" jsonschema:"List of task description strings"`
}

func ExecuteOpenpyxlHandler(ctx context.Context, input ExecuteOpenpyxlInput) (string, error) {
	if input.WorkbookPath == "" {
		return "", fmt.Errorf("workbook_path is required")
	}
	if len(input.Tasks) == 0 {
		return "", fmt.Errorf("at least one task is required")
	}
	if len(input.Tasks) > cfg.MaxTasks {
		return "", fmt.Errorf("too many tasks (max %d)", cfg.MaxTasks)
	}
	for _, t := range input.Tasks {
		if len(t) > cfg.MaxTaskLen {
			return "", fmt.Errorf("task too long (max %d chars)", cfg.MaxTaskLen)
		}
	}
	if ctx.Err() != nil {
		return "", ctx.Err()
	}

	tasksJSON, _ := json.Marshal(input.Tasks)
	script := fmt.Sprintf(`
import json
from pathlib import Path
from excel_engine import ExcelEngine, EngineConfig
from excel_engine.config import Layer
from excel_engine.parsers.task_extractor import TaskExtractor

wb = Path(%q).expanduser().resolve()
tasks_raw = json.loads(%q)

extractor = TaskExtractor()
task_text = "\n".join(f"- {t}" for t in tasks_raw)
task_objs = extractor.extract(task_text)

config = EngineConfig()
config.layer_order = [Layer.OPENPYXL]
engine = ExcelEngine(config=config)
result = engine.run(workbook=wb, tasks=task_objs)

print(json.dumps({
    "success": result.success,
    "tasks_completed": result.tasks_completed,
    "tasks_total": result.tasks_total,
    "elapsed_seconds": round(result.elapsed_seconds, 2),
    "errors": result.errors,
    "summary": result.summary(),
}))
`, input.WorkbookPath, string(tasksJSON))

	raw, err := pythonBridge(ctx, cfg, script)
	if err != nil {
		return "", err
	}
	return string(raw), nil
}

// --- Tool 4: VerifyWorkbook ---

type VerifyWorkbookInput struct {
	WorkbookPath  string   `json:"workbook_path" jsonschema:"Path to the .xlsx workbook"`
	ExpectedTasks []string `json:"expected_tasks,omitempty" jsonschema:"Optional list of task descriptions to verify"`
}

func VerifyWorkbookHandler(ctx context.Context, input VerifyWorkbookInput) (string, error) {
	if input.WorkbookPath == "" {
		return "", fmt.Errorf("workbook_path is required")
	}
	if ctx.Err() != nil {
		return "", ctx.Err()
	}

	tasksJSON := "null"
	if len(input.ExpectedTasks) > 0 {
		b, _ := json.Marshal(input.ExpectedTasks)
		tasksJSON = string(b)
	}

	script := fmt.Sprintf(`
import json
from pathlib import Path
from excel_engine.parsers.task_extractor import TaskExtractor
from excel_engine.verifier.workbook_verifier import WorkbookVerifier

wb = Path(%q).expanduser().resolve()
expected = json.loads(%q)

verifier = WorkbookVerifier()
verifier.load(wb)

if expected:
    extractor = TaskExtractor()
    task_text = "\n".join(f"- {t}" for t in expected)
    task_objs = extractor.extract(task_text)
    verification = verifier.verify_section("go-mcp-check", task_objs)
    verifier.close()
    print(json.dumps({
        "workbook": str(wb),
        "total_tasks": len(verification.results),
        "passed": verification.pass_count,
        "failed": verification.fail_count,
        "all_passed": verification.all_passed,
        "score": f"{verification.pass_count}/{len(verification.results)}",
        "results": [{"task_id": r.task_id, "task_type": r.task_type.value,
                     "passed": r.passed, "message": r.message} for r in verification.results],
    }))
else:
    from openpyxl import load_workbook
    oxwb = load_workbook(str(wb), data_only=False)
    report = {
        "workbook": str(wb),
        "sheets": oxwb.sheetnames,
        "sheet_count": len(oxwb.sheetnames),
        "formulas_found": 0,
    }
    for ws_name in oxwb.sheetnames:
        ws = oxwb[ws_name]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    report["formulas_found"] += 1
    oxwb.close()
    verifier.close()
    print(json.dumps(report))
`, input.WorkbookPath, tasksJSON)

	raw, err := pythonBridge(ctx, cfg, script)
	if err != nil {
		return "", err
	}
	return string(raw), nil
}

// --- Tool 5: GetEngineStatus ---

type GetEngineStatusInput struct{}

func GetEngineStatusHandler(ctx context.Context, input GetEngineStatusInput) (string, error) {
	if ctx.Err() != nil {
		return "", ctx.Err()
	}

	script := `
import json
from excel_engine import __version__
from excel_engine.config import Layer, TaskType, TASK_LAYER_MAP, EngineConfig

config = EngineConfig()
print(json.dumps({
    "engine_version": __version__,
    "go_mcp_server_version": "1.0.0",
    "available_layers": [{"number": l.value, "name": l.name} for l in Layer],
    "supported_task_types": [tt.value for tt in TaskType],
    "config": {
        "max_retries": config.max_retries,
        "layer_order": [l.name for l in config.layer_order],
        "verify_after_each_section": config.verify_after_each_section,
    },
}))
`
	raw, err := pythonBridge(ctx, cfg, script)
	if err != nil {
		return "", err
	}
	return string(raw), nil
}

// --- Tool 6: ExecuteLive ---

type ExecuteLiveInput struct {
	WorkbookPath string   `json:"workbook_path" jsonschema:"Path to the .xlsx workbook"`
	Tasks        []string `json:"tasks" jsonschema:"List of task description strings for live Excel operations"`
}

func ExecuteLiveHandler(ctx context.Context, input ExecuteLiveInput) (string, error) {
	if input.WorkbookPath == "" {
		return "", fmt.Errorf("workbook_path is required")
	}
	if len(input.Tasks) == 0 {
		return "", fmt.Errorf("at least one task is required")
	}
	if len(input.Tasks) > cfg.MaxTasks {
		return "", fmt.Errorf("too many tasks (max %d)", cfg.MaxTasks)
	}
	if ctx.Err() != nil {
		return "", ctx.Err()
	}

	tasksJSON, _ := json.Marshal(input.Tasks)

	// Build the task list joined with newlines
	taskLines := make([]string, len(input.Tasks))
	for i, t := range input.Tasks {
		taskLines[i] = fmt.Sprintf("- %s", t)
	}

	script := fmt.Sprintf(`
import json
from pathlib import Path
from excel_engine import ExcelEngine, EngineConfig
from excel_engine.config import Layer
from excel_engine.parsers.task_extractor import TaskExtractor

wb = Path(%q).expanduser().resolve()
tasks_raw = json.loads(%q)

extractor = TaskExtractor()
task_text = "\n".join(f"- {t}" for t in tasks_raw)
task_objs = extractor.extract(task_text)

config = EngineConfig()
config.layer_order = [Layer.XLWINGS, Layer.APPLESCRIPT, Layer.SYSTEM_EVENTS, Layer.VBA, Layer.PYAUTOGUI]
engine = ExcelEngine(config=config)
result = engine.run(workbook=wb, tasks=task_objs)

print(json.dumps({
    "success": result.success,
    "tasks_completed": result.tasks_completed,
    "tasks_total": result.tasks_total,
    "elapsed_seconds": round(result.elapsed_seconds, 2),
    "errors": result.errors,
    "summary": result.summary(),
}))
`, input.WorkbookPath, string(tasksJSON))

	raw, err := pythonBridge(ctx, cfg, script)
	if err != nil {
		return "", err
	}
	return string(raw), nil
}
