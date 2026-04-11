package tools

import (
	"context"
	"encoding/json"
	"strings"
	"testing"
	"time"

	"github.com/mmanthe37/excel-engine/go-mcp-server/config"
)

// testConfig returns a Config that uses python3 to run scripts.
func testConfig() *config.Config {
	return &config.Config{
		PythonPath: "python3",
		EngineRoot: "/nonexistent", // not needed for simple scripts
	}
}

func TestPythonBridge_Success(t *testing.T) {
	cfg := testConfig()
	script := `import json; print(json.dumps({"status": "ok", "value": 42}))`

	raw, err := pythonBridge(context.Background(), cfg, script)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(raw, &result); err != nil {
		t.Fatalf("invalid JSON returned: %v", err)
	}
	if result["status"] != "ok" {
		t.Errorf("expected status=ok, got %v", result["status"])
	}
	if result["value"] != float64(42) {
		t.Errorf("expected value=42, got %v", result["value"])
	}
}

func TestPythonBridge_EmptyOutput(t *testing.T) {
	cfg := testConfig()
	// Script that produces no stdout
	script := `pass`

	raw, err := pythonBridge(context.Background(), cfg, script)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(raw, &result); err != nil {
		t.Fatalf("invalid JSON: %v", err)
	}
	if _, ok := result["error"]; !ok {
		t.Error("expected error key in empty-response JSON")
	}
}

func TestPythonBridge_InvalidJSON(t *testing.T) {
	cfg := testConfig()
	script := `print("not json at all")`

	_, err := pythonBridge(context.Background(), cfg, script)
	if err == nil {
		t.Fatal("expected error for invalid JSON output")
	}
	if !strings.Contains(err.Error(), "invalid JSON") {
		t.Errorf("expected 'invalid JSON' in error, got: %v", err)
	}
}

func TestPythonBridge_ScriptFailure(t *testing.T) {
	cfg := testConfig()
	script := `raise ValueError("boom")`

	_, err := pythonBridge(context.Background(), cfg, script)
	if err == nil {
		t.Fatal("expected error for script that raises exception")
	}
	if !strings.Contains(err.Error(), "python") {
		t.Errorf("expected 'python' in error message, got: %v", err)
	}
}

func TestPythonBridge_NonZeroExit(t *testing.T) {
	cfg := testConfig()
	script := `import sys; sys.exit(1)`

	_, err := pythonBridge(context.Background(), cfg, script)
	if err == nil {
		t.Fatal("expected error for non-zero exit")
	}
}

func TestPythonBridge_ContextCancellation(t *testing.T) {
	cfg := testConfig()
	// Script that sleeps long enough for cancellation to fire
	script := `import time; time.sleep(30); import json; print(json.dumps({"done": True}))`

	ctx, cancel := context.WithCancel(context.Background())
	// Cancel almost immediately
	go func() {
		time.Sleep(100 * time.Millisecond)
		cancel()
	}()

	_, err := pythonBridge(ctx, cfg, script)
	if err == nil {
		t.Fatal("expected error after context cancellation")
	}
}

func TestPythonBridge_Timeout(t *testing.T) {
	// Use a config with python3 but the bridge sets a 120s timeout internally;
	// we test with a tighter parent context to simulate timeout behavior.
	cfg := testConfig()
	script := `import time; time.sleep(30)`

	ctx, cancel := context.WithTimeout(context.Background(), 200*time.Millisecond)
	defer cancel()

	_, err := pythonBridge(ctx, cfg, script)
	if err == nil {
		t.Fatal("expected error from timeout")
	}
}

func TestResolvePath_Empty(t *testing.T) {
	_, err := resolvePath("")
	if err == nil {
		t.Fatal("expected error for empty path")
	}
}

func TestResolvePath_Absolute(t *testing.T) {
	p, err := resolvePath("/some/file.xlsx")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if p != "/some/file.xlsx" {
		t.Errorf("expected /some/file.xlsx, got %s", p)
	}
}

func TestResolvePath_Relative(t *testing.T) {
	p, err := resolvePath("data/file.xlsx")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if p == "" {
		t.Error("expected non-empty resolved path")
	}
	if !strings.HasSuffix(p, "data/file.xlsx") {
		t.Errorf("expected path ending in data/file.xlsx, got %s", p)
	}
}

func TestMin(t *testing.T) {
	if min(3, 5) != 3 {
		t.Error("min(3,5) should be 3")
	}
	if min(10, 2) != 2 {
		t.Error("min(10,2) should be 2")
	}
	if min(4, 4) != 4 {
		t.Error("min(4,4) should be 4")
	}
}
