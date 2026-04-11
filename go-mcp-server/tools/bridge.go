package tools

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"os/exec"
	"path/filepath"
	"strings"
	"time"

	"github.com/mmanthe37/excel-engine/go-mcp-server/config"
)

// pythonBridge calls the Python excel-engine via subprocess and returns the
// JSON result. This is the core integration — Go handles the MCP protocol
// while Python handles the actual Excel automation.
func pythonBridge(ctx context.Context, cfg *config.Config, script string) (json.RawMessage, error) {
	pythonPath := cfg.PythonPath
	engineRoot := cfg.EngineRoot

	// Build the Python command with the engine root on sys.path
	args := []string{"-c", fmt.Sprintf(
		"import sys; sys.path.insert(0, %q); %s",
		engineRoot, script,
	)}

	ctx, cancel := context.WithTimeout(ctx, 120*time.Second)
	defer cancel()

	cmd := exec.CommandContext(ctx, pythonPath, args...)

	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr

	if err := cmd.Run(); err != nil {
		stderrStr := strings.TrimSpace(stderr.String())
		if stderrStr != "" {
			return nil, fmt.Errorf("python error: %s", stderrStr)
		}
		return nil, fmt.Errorf("python execution failed: %w", err)
	}

	raw := bytes.TrimSpace(stdout.Bytes())
	if len(raw) == 0 {
		return json.RawMessage(`{"error":"empty response from python"}`), nil
	}

	// Validate it's valid JSON
	if !json.Valid(raw) {
		return nil, fmt.Errorf("invalid JSON from python: %s", string(raw[:min(len(raw), 200)]))
	}

	return json.RawMessage(raw), nil
}

// resolvePath validates and resolves a file path, ensuring it's under the home directory.
func resolvePath(p string) (string, error) {
	if p == "" {
		return "", fmt.Errorf("path cannot be empty")
	}

	// Expand ~ to home directory
	if strings.HasPrefix(p, "~/") {
		home, err := filepath.Abs(filepath.Join("~"))
		if err != nil {
			return "", fmt.Errorf("cannot resolve home: %w", err)
		}
		p = filepath.Join(home, p[2:])
	}

	abs, err := filepath.Abs(p)
	if err != nil {
		return "", fmt.Errorf("cannot resolve path: %w", err)
	}

	return abs, nil
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
