package main

import (
	"context"
	"os"
	"testing"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/mmanthe37/excel-engine/go-mcp-server/config"
	"github.com/mmanthe37/excel-engine/go-mcp-server/tools"
)

func TestConfigLoadDefaults(t *testing.T) {
	os.Unsetenv("SERVER_NAME")
	os.Unsetenv("VERSION")
	os.Unsetenv("PYTHON_PATH")
	os.Unsetenv("ENGINE_ROOT")

	cfg := config.Load()

	if cfg.ServerName != "excel-engine-go" {
		t.Errorf("expected ServerName=excel-engine-go, got %s", cfg.ServerName)
	}
	if cfg.Version != "v1.1.0" {
		t.Errorf("expected Version=v1.1.0, got %s", cfg.Version)
	}
	if cfg.PythonPath != "python3" {
		t.Errorf("expected PythonPath=python3, got %s", cfg.PythonPath)
	}
	if cfg.LogLevel != "info" {
		t.Errorf("expected LogLevel=info, got %s", cfg.LogLevel)
	}
	if cfg.MaxTasks != 500 {
		t.Errorf("expected MaxTasks=500, got %d", cfg.MaxTasks)
	}
}

func TestConfigLoadFromEnv(t *testing.T) {
	t.Setenv("SERVER_NAME", "test-server")
	t.Setenv("VERSION", "v9.9.9")
	t.Setenv("PYTHON_PATH", "/usr/local/bin/python3.12")
	t.Setenv("ENGINE_ROOT", "/opt/excel-engine")
	t.Setenv("LOG_LEVEL", "debug")

	cfg := config.Load()

	if cfg.ServerName != "test-server" {
		t.Errorf("expected ServerName=test-server, got %s", cfg.ServerName)
	}
	if cfg.Version != "v9.9.9" {
		t.Errorf("expected Version=v9.9.9, got %s", cfg.Version)
	}
	if cfg.PythonPath != "/usr/local/bin/python3.12" {
		t.Errorf("expected PythonPath from env, got %s", cfg.PythonPath)
	}
	if cfg.EngineRoot != "/opt/excel-engine" {
		t.Errorf("expected EngineRoot from env, got %s", cfg.EngineRoot)
	}
	if cfg.LogLevel != "debug" {
		t.Errorf("expected LogLevel=debug, got %s", cfg.LogLevel)
	}
}

func TestServerInitialization(t *testing.T) {
	cfg := config.Load()
	tools.SetConfig(cfg)

	server := mcp.NewServer(
		&mcp.Implementation{Name: cfg.ServerName, Version: cfg.Version},
		nil,
	)
	if server == nil {
		t.Fatal("expected non-nil server")
	}
}

func TestToolRegistration(t *testing.T) {
	cfg := config.Load()
	tools.SetConfig(cfg)

	server := mcp.NewServer(
		&mcp.Implementation{Name: cfg.ServerName, Version: cfg.Version},
		nil,
	)

	// RegisterTools should not panic
	tools.RegisterTools(server)
}

func TestAllToolsRegistered(t *testing.T) {
	cfg := config.Load()
	tools.SetConfig(cfg)

	server := mcp.NewServer(
		&mcp.Implementation{Name: cfg.ServerName, Version: cfg.Version},
		nil,
	)
	tools.RegisterTools(server)

	// Connect a client via in-memory transports to list tools
	clientTransport, serverTransport := mcp.NewInMemoryTransports()

	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()

	go func() {
		_ = server.Run(ctx, serverTransport)
	}()

	client := mcp.NewClient(
		&mcp.Implementation{Name: "test-client", Version: "0.0.1"},
		nil,
	)
	cs, err := client.Connect(ctx, clientTransport)
	if err != nil {
		t.Fatalf("client connect failed: %v", err)
	}
	defer cs.Close()

	result, err := cs.ListTools(ctx, nil)
	if err != nil {
		t.Fatalf("ListTools failed: %v", err)
	}

	expectedTools := []string{
		"complete_assignment",
		"parse_instructions",
		"execute_openpyxl",
		"execute_live",
		"verify_workbook",
		"get_engine_status",
	}

	registered := make(map[string]bool)
	for _, tool := range result.Tools {
		registered[tool.Name] = true
	}

	for _, name := range expectedTools {
		if !registered[name] {
			t.Errorf("tool %q not registered", name)
		}
	}

	if len(result.Tools) != len(expectedTools) {
		t.Errorf("expected %d tools, got %d", len(expectedTools), len(result.Tools))
	}
}
