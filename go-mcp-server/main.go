package main

import (
	"context"
	"log"
	"os"
	"os/signal"
	"syscall"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/mmanthe37/excel-engine/go-mcp-server/config"
	"github.com/mmanthe37/excel-engine/go-mcp-server/tools"
)

func main() {
	cfg := config.Load()
	tools.SetConfig(cfg)

	server := mcp.NewServer(
		&mcp.Implementation{Name: cfg.ServerName, Version: cfg.Version},
		nil,
	)
	tools.RegisterTools(server)

	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()

	// Graceful shutdown on SIGINT/SIGTERM
	sigCh := make(chan os.Signal, 1)
	signal.Notify(sigCh, syscall.SIGINT, syscall.SIGTERM)
	go func() {
		<-sigCh
		log.Println("Shutting down...")
		cancel()
	}()

	log.Printf("Starting %s %s via stdio transport", cfg.ServerName, cfg.Version)

	if err := server.Run(ctx, mcp.NewStdioTransport()); err != nil {
		log.Fatalf("Server error: %v", err)
	}
}
