package tools

import (
	"context"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

// RegisterTools registers all Excel Engine tools with the MCP server.
func RegisterTools(server *mcp.Server) {
	mcp.AddTool(server,
		&mcp.Tool{
			Name:        "complete_assignment",
			Description: "Complete an Excel assignment autonomously. Parses the instruction file, plans execution across 6 automation layers, executes all tasks, and verifies results.",
		},
		func(ctx context.Context, ss *mcp.ServerSession, params *mcp.CallToolParamsFor[CompleteAssignmentInput]) (*mcp.CallToolResultFor[any], error) {
			result, err := CompleteAssignmentHandler(ctx, params.Arguments)
			if err != nil {
				return nil, err
			}
			return &mcp.CallToolResultFor[any]{
				Content: []mcp.Content{&mcp.TextContent{Text: result}},
			}, nil
		},
	)

	mcp.AddTool(server,
		&mcp.Tool{
			Name:        "parse_instructions",
			Description: "Parse an instruction file into structured Excel tasks. Reads .docx, .pdf, or .txt files and extracts individual operations (formulas, formatting, charts, etc.).",
		},
		func(ctx context.Context, ss *mcp.ServerSession, params *mcp.CallToolParamsFor[ParseInstructionsInput]) (*mcp.CallToolResultFor[any], error) {
			result, err := ParseInstructionsHandler(ctx, params.Arguments)
			if err != nil {
				return nil, err
			}
			return &mcp.CallToolResultFor[any]{
				Content: []mcp.Content{&mcp.TextContent{Text: result}},
			}, nil
		},
	)

	mcp.AddTool(server,
		&mcp.Tool{
			Name:        "execute_openpyxl",
			Description: "Run Phase 1 offline operations via the openpyxl layer. Executes tasks against the workbook entirely offline — Excel does not need to be open.",
		},
		func(ctx context.Context, ss *mcp.ServerSession, params *mcp.CallToolParamsFor[ExecuteOpenpyxlInput]) (*mcp.CallToolResultFor[any], error) {
			result, err := ExecuteOpenpyxlHandler(ctx, params.Arguments)
			if err != nil {
				return nil, err
			}
			return &mcp.CallToolResultFor[any]{
				Content: []mcp.Content{&mcp.TextContent{Text: result}},
			}, nil
		},
	)

	mcp.AddTool(server,
		&mcp.Tool{
			Name:        "execute_live",
			Description: "Run Phase 2 live Excel operations using xlwings, AppleScript, System Events, VBA, or PyAutoGUI (Layers 2-6). Requires Excel to be open on macOS.",
		},
		func(ctx context.Context, ss *mcp.ServerSession, params *mcp.CallToolParamsFor[ExecuteLiveInput]) (*mcp.CallToolResultFor[any], error) {
			result, err := ExecuteLiveHandler(ctx, params.Arguments)
			if err != nil {
				return nil, err
			}
			return &mcp.CallToolResultFor[any]{
				Content: []mcp.Content{&mcp.TextContent{Text: result}},
			}, nil
		},
	)

	mcp.AddTool(server,
		&mcp.Tool{
			Name:        "verify_workbook",
			Description: "Verify assignment completion against expected tasks. Opens the workbook and checks whether each task was completed. General structural audit if no tasks given.",
		},
		func(ctx context.Context, ss *mcp.ServerSession, params *mcp.CallToolParamsFor[VerifyWorkbookInput]) (*mcp.CallToolResultFor[any], error) {
			result, err := VerifyWorkbookHandler(ctx, params.Arguments)
			if err != nil {
				return nil, err
			}
			return &mcp.CallToolResultFor[any]{
				Content: []mcp.Content{&mcp.TextContent{Text: result}},
			}, nil
		},
	)

	mcp.AddTool(server,
		&mcp.Tool{
			Name:        "get_engine_status",
			Description: "Get current engine configuration, version, available layers, and supported task types.",
		},
		func(ctx context.Context, ss *mcp.ServerSession, params *mcp.CallToolParamsFor[GetEngineStatusInput]) (*mcp.CallToolResultFor[any], error) {
			result, err := GetEngineStatusHandler(ctx, params.Arguments)
			if err != nil {
				return nil, err
			}
			return &mcp.CallToolResultFor[any]{
				Content: []mcp.Content{&mcp.TextContent{Text: result}},
			}, nil
		},
	)
}
