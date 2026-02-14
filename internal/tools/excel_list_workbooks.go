package tools

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
)

func AddExcelListWorkbooksTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_list_workbooks",
		mcp.WithDescription("List all workbooks currently open in Excel (Windows only)"),
	), WithRecovery(handleListWorkbooks))
}

func handleListWorkbooks(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	return listWorkbooks()
}

func listWorkbooks() (*mcp.CallToolResult, error) {
	app, err := excel.NewExcelApp()
	if err != nil {
		return nil, err
	}
	defer app.Release()

	workbooks, err := app.ListWorkbooks()
	if err != nil {
		return nil, err
	}

	jsonData, err := json.MarshalIndent(workbooks, "", "  ")
	if err != nil {
		return nil, fmt.Errorf("failed to marshal workbook list: %w", err)
	}

	result := "# Open Workbooks\n"
	result += fmt.Sprintf("Found %d open workbook(s):\n\n", len(workbooks))
	result += "```json\n"
	result += string(jsonData) + "\n"
	result += "```\n"
	return mcp.NewToolResultText(result), nil
}
