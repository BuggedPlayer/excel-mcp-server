package tools

import (
	"context"
	"fmt"
	"os/exec"
	"runtime"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ExcelOpenWorkbookArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
}

var excelOpenWorkbookArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
})

func AddExcelOpenWorkbookTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_open_workbook",
		mcp.WithDescription("Open a workbook file in Excel application (Windows only)"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file to open"),
		),
	), WithRecovery(handleOpenWorkbook))
}

func handleOpenWorkbook(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelOpenWorkbookArguments{}
	if issues := excelOpenWorkbookArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return openWorkbook(args.FileAbsolutePath)
}

func openWorkbook(fileAbsolutePath string) (*mcp.CallToolResult, error) {
	// Try OLE first for richer integration
	app, err := excel.NewExcelApp()
	if err == nil {
		defer app.Release()
		if err := app.OpenWorkbook(fileAbsolutePath); err == nil {
			result := "# Notice\n"
			result += fmt.Sprintf("Workbook opened in Excel via OLE: %s\n", fileAbsolutePath)
			return mcp.NewToolResultText(result), nil
		}
	}

	// Fallback: use OS shell to open the file with the default application
	if runtime.GOOS == "windows" {
		cmd := exec.Command("cmd", "/c", "start", "", fileAbsolutePath)
		if err := cmd.Start(); err != nil {
			return nil, fmt.Errorf("failed to open workbook: %w", err)
		}
		result := "# Notice\n"
		result += fmt.Sprintf("Workbook opened in Excel: %s\n", fileAbsolutePath)
		return mcp.NewToolResultText(result), nil
	}

	return nil, fmt.Errorf("opening workbooks in Excel is only supported on Windows")
}
