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

type ExcelCreateWorkbookArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
}

var excelCreateWorkbookArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
})

func AddExcelCreateWorkbookTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_create_workbook",
		mcp.WithDescription("Create a new workbook in Excel application and save it (Windows only)"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path where the new Excel file will be saved"),
		),
	), WithRecovery(handleCreateWorkbook))
}

func handleCreateWorkbook(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelCreateWorkbookArguments{}
	if issues := excelCreateWorkbookArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return createWorkbook(args.FileAbsolutePath)
}

func createWorkbook(fileAbsolutePath string) (*mcp.CallToolResult, error) {
	// Try OLE first for richer integration
	app, oleErr := excel.NewExcelApp()
	if oleErr == nil {
		defer app.Release()
		if err := app.CreateWorkbook(fileAbsolutePath); err == nil {
			result := "# Notice\n"
			result += fmt.Sprintf("New workbook created and opened in Excel: %s\n", fileAbsolutePath)
			return mcp.NewToolResultText(result), nil
		}
	}

	// Fallback: create the file with excelize, then open it with the OS shell
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, fmt.Errorf("failed to create workbook: %w", err)
	}
	if err := workbook.Save(); err != nil {
		release()
		return nil, fmt.Errorf("failed to save workbook: %w", err)
	}
	release()

	// Open the file in Excel via OS shell
	if runtime.GOOS == "windows" {
		cmd := exec.Command("cmd", "/c", "start", "", fileAbsolutePath)
		_ = cmd.Start()
	}

	result := "# Notice\n"
	result += fmt.Sprintf("New workbook created and saved to: %s\n", fileAbsolutePath)
	return mcp.NewToolResultText(result), nil
}
