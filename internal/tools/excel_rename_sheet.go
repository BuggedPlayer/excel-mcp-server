package tools

import (
	"context"
	"fmt"
	"html"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ExcelRenameSheetArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	OldSheetName     string `zog:"oldSheetName"`
	NewSheetName     string `zog:"newSheetName"`
}

var excelRenameSheetArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"oldSheetName":     z.String().Required(),
	"newSheetName":     z.String().Required(),
})

func AddExcelRenameSheetTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_rename_sheet",
		mcp.WithDescription("Rename a sheet in the Excel file"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("oldSheetName",
			mcp.Required(),
			mcp.Description("Current name of the sheet"),
		),
		mcp.WithString("newSheetName",
			mcp.Required(),
			mcp.Description("New name for the sheet"),
		),
	), WithRecovery(handleRenameSheet))
}

func handleRenameSheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelRenameSheetArguments{}
	if issues := excelRenameSheetArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return renameSheet(args.FileAbsolutePath, args.OldSheetName, args.NewSheetName)
}

func renameSheet(fileAbsolutePath string, oldSheetName string, newSheetName string) (*mcp.CallToolResult, error) {
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer release()

	if err := workbook.RenameSheet(oldSheetName, newSheetName); err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Sheet [%s] renamed to [%s].\n", html.EscapeString(oldSheetName), html.EscapeString(newSheetName))
	return mcp.NewToolResultText(result), nil
}
