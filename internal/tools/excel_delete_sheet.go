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

type ExcelDeleteSheetArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
}

var excelDeleteSheetArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
})

func AddExcelDeleteSheetTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_delete_sheet",
		mcp.WithDescription("Delete a sheet from the Excel file"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Name of the sheet to delete"),
		),
	), WithRecovery(handleDeleteSheet))
}

func handleDeleteSheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelDeleteSheetArguments{}
	if issues := excelDeleteSheetArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return deleteSheet(args.FileAbsolutePath, args.SheetName)
}

func deleteSheet(fileAbsolutePath string, sheetName string) (*mcp.CallToolResult, error) {
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer release()

	if err := workbook.DeleteSheet(sheetName); err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Sheet [%s] deleted.\n", html.EscapeString(sheetName))
	return mcp.NewToolResultText(result), nil
}
