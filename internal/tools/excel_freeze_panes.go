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

type ExcelFreezePanesArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Cell             string `zog:"cell"`
}

var excelFreezePanesArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"cell":             z.String().Required(),
})

func AddExcelFreezePanesTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_freeze_panes",
		mcp.WithDescription("Freeze rows and columns at the specified cell"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("cell",
			mcp.Required(),
			mcp.Description("Cell below and to the right of the freeze (e.g., \"B2\" freezes row 1 and column A)"),
		),
	), WithRecovery(handleFreezePanes))
}

func handleFreezePanes(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelFreezePanesArguments{}
	if issues := excelFreezePanesArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return freezePanes(args.FileAbsolutePath, args.SheetName, args.Cell)
}

func freezePanes(fileAbsolutePath string, sheetName string, cell string) (*mcp.CallToolResult, error) {
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer release()

	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	defer worksheet.Release()

	if err := worksheet.FreezePanes(cell); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Panes frozen at %s in sheet [%s].\n", cell, html.EscapeString(sheetName))
	return mcp.NewToolResultText(result), nil
}
