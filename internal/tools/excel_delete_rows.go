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

type ExcelDeleteRowsArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Row              int    `zog:"row"`
	Count            int    `zog:"count"`
}

var excelDeleteRowsArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"row":              z.Int().GTE(1).Required(),
	"count":            z.Int().GTE(1).Required(),
})

func AddExcelDeleteRowsTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_delete_rows",
		mcp.WithDescription("Delete rows at the specified position"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithNumber("row",
			mcp.Required(),
			mcp.Description("Starting row number to delete (1-based)"),
		),
		mcp.WithNumber("count",
			mcp.Required(),
			mcp.Description("Number of rows to delete"),
		),
	), WithRecovery(handleDeleteRows))
}

func handleDeleteRows(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelDeleteRowsArguments{}
	if issues := excelDeleteRowsArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return deleteRows(args.FileAbsolutePath, args.SheetName, args.Row, args.Count)
}

func deleteRows(fileAbsolutePath string, sheetName string, row int, count int) (*mcp.CallToolResult, error) {
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

	if err := worksheet.DeleteRows(row, count); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Deleted %d row(s) starting at row %d in sheet [%s].\n", count, row, html.EscapeString(sheetName))
	return mcp.NewToolResultText(result), nil
}
