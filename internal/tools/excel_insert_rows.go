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

type ExcelInsertRowsArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Row              int    `zog:"row"`
	Count            int    `zog:"count"`
}

var excelInsertRowsArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"row":              z.Int().GTE(1).Required(),
	"count":            z.Int().GTE(1).Required(),
})

func AddExcelInsertRowsTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_insert_rows",
		mcp.WithDescription("Insert empty rows at the specified position"),
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
			mcp.Description("Row number before which to insert (1-based)"),
		),
		mcp.WithNumber("count",
			mcp.Required(),
			mcp.Description("Number of rows to insert"),
		),
	), WithRecovery(handleInsertRows))
}

func handleInsertRows(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelInsertRowsArguments{}
	if issues := excelInsertRowsArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return insertRows(args.FileAbsolutePath, args.SheetName, args.Row, args.Count)
}

func insertRows(fileAbsolutePath string, sheetName string, row int, count int) (*mcp.CallToolResult, error) {
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

	if err := worksheet.InsertRows(row, count); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Inserted %d row(s) at row %d in sheet [%s].\n", count, row, html.EscapeString(sheetName))
	return mcp.NewToolResultText(result), nil
}
