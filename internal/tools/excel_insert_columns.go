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

type ExcelInsertColumnsArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Column           string `zog:"column"`
	Count            int    `zog:"count"`
}

var excelInsertColumnsArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"column":           z.String().Required(),
	"count":            z.Int().GTE(1).Required(),
})

func AddExcelInsertColumnsTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_insert_columns",
		mcp.WithDescription("Insert empty columns at the specified position"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("column",
			mcp.Required(),
			mcp.Description("Column letter before which to insert (e.g., \"B\")"),
		),
		mcp.WithNumber("count",
			mcp.Required(),
			mcp.Description("Number of columns to insert"),
		),
	), WithRecovery(handleInsertColumns))
}

func handleInsertColumns(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelInsertColumnsArguments{}
	if issues := excelInsertColumnsArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return insertColumns(args.FileAbsolutePath, args.SheetName, args.Column, args.Count)
}

func insertColumns(fileAbsolutePath string, sheetName string, column string, count int) (*mcp.CallToolResult, error) {
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

	if err := worksheet.InsertColumns(column, count); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Inserted %d column(s) at column %s in sheet [%s].\n", count, column, html.EscapeString(sheetName))
	return mcp.NewToolResultText(result), nil
}
