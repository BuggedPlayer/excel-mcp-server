package tools

import (
	"context"
	"fmt"
	"html"
	"strings"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ExcelSetColumnWidthArguments struct {
	FileAbsolutePath string  `zog:"fileAbsolutePath"`
	SheetName        string  `zog:"sheetName"`
	Column           string  `zog:"column"`
	Width            float64 `zog:"width"`
}

var excelSetColumnWidthArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"column":           z.String().Required(),
	"width":            z.Float64().GTE(0).LTE(255).Required(),
})

func AddExcelSetColumnWidthTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_set_column_width",
		mcp.WithDescription("Set the width of one or more columns"),
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
			mcp.Description("Column letter or range (e.g., \"A\" or \"A:C\")"),
		),
		mcp.WithNumber("width",
			mcp.Required(),
			mcp.Description("Column width in character units (0-255)"),
		),
	), WithRecovery(handleSetColumnWidth))
}

func handleSetColumnWidth(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelSetColumnWidthArguments{}
	if issues := excelSetColumnWidthArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return setColumnWidth(args.FileAbsolutePath, args.SheetName, args.Column, args.Width)
}

func setColumnWidth(fileAbsolutePath string, sheetName string, column string, width float64) (*mcp.CallToolResult, error) {
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

	startCol := column
	endCol := column
	if parts := strings.SplitN(column, ":", 2); len(parts) == 2 {
		startCol = parts[0]
		endCol = parts[1]
	}

	if err := worksheet.SetColumnWidth(startCol, endCol, width); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Column width set to %.1f for [%s] in sheet [%s].\n", width, column, html.EscapeString(sheetName))
	return mcp.NewToolResultText(result), nil
}
