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

type ExcelSetRowHeightArguments struct {
	FileAbsolutePath string  `zog:"fileAbsolutePath"`
	SheetName        string  `zog:"sheetName"`
	Row              int     `zog:"row"`
	EndRow           int     `zog:"endRow"`
	Height           float64 `zog:"height"`
}

var excelSetRowHeightArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"row":              z.Int().GTE(1).Required(),
	"endRow":           z.Int().GTE(1),
	"height":           z.Float64().GTE(0).LTE(409).Required(),
})

func AddExcelSetRowHeightTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_set_row_height",
		mcp.WithDescription("Set the height of one or more rows"),
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
			mcp.Description("Starting row number (1-based)"),
		),
		mcp.WithNumber("endRow",
			mcp.Description("Ending row number (1-based, defaults to row for single row)"),
		),
		mcp.WithNumber("height",
			mcp.Required(),
			mcp.Description("Row height in points (0-409)"),
		),
	), WithRecovery(handleSetRowHeight))
}

func handleSetRowHeight(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelSetRowHeightArguments{}
	if issues := excelSetRowHeightArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return setRowHeight(args.FileAbsolutePath, args.SheetName, args.Row, args.EndRow, args.Height)
}

func setRowHeight(fileAbsolutePath string, sheetName string, row int, endRow int, height float64) (*mcp.CallToolResult, error) {
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

	if endRow == 0 || endRow < row {
		endRow = row
	}
	for r := row; r <= endRow; r++ {
		if err := worksheet.SetRowHeight(r, height); err != nil {
			return nil, err
		}
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	if row == endRow {
		result += fmt.Sprintf("Row height set to %.1f for row %d in sheet [%s].\n", height, row, html.EscapeString(sheetName))
	} else {
		result += fmt.Sprintf("Row height set to %.1f for rows %d-%d in sheet [%s].\n", height, row, endRow, html.EscapeString(sheetName))
	}
	return mcp.NewToolResultText(result), nil
}
