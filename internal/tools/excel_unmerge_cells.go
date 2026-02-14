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

type ExcelUnmergeCellsArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Range            string `zog:"range"`
}

var excelUnmergeCellsArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"range":            z.String().Required(),
})

func AddExcelUnmergeCellsTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_unmerge_cells",
		mcp.WithDescription("Unmerge previously merged cells in the specified range"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("range",
			mcp.Required(),
			mcp.Description("Range of cells to unmerge (e.g., \"A1:C3\")"),
		),
	), WithRecovery(handleUnmergeCells))
}

func handleUnmergeCells(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelUnmergeCellsArguments{}
	if issues := excelUnmergeCellsArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return unmergeCells(args.FileAbsolutePath, args.SheetName, args.Range)
}

func unmergeCells(fileAbsolutePath string, sheetName string, mergeRange string) (*mcp.CallToolResult, error) {
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

	if err := worksheet.UnmergeCells(mergeRange); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Cells [%s] unmerged in sheet [%s].\n", mergeRange, html.EscapeString(sheetName))
	return mcp.NewToolResultText(result), nil
}
