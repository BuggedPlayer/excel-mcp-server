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

type ExcelAddChartArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	ChartType        string `zog:"chartType"`
	DataRange        string `zog:"dataRange"`
	Title            string `zog:"title"`
	Position         string `zog:"position"`
}

var excelAddChartArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"chartType":        z.String().Required(),
	"dataRange":        z.String().Required(),
	"title":            z.String(),
	"position":         z.String().Required(),
})

func AddExcelAddChartTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_add_chart",
		mcp.WithDescription("Create a chart in the Excel sheet"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name where the chart is placed"),
		),
		mcp.WithString("chartType",
			mcp.Required(),
			mcp.Description("Chart type: \"col\", \"bar\", \"line\", \"pie\", \"area\", or \"scatter\""),
		),
		mcp.WithString("dataRange",
			mcp.Required(),
			mcp.Description("Data range for the chart (e.g., \"A1:D5\")"),
		),
		mcp.WithString("title",
			mcp.Description("Chart title"),
		),
		mcp.WithString("position",
			mcp.Required(),
			mcp.Description("Cell where the top-left of the chart is anchored (e.g., \"F1\")"),
		),
	), WithRecovery(handleAddChart))
}

func handleAddChart(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelAddChartArguments{}
	if issues := excelAddChartArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return addChart(args.FileAbsolutePath, args.SheetName, args.ChartType, args.DataRange, args.Title, args.Position)
}

func addChart(fileAbsolutePath string, sheetName string, chartType string, dataRange string, title string, position string) (*mcp.CallToolResult, error) {
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

	if err := worksheet.AddChart(position, chartType, dataRange, title); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Chart (%s) created at %s in sheet [%s].\n", chartType, position, html.EscapeString(sheetName))
	return mcp.NewToolResultText(result), nil
}
