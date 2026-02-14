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

type ExcelSetConditionalFormatArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Range            string `zog:"range"`
	Type             string `zog:"type"`
	Criteria         string `zog:"criteria"`
	Value            string `zog:"value"`
	Value2           string `zog:"value2"`
	FontColor        string `zog:"fontColor"`
	BgColor          string `zog:"bgColor"`
}

var excelSetConditionalFormatArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"range":            z.String().Required(),
	"type":             z.String().Required(),
	"criteria":         z.String(),
	"value":            z.String(),
	"value2":           z.String(),
	"fontColor":        z.String(),
	"bgColor":          z.String(),
})

func AddExcelSetConditionalFormatTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_set_conditional_format",
		mcp.WithDescription("Apply conditional formatting rules to a range of cells"),
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
			mcp.Description("Range to apply formatting (e.g., \"A1:A100\")"),
		),
		mcp.WithString("type",
			mcp.Required(),
			mcp.Description("Rule type: \"cell\", \"top\", \"duplicate\", \"colorScale\", or \"dataBar\""),
		),
		mcp.WithString("criteria",
			mcp.Description("For \"cell\" type: \">\", \"<\", \">=\", \"<=\", \"==\", \"!=\", or \"between\""),
		),
		mcp.WithString("value",
			mcp.Description("Comparison value (or min value for \"between\")"),
		),
		mcp.WithString("value2",
			mcp.Description("Max value for \"between\" criteria"),
		),
		mcp.WithString("fontColor",
			mcp.Description("Font color in hex format (e.g., \"#FF0000\")"),
		),
		mcp.WithString("bgColor",
			mcp.Description("Background color in hex format (e.g., \"#FFFF00\")"),
		),
	), WithRecovery(handleSetConditionalFormat))
}

func handleSetConditionalFormat(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelSetConditionalFormatArguments{}
	if issues := excelSetConditionalFormatArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return setConditionalFormat(args.FileAbsolutePath, args.SheetName, args.Range, args.Type, args.Criteria, args.Value, args.Value2, args.FontColor, args.BgColor)
}

func setConditionalFormat(fileAbsolutePath string, sheetName string, formatRange string, ruleType string, criteria string, value string, value2 string, fontColor string, bgColor string) (*mcp.CallToolResult, error) {
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

	if err := worksheet.SetConditionalFormat(formatRange, ruleType, criteria, value, value2, fontColor, bgColor); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Conditional formatting (%s) applied to %s in sheet [%s].\n", ruleType, formatRange, html.EscapeString(sheetName))
	return mcp.NewToolResultText(result), nil
}
