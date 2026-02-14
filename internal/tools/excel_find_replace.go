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

type ExcelFindReplaceArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Find             string `zog:"find"`
	Replace          string `zog:"replace"`
	Range            string `zog:"range"`
	MatchCase        bool   `zog:"matchCase"`
	MatchEntireCell  bool   `zog:"matchEntireCell"`
}

var excelFindReplaceArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"find":             z.String().Required(),
	"replace":          z.String().Required(),
	"range":            z.String(),
	"matchCase":        z.Bool().Default(false),
	"matchEntireCell":  z.Bool().Default(false),
})

func AddExcelFindReplaceTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_find_replace",
		mcp.WithDescription("Find and replace values in the Excel sheet"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("find",
			mcp.Required(),
			mcp.Description("Text to find"),
		),
		mcp.WithString("replace",
			mcp.Required(),
			mcp.Description("Replacement text"),
		),
		mcp.WithString("range",
			mcp.Description("Range to search within (default: entire sheet)"),
		),
		mcp.WithBoolean("matchCase",
			mcp.Description("Case-sensitive search (default: false)"),
		),
		mcp.WithBoolean("matchEntireCell",
			mcp.Description("Match entire cell contents (default: false)"),
		),
	), WithRecovery(handleFindReplace))
}

func handleFindReplace(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelFindReplaceArguments{}
	if issues := excelFindReplaceArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return findReplace(args.FileAbsolutePath, args.SheetName, args.Find, args.Replace, args.Range, args.MatchCase, args.MatchEntireCell)
}

func findReplace(fileAbsolutePath string, sheetName string, find string, replace string, searchRange string, matchCase bool, matchEntireCell bool) (*mcp.CallToolResult, error) {
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

	count, err := worksheet.FindReplace(searchRange, find, replace, matchCase, matchEntireCell)
	if err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	if count >= 0 {
		result += fmt.Sprintf("Replaced %d occurrence(s) of \"%s\" with \"%s\" in sheet [%s].\n", count, html.EscapeString(find), html.EscapeString(replace), html.EscapeString(sheetName))
	} else {
		result += fmt.Sprintf("Replaced \"%s\" with \"%s\" in sheet [%s].\n", html.EscapeString(find), html.EscapeString(replace), html.EscapeString(sheetName))
	}
	return mcp.NewToolResultText(result), nil
}
