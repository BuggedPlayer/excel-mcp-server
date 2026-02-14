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

type ExcelSetNamedRangeArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	Name             string `zog:"name"`
	RefersTo         string `zog:"refersTo"`
	Scope            string `zog:"scope"`
}

var excelSetNamedRangeArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"name":             z.String().Required(),
	"refersTo":         z.String().Required(),
	"scope":            z.String(),
})

func AddExcelSetNamedRangeTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_set_named_range",
		mcp.WithDescription("Create or update a named range in the workbook"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("name",
			mcp.Required(),
			mcp.Description("Name for the range (e.g., \"SalesData\")"),
		),
		mcp.WithString("refersTo",
			mcp.Required(),
			mcp.Description("Range reference (e.g., \"Sheet1!A1:C10\")"),
		),
		mcp.WithString("scope",
			mcp.Description("Scope: \"workbook\" (default) or sheet name for sheet-scoped"),
		),
	), WithRecovery(handleSetNamedRange))
}

func handleSetNamedRange(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelSetNamedRangeArguments{}
	if issues := excelSetNamedRangeArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return setNamedRange(args.FileAbsolutePath, args.Name, args.RefersTo, args.Scope)
}

func setNamedRange(fileAbsolutePath string, name string, refersTo string, scope string) (*mcp.CallToolResult, error) {
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer release()

	if err := workbook.SetDefinedName(name, refersTo, scope); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Named range \"%s\" set to %s.\n", html.EscapeString(name), html.EscapeString(refersTo))
	return mcp.NewToolResultText(result), nil
}
