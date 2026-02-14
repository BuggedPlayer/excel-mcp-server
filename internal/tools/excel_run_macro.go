package tools

import (
	"context"
	"fmt"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ExcelRunMacroArguments struct {
	FileAbsolutePath string   `zog:"fileAbsolutePath"`
	MacroName        string   `zog:"macroName"`
	Args             []string `zog:"args"`
}

var excelRunMacroArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"macroName":        z.String().Required(),
	"args":             z.Slice(z.String()),
})

func AddExcelRunMacroTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_run_macro",
		mcp.WithDescription("Run a VBA macro in an Excel workbook (Windows only)"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file containing the macro"),
		),
		mcp.WithString("macroName",
			mcp.Required(),
			mcp.Description("Name of the macro to run (e.g., \"Sheet1.MyMacro\")"),
		),
		mcp.WithArray("args",
			mcp.Description("Arguments to pass to the macro (up to 10 string arguments)"),
			mcp.Items(map[string]any{
				"type": "string",
			}),
		),
	), WithRecovery(handleRunMacro))
}

func handleRunMacro(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelRunMacroArguments{}
	if issues := excelRunMacroArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return runMacro(args.FileAbsolutePath, args.MacroName, args.Args)
}

func runMacro(fileAbsolutePath string, macroName string, args []string) (*mcp.CallToolResult, error) {
	if len(args) > 10 {
		return imcp.NewToolResultInvalidArgumentError("macro supports at most 10 arguments"), nil
	}

	// First ensure the workbook is open in Excel
	app, err := excel.NewExcelApp()
	if err != nil {
		return nil, err
	}
	defer app.Release()

	// Open the workbook if not already open
	if err := app.OpenWorkbook(fileAbsolutePath); err != nil {
		// Ignore error if workbook is already open
		_ = err
	}

	// Run the macro
	macroResult, err := app.RunMacro(macroName, args)
	if err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("Macro '%s' executed successfully.\n", macroName)
	if macroResult != "" {
		result += fmt.Sprintf("Return value: %s\n", macroResult)
	}
	return mcp.NewToolResultText(result), nil
}
