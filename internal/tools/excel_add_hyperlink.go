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

type ExcelAddHyperlinkArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Cell             string `zog:"cell"`
	Url              string `zog:"url"`
	Display          string `zog:"display"`
}

var excelAddHyperlinkArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"cell":             z.String().Required(),
	"url":              z.String().Required(),
	"display":          z.String(),
})

func AddExcelAddHyperlinkTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_add_hyperlink",
		mcp.WithDescription("Add a hyperlink to a cell in the Excel sheet"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("cell",
			mcp.Required(),
			mcp.Description("Cell reference (e.g., \"A1\")"),
		),
		mcp.WithString("url",
			mcp.Required(),
			mcp.Description("URL or internal reference (e.g., \"https://example.com\" or \"Sheet2!A1\")"),
		),
		mcp.WithString("display",
			mcp.Description("Display text for the hyperlink"),
		),
	), WithRecovery(handleAddHyperlink))
}

func handleAddHyperlink(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelAddHyperlinkArguments{}
	if issues := excelAddHyperlinkArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return addHyperlink(args.FileAbsolutePath, args.SheetName, args.Cell, args.Url, args.Display)
}

func addHyperlink(fileAbsolutePath string, sheetName string, cell string, url string, display string) (*mcp.CallToolResult, error) {
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

	if err := worksheet.AddHyperlink(cell, url, display); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Hyperlink added to cell %s in sheet [%s] -> %s\n", cell, html.EscapeString(sheetName), html.EscapeString(url))
	return mcp.NewToolResultText(result), nil
}
