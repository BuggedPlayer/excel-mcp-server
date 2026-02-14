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

type ExcelAddCommentArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Cell             string `zog:"cell"`
	Author           string `zog:"author"`
	Text             string `zog:"text"`
}

var excelAddCommentArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"cell":             z.String().Required(),
	"author":           z.String(),
	"text":             z.String().Required(),
})

func AddExcelAddCommentTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_add_comment",
		mcp.WithDescription("Add a comment to a cell in the Excel sheet"),
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
		mcp.WithString("author",
			mcp.Description("Comment author name"),
		),
		mcp.WithString("text",
			mcp.Required(),
			mcp.Description("Comment text"),
		),
	), WithRecovery(handleAddComment))
}

func handleAddComment(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelAddCommentArguments{}
	if issues := excelAddCommentArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return addComment(args.FileAbsolutePath, args.SheetName, args.Cell, args.Author, args.Text)
}

func addComment(fileAbsolutePath string, sheetName string, cell string, author string, text string) (*mcp.CallToolResult, error) {
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

	if err := worksheet.AddComment(cell, author, text); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Comment added to cell %s in sheet [%s].\n", cell, html.EscapeString(sheetName))
	return mcp.NewToolResultText(result), nil
}
