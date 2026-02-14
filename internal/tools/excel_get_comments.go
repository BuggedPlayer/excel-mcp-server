package tools

import (
	"context"
	"encoding/json"
	"fmt"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ExcelGetCommentsArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
}

var excelGetCommentsArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
})

func AddExcelGetCommentsTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_get_comments",
		mcp.WithDescription("Get all comments from a sheet in the Excel file"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
	), WithRecovery(handleGetComments))
}

func handleGetComments(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelGetCommentsArguments{}
	if issues := excelGetCommentsArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return getComments(args.FileAbsolutePath, args.SheetName)
}

func getComments(fileAbsolutePath string, sheetName string) (*mcp.CallToolResult, error) {
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

	comments, err := worksheet.GetComments()
	if err != nil {
		return nil, err
	}

	jsonData, err := json.MarshalIndent(comments, "", "  ")
	if err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Found %d comment(s) in sheet [%s].\n\n", len(comments), sheetName)
	result += string(jsonData) + "\n"
	return mcp.NewToolResultText(result), nil
}
