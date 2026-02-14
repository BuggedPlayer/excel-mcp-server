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

type ExcelAddDataValidationArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Range            string `zog:"range"`
	Type             string `zog:"type"`
	Formula1         string `zog:"formula1"`
	Formula2         string `zog:"formula2"`
	AllowBlank       bool   `zog:"allowBlank"`
}

var excelAddDataValidationArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"range":            z.String().Required(),
	"type":             z.String().Required(),
	"formula1":         z.String().Required(),
	"formula2":         z.String(),
	"allowBlank":       z.Bool().Default(true),
})

func AddExcelAddDataValidationTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_add_data_validation",
		mcp.WithDescription("Add data validation rules to a range of cells"),
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
			mcp.Description("Range to apply validation (e.g., \"A1:A100\")"),
		),
		mcp.WithString("type",
			mcp.Required(),
			mcp.Description("Validation type: \"list\", \"whole\", or \"decimal\""),
		),
		mcp.WithString("formula1",
			mcp.Required(),
			mcp.Description("For list: comma-separated values. For whole/decimal: minimum value"),
		),
		mcp.WithString("formula2",
			mcp.Description("For whole/decimal: maximum value"),
		),
		mcp.WithBoolean("allowBlank",
			mcp.Description("Allow blank cells (default: true)"),
		),
	), WithRecovery(handleAddDataValidation))
}

func handleAddDataValidation(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelAddDataValidationArguments{}
	if issues := excelAddDataValidationArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return addDataValidation(args.FileAbsolutePath, args.SheetName, args.Range, args.Type, args.Formula1, args.Formula2, args.AllowBlank)
}

func addDataValidation(fileAbsolutePath string, sheetName string, validationRange string, validationType string, formula1 string, formula2 string, allowBlank bool) (*mcp.CallToolResult, error) {
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

	if err := worksheet.AddDataValidation(validationRange, validationType, formula1, formula2, allowBlank); err != nil {
		return nil, err
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Data validation (%s) added to range [%s] in sheet [%s].\n", validationType, validationRange, html.EscapeString(sheetName))
	return mcp.NewToolResultText(result), nil
}
