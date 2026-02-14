package tools

import (
	"context"
	"encoding/csv"
	"fmt"
	"html"
	"os"
	"path/filepath"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
	"github.com/xuri/excelize/v2"
)

type ExcelExportCsvArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	OutputPath       string `zog:"outputPath"`
	Range            string `zog:"range"`
	Delimiter        string `zog:"delimiter"`
}

var excelExportCsvArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"outputPath":       z.String().Test(AbsolutePathTest()).Required(),
	"range":            z.String(),
	"delimiter":        z.String().Default(","),
})

func AddExcelExportCsvTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_export_csv",
		mcp.WithDescription("Export an Excel sheet to a CSV file"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("outputPath",
			mcp.Required(),
			mcp.Description("Absolute path for the output CSV file"),
		),
		mcp.WithString("range",
			mcp.Description("Range to export (e.g., \"A1:D10\"). If omitted, exports the entire used range"),
		),
		mcp.WithString("delimiter",
			mcp.Description("CSV delimiter character (default: \",\")"),
		),
	), WithRecovery(handleExportCsv))
}

func handleExportCsv(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelExportCsvArguments{}
	if issues := excelExportCsvArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return exportCsv(args.FileAbsolutePath, args.SheetName, args.OutputPath, args.Range, args.Delimiter)
}

func exportCsv(fileAbsolutePath string, sheetName string, outputPath string, rangeStr string, delimiter string) (*mcp.CallToolResult, error) {
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

	// Determine the range to export
	if rangeStr == "" {
		dim, err := worksheet.GetDimension()
		if err != nil {
			return nil, err
		}
		rangeStr = dim
	}

	startCol, startRow, endCol, endRow, err := excel.ParseRange(rangeStr)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}

	// Determine delimiter rune
	delimRune := ','
	if len(delimiter) > 0 {
		delimRune = rune(delimiter[0])
	}

	// Create output directory if needed
	if err := os.MkdirAll(filepath.Dir(outputPath), os.ModePerm); err != nil {
		return nil, fmt.Errorf("failed to create output directory: %w", err)
	}

	// Create output file
	file, err := os.Create(outputPath)
	if err != nil {
		return nil, fmt.Errorf("failed to create output file: %w", err)
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	writer.Comma = delimRune

	rowCount := 0
	for row := startRow; row <= endRow; row++ {
		record := make([]string, 0, endCol-startCol+1)
		for col := startCol; col <= endCol; col++ {
			cellName, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return nil, err
			}
			val, err := worksheet.GetValue(cellName)
			if err != nil {
				return nil, err
			}
			record = append(record, val)
		}
		if err := writer.Write(record); err != nil {
			return nil, fmt.Errorf("failed to write CSV row: %w", err)
		}
		rowCount++
	}
	writer.Flush()
	if err := writer.Error(); err != nil {
		return nil, fmt.Errorf("CSV write error: %w", err)
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Exported %d rows from sheet [%s] to %s\n", rowCount, html.EscapeString(sheetName), outputPath)
	return mcp.NewToolResultText(result), nil
}
