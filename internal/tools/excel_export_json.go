package tools

import (
	"context"
	"encoding/json"
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

type ExcelExportJsonArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	OutputPath       string `zog:"outputPath"`
	Range            string `zog:"range"`
	HeaderRow        bool   `zog:"headerRow"`
}

var excelExportJsonArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"outputPath":       z.String().Test(AbsolutePathTest()).Required(),
	"range":            z.String(),
	"headerRow":        z.Bool().Default(true),
})

func AddExcelExportJsonTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_export_json",
		mcp.WithDescription("Export an Excel sheet to a JSON file"),
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
			mcp.Description("Absolute path for the output JSON file"),
		),
		mcp.WithString("range",
			mcp.Description("Range to export (e.g., \"A1:D10\"). If omitted, exports the entire used range"),
		),
		mcp.WithBoolean("headerRow",
			mcp.Description("Use the first row as JSON keys (default: true)"),
		),
	), WithRecovery(handleExportJson))
}

func handleExportJson(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelExportJsonArguments{}
	if issues := excelExportJsonArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return exportJson(args.FileAbsolutePath, args.SheetName, args.OutputPath, args.Range, args.HeaderRow)
}

func exportJson(fileAbsolutePath string, sheetName string, outputPath string, rangeStr string, headerRow bool) (*mcp.CallToolResult, error) {
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

	var result interface{}
	rowCount := 0

	if headerRow {
		// Read header row
		headers := make([]string, 0, endCol-startCol+1)
		for col := startCol; col <= endCol; col++ {
			cellName, err := excelize.CoordinatesToCellName(col, startRow)
			if err != nil {
				return nil, err
			}
			val, err := worksheet.GetValue(cellName)
			if err != nil {
				return nil, err
			}
			headers = append(headers, val)
		}

		// Read data rows as objects
		var objects []map[string]interface{}
		for row := startRow + 1; row <= endRow; row++ {
			obj := make(map[string]interface{})
			for col := startCol; col <= endCol; col++ {
				cellName, err := excelize.CoordinatesToCellName(col, row)
				if err != nil {
					return nil, err
				}
				val, err := worksheet.GetValue(cellName)
				if err != nil {
					return nil, err
				}
				headerIdx := col - startCol
				if headerIdx < len(headers) {
					obj[headers[headerIdx]] = val
				}
			}
			objects = append(objects, obj)
			rowCount++
		}
		result = objects
	} else {
		// Read all rows as arrays
		var rows [][]string
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
			rows = append(rows, record)
			rowCount++
		}
		result = rows
	}

	// Create output directory if needed
	if err := os.MkdirAll(filepath.Dir(outputPath), os.ModePerm); err != nil {
		return nil, fmt.Errorf("failed to create output directory: %w", err)
	}

	// Write JSON
	jsonData, err := json.MarshalIndent(result, "", "  ")
	if err != nil {
		return nil, fmt.Errorf("failed to marshal JSON: %w", err)
	}
	if err := os.WriteFile(outputPath, jsonData, os.ModePerm); err != nil {
		return nil, fmt.Errorf("failed to write JSON file: %w", err)
	}

	resultText := "# Notice\n"
	resultText += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	resultText += fmt.Sprintf("Exported %d rows from sheet [%s] to %s\n", rowCount, html.EscapeString(sheetName), outputPath)
	return mcp.NewToolResultText(resultText), nil
}
