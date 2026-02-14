package tools

import (
	"context"
	"encoding/json"
	"fmt"
	"html"
	"os"
	"sort"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
	"github.com/xuri/excelize/v2"
)

type ExcelImportJsonArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	JsonPath         string `zog:"jsonPath"`
	StartCell        string `zog:"startCell"`
	HeaderRow        bool   `zog:"headerRow"`
	NewSheet         bool   `zog:"newSheet"`
}

var excelImportJsonArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"jsonPath":         z.String().Test(AbsolutePathTest()).Required(),
	"startCell":        z.String().Default("A1"),
	"headerRow":        z.Bool().Default(true),
	"newSheet":         z.Bool().Default(false),
})

func AddExcelImportJsonTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_import_json",
		mcp.WithDescription("Import a JSON file into an Excel sheet"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("jsonPath",
			mcp.Required(),
			mcp.Description("Absolute path to the JSON file to import"),
		),
		mcp.WithString("startCell",
			mcp.Description("Cell where data import starts (default: \"A1\")"),
		),
		mcp.WithBoolean("headerRow",
			mcp.Description("Write keys as header row (default: true). Only applies when JSON is an array of objects"),
		),
		mcp.WithBoolean("newSheet",
			mcp.Description("Create a new sheet if true (default: false)"),
		),
	), WithRecovery(handleImportJson))
}

func handleImportJson(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelImportJsonArguments{}
	if issues := excelImportJsonArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return importJson(args.FileAbsolutePath, args.SheetName, args.JsonPath, args.StartCell, args.HeaderRow, args.NewSheet)
}

func importJson(fileAbsolutePath string, sheetName string, jsonPath string, startCell string, headerRow bool, newSheet bool) (*mcp.CallToolResult, error) {
	// Read JSON file
	jsonData, err := os.ReadFile(jsonPath)
	if err != nil {
		return nil, fmt.Errorf("failed to read JSON file: %w", err)
	}

	// Try to parse as array of objects first
	var objects []map[string]interface{}
	if err := json.Unmarshal(jsonData, &objects); err == nil && len(objects) > 0 {
		return importJsonObjects(fileAbsolutePath, sheetName, startCell, headerRow, newSheet, objects, jsonPath)
	}

	// Try to parse as array of arrays
	var arrays [][]interface{}
	if err := json.Unmarshal(jsonData, &arrays); err == nil {
		return importJsonArrays(fileAbsolutePath, sheetName, startCell, newSheet, arrays, jsonPath)
	}

	return imcp.NewToolResultInvalidArgumentError("JSON must be an array of objects or an array of arrays"), nil
}

func importJsonObjects(fileAbsolutePath string, sheetName string, startCell string, headerRow bool, newSheet bool, objects []map[string]interface{}, jsonPath string) (*mcp.CallToolResult, error) {
	// Parse start cell
	startCol, startRow, err := excelize.CellNameToCoordinates(startCell)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("invalid start cell: %s", startCell)), nil
	}

	// Collect all keys maintaining a stable order
	keySet := make(map[string]bool)
	for _, obj := range objects {
		for k := range obj {
			keySet[k] = true
		}
	}
	keys := make([]string, 0, len(keySet))
	for k := range keySet {
		keys = append(keys, k)
	}
	sort.Strings(keys)

	// Open Excel file
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer release()

	if newSheet {
		if err := workbook.CreateNewSheet(sheetName); err != nil {
			return nil, err
		}
	}

	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	defer worksheet.Release()

	currentRow := startRow
	rowCount := 0

	// Write header row
	if headerRow {
		for j, key := range keys {
			cellName, err := excelize.CoordinatesToCellName(startCol+j, currentRow)
			if err != nil {
				return nil, err
			}
			if err := worksheet.SetValue(cellName, key); err != nil {
				return nil, err
			}
		}
		currentRow++
	}

	// Write data rows
	for _, obj := range objects {
		for j, key := range keys {
			cellName, err := excelize.CoordinatesToCellName(startCol+j, currentRow)
			if err != nil {
				return nil, err
			}
			val := obj[key]
			if val == nil {
				val = ""
			}
			if err := worksheet.SetValue(cellName, val); err != nil {
				return nil, err
			}
		}
		currentRow++
		rowCount++
	}

	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Imported %d rows from %s into sheet [%s] starting at %s\n", rowCount, jsonPath, html.EscapeString(sheetName), startCell)
	return mcp.NewToolResultText(result), nil
}

func importJsonArrays(fileAbsolutePath string, sheetName string, startCell string, newSheet bool, arrays [][]interface{}, jsonPath string) (*mcp.CallToolResult, error) {
	// Parse start cell
	startCol, startRow, err := excelize.CellNameToCoordinates(startCell)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("invalid start cell: %s", startCell)), nil
	}

	// Open Excel file
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer release()

	if newSheet {
		if err := workbook.CreateNewSheet(sheetName); err != nil {
			return nil, err
		}
	}

	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	defer worksheet.Release()

	for i, row := range arrays {
		for j, val := range row {
			cellName, err := excelize.CoordinatesToCellName(startCol+j, startRow+i)
			if err != nil {
				return nil, err
			}
			if val == nil {
				val = ""
			}
			if err := worksheet.SetValue(cellName, val); err != nil {
				return nil, err
			}
		}
	}

	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Imported %d rows from %s into sheet [%s] starting at %s\n", len(arrays), jsonPath, html.EscapeString(sheetName), startCell)
	return mcp.NewToolResultText(result), nil
}
