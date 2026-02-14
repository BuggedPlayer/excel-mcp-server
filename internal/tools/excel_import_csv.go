package tools

import (
	"context"
	"encoding/csv"
	"fmt"
	"html"
	"io"
	"os"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
	"github.com/xuri/excelize/v2"
)

type ExcelImportCsvArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	CsvPath          string `zog:"csvPath"`
	StartCell        string `zog:"startCell"`
	Delimiter        string `zog:"delimiter"`
	NewSheet         bool   `zog:"newSheet"`
}

var excelImportCsvArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"csvPath":          z.String().Test(AbsolutePathTest()).Required(),
	"startCell":        z.String().Default("A1"),
	"delimiter":        z.String().Default(","),
	"newSheet":         z.Bool().Default(false),
})

func AddExcelImportCsvTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_import_csv",
		mcp.WithDescription("Import a CSV file into an Excel sheet"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("csvPath",
			mcp.Required(),
			mcp.Description("Absolute path to the CSV file to import"),
		),
		mcp.WithString("startCell",
			mcp.Description("Cell where data import starts (default: \"A1\")"),
		),
		mcp.WithString("delimiter",
			mcp.Description("CSV delimiter character (default: \",\")"),
		),
		mcp.WithBoolean("newSheet",
			mcp.Description("Create a new sheet if true (default: false)"),
		),
	), WithRecovery(handleImportCsv))
}

func handleImportCsv(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelImportCsvArguments{}
	if issues := excelImportCsvArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return importCsv(args.FileAbsolutePath, args.SheetName, args.CsvPath, args.StartCell, args.Delimiter, args.NewSheet)
}

func importCsv(fileAbsolutePath string, sheetName string, csvPath string, startCell string, delimiter string, newSheet bool) (*mcp.CallToolResult, error) {
	// Open and read CSV file
	csvFile, err := os.Open(csvPath)
	if err != nil {
		return nil, fmt.Errorf("failed to open CSV file: %w", err)
	}
	defer csvFile.Close()

	reader := csv.NewReader(csvFile)
	if len(delimiter) > 0 {
		reader.Comma = rune(delimiter[0])
	}
	reader.FieldsPerRecord = -1 // Allow variable number of fields

	// Read all records
	var records [][]string
	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("failed to read CSV: %w", err)
		}
		records = append(records, record)
	}

	if len(records) == 0 {
		return imcp.NewToolResultInvalidArgumentError("CSV file is empty"), nil
	}

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

	// Write records to cells
	for i, record := range records {
		for j, value := range record {
			cellName, err := excelize.CoordinatesToCellName(startCol+j, startRow+i)
			if err != nil {
				return nil, err
			}
			if err := worksheet.SetValue(cellName, value); err != nil {
				return nil, err
			}
		}
	}

	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Imported %d rows from %s into sheet [%s] starting at %s\n", len(records), csvPath, html.EscapeString(sheetName), startCell)
	return mcp.NewToolResultText(result), nil
}
