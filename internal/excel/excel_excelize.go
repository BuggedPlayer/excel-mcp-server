package excel

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

type ExcelizeExcel struct {
	file *excelize.File
}

func NewExcelizeExcel(file *excelize.File) Excel {
	return &ExcelizeExcel{file: file}
}

func (e *ExcelizeExcel) GetBackendName() string {
	return "excelize"
}

func (e *ExcelizeExcel) FindSheet(sheetName string) (Worksheet, error) {
	index, err := e.file.GetSheetIndex(sheetName)
	if err != nil {
		return nil, fmt.Errorf("sheet not found: %s: %w", sheetName, err)
	}
	if index < 0 {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}
	return &ExcelizeWorksheet{file: e.file, sheetName: sheetName}, nil
}

func (e *ExcelizeExcel) CreateNewSheet(sheetName string) error {
	_, err := e.file.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to create new sheet: %w", err)
	}
	return nil
}

func (e *ExcelizeExcel) CopySheet(srcSheetName string, destSheetName string) error {
	srcIndex, err := e.file.GetSheetIndex(srcSheetName)
	if err != nil {
		return fmt.Errorf("source sheet not found: %s: %w", srcSheetName, err)
	}
	if srcIndex < 0 {
		return fmt.Errorf("source sheet not found: %s", srcSheetName)
	}
	destIndex, err := e.file.NewSheet(destSheetName)
	if err != nil {
		return fmt.Errorf("failed to create destination sheet: %w", err)
	}
	if err := e.file.CopySheet(srcIndex, destIndex); err != nil {
		return fmt.Errorf("failed to copy sheet: %w", err)
	}
	srcNext := e.file.GetSheetList()[srcIndex+1]
	if srcNext != srcSheetName {
		e.file.MoveSheet(destSheetName, srcNext)
	}
	return nil
}

func (e *ExcelizeExcel) GetSheets() ([]Worksheet, error) {
	sheetList := e.file.GetSheetList()
	worksheets := make([]Worksheet, len(sheetList))
	for i, sheetName := range sheetList {
		worksheets[i] = &ExcelizeWorksheet{file: e.file, sheetName: sheetName}
	}
	return worksheets, nil
}

// SaveExcelize saves the Excel file to the specified path.
// Excelize's Save method restricts the file path length to 207 characters,
// but since this limitation has been relaxed in some environments,
// we ignore this restriction.
// https://github.com/qax-os/excelize/blob/v2.9.0/file.go#L71-L73
func (w *ExcelizeExcel) Save() error {
	file, err := os.OpenFile(filepath.Clean(w.file.Path), os.O_WRONLY|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if err != nil {
		return err
	}
	defer file.Close()
	return w.file.Write(file)
}

type ExcelizeWorksheet struct {
	file      *excelize.File
	sheetName string
}

func (w *ExcelizeWorksheet) Release() {
	// No resources to release in excelize
}

func (w *ExcelizeWorksheet) Name() (string, error) {
	return w.sheetName, nil
}

func (w *ExcelizeWorksheet) GetTables() ([]Table, error) {
	tables, err := w.file.GetTables(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get tables: %w", err)
	}
	tableList := make([]Table, len(tables))
	for i, table := range tables {
		tableList[i] = Table{
			Name:  table.Name,
			Range: NormalizeRange(table.Range),
		}
	}
	return tableList, nil
}

func (w *ExcelizeWorksheet) GetPivotTables() ([]PivotTable, error) {
	pivotTables, err := w.file.GetPivotTables(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get pivot tables: %w", err)
	}
	pivotTableList := make([]PivotTable, len(pivotTables))
	for i, pivotTable := range pivotTables {
		pivotTableList[i] = PivotTable{
			Name:  pivotTable.Name,
			Range: NormalizeRange(pivotTable.PivotTableRange),
		}
	}
	return pivotTableList, nil
}

func (w *ExcelizeWorksheet) SetValue(cell string, value any) error {
	if err := w.file.SetCellValue(w.sheetName, cell, value); err != nil {
		return err
	}
	if err := w.updateDimension(cell); err != nil {
		return fmt.Errorf("failed to update dimension: %w", err)
	}
	return nil
}

func (w *ExcelizeWorksheet) SetFormula(cell string, formula string) error {
	if err := w.file.SetCellFormula(w.sheetName, cell, formula); err != nil {
		return err
	}
	if err := w.updateDimension(cell); err != nil {
		return fmt.Errorf("failed to update dimension: %w", err)
	}
	return nil
}

func (w *ExcelizeWorksheet) GetValue(cell string) (string, error) {
	value, err := w.file.GetCellValue(w.sheetName, cell)
	if err != nil {
		return "", fmt.Errorf("failed to get cell value: %w", err)
	}
	if value == "" {
		// try to get calculated value
		formula, err := w.file.GetCellFormula(w.sheetName, cell)
		if err != nil {
			return "", fmt.Errorf("failed to get formula: %w", err)
		}
		if formula != "" {
			return w.file.CalcCellValue(w.sheetName, cell)
		}
	}
	return value, nil
}

func (w *ExcelizeWorksheet) GetFormula(cell string) (string, error) {
	formula, err := w.file.GetCellFormula(w.sheetName, cell)
	if err != nil {
		return "", fmt.Errorf("failed to get formula: %w", err)
	}
	if formula == "" {
		// fallback
		return w.GetValue(cell)
	}
	if !strings.HasPrefix(formula, "=") {
		formula = "=" + formula
	}
	return formula, nil
}

func (w *ExcelizeWorksheet) GetDimension() (string, error) {
	return w.file.GetSheetDimension(w.sheetName)
}

func (w *ExcelizeWorksheet) GetPagingStrategy(pageSize int) (PagingStrategy, error) {
	return NewExcelizeFixedSizePagingStrategy(pageSize, w)
}

func (w *ExcelizeWorksheet) CapturePicture(captureRange string) (string, error) {
	return "", fmt.Errorf("CapturePicture is not supported in Excelize")
}

func (w *ExcelizeWorksheet) AddTable(tableRange, tableName string) error {
	enable := true
	if err := w.file.AddTable(w.sheetName, &excelize.Table{
		Range:             tableRange,
		Name:              tableName,
		StyleName:         "TableStyleMedium2",
		ShowColumnStripes: true,
		ShowFirstColumn:   false,
		ShowHeaderRow:     &enable,
		ShowLastColumn:    false,
		ShowRowStripes:    &enable,
	}); err != nil {
		return err
	}
	return nil
}

func (w *ExcelizeWorksheet) GetCellStyle(cell string) (*CellStyle, error) {
	styleID, err := w.file.GetCellStyle(w.sheetName, cell)
	if err != nil {
		return nil, fmt.Errorf("failed to get cell style: %w", err)
	}

	style, err := w.file.GetStyle(styleID)
	if err != nil {
		return nil, fmt.Errorf("failed to get style details: %w", err)
	}

	return convertExcelizeStyleToCellStyle(style), nil
}

func (w *ExcelizeWorksheet) SetCellStyle(cell string, style *CellStyle) error {
	excelizeStyle := convertCellStyleToExcelizeStyle(style)

	styleID, err := w.file.NewStyle(excelizeStyle)
	if err != nil {
		return fmt.Errorf("failed to create style: %w", err)
	}

	if err := w.file.SetCellStyle(w.sheetName, cell, cell, styleID); err != nil {
		return fmt.Errorf("failed to set cell style: %w", err)
	}

	return nil
}

func convertCellStyleToExcelizeStyle(style *CellStyle) *excelize.Style {
	result := &excelize.Style{}

	// Border
	if len(style.Border) > 0 {
		borders := make([]excelize.Border, len(style.Border))
		for i, border := range style.Border {
			excelizeBorder := excelize.Border{
				Type: border.Type.String(),
			}
			if border.Color != "" {
				excelizeBorder.Color = strings.TrimPrefix(border.Color, "#")
			}
			excelizeBorder.Style = borderStyleNameToInt(border.Style)
			borders[i] = excelizeBorder
		}
		result.Border = borders
	}

	// Font
	if style.Font != nil {
		font := &excelize.Font{}
		if style.Font.Bold != nil {
			font.Bold = *style.Font.Bold
		}
		if style.Font.Italic != nil {
			font.Italic = *style.Font.Italic
		}
		if style.Font.Underline != nil {
			font.Underline = style.Font.Underline.String()
		}
		if style.Font.Size != nil && *style.Font.Size > 0 {
			font.Size = float64(*style.Font.Size)
		}
		if style.Font.Strike != nil {
			font.Strike = *style.Font.Strike
		}
		if style.Font.Color != nil && *style.Font.Color != "" {
			font.Color = strings.TrimPrefix(*style.Font.Color, "#")
		}
		if style.Font.VertAlign != nil {
			font.VertAlign = style.Font.VertAlign.String()
		}
		result.Font = font
	}

	// Fill
	if style.Fill != nil {
		fill := excelize.Fill{}
		if style.Fill.Type != "" {
			fill.Type = style.Fill.Type.String()
		}
		fill.Pattern = fillPatternNameToInt(style.Fill.Pattern)
		if len(style.Fill.Color) > 0 {
			colors := make([]string, len(style.Fill.Color))
			for i, color := range style.Fill.Color {
				colors[i] = strings.TrimPrefix(color, "#")
			}
			fill.Color = colors
		}
		if style.Fill.Shading != nil {
			fill.Shading = fillShadingNameToInt(*style.Fill.Shading)
		}
		result.Fill = fill
	}

	// NumFmt
	if style.NumFmt != nil && *style.NumFmt != "" {
		result.CustomNumFmt = style.NumFmt
	}

	// DecimalPlaces
	if style.DecimalPlaces != nil && *style.DecimalPlaces > 0 {
		result.DecimalPlaces = style.DecimalPlaces
	}

	return result
}

func convertExcelizeStyleToCellStyle(style *excelize.Style) *CellStyle {
	result := &CellStyle{}

	// Border
	if len(style.Border) > 0 {
		var borders []Border
		for _, border := range style.Border {
			borderStyle := Border{
				Type: BorderType(border.Type),
			}
			if border.Color != "" {
				borderStyle.Color = "#" + strings.ToUpper(border.Color)
			}
			if border.Style != 0 {
				borderStyle.Style = intToBorderStyleName(border.Style)
			}
			borders = append(borders, borderStyle)
		}
		if len(borders) > 0 {
			result.Border = borders
		}
	}

	// Font
	if style.Font != nil {
		font := &FontStyle{}
		if style.Font.Bold {
			font.Bold = &style.Font.Bold
		}
		if style.Font.Italic {
			font.Italic = &style.Font.Italic
		}
		if style.Font.Underline != "" {
			underline := FontUnderline(style.Font.Underline)
			font.Underline = &underline
		}
		if style.Font.Size > 0 {
			size := int(style.Font.Size)
			font.Size = &size
		}
		if style.Font.Strike {
			font.Strike = &style.Font.Strike
		}
		if style.Font.Color != "" {
			color := "#" + strings.ToUpper(style.Font.Color)
			font.Color = &color
		}
		if style.Font.VertAlign != "" {
			vertAlign := FontVertAlign(style.Font.VertAlign)
			font.VertAlign = &vertAlign
		}
		if font.Bold != nil || font.Italic != nil || font.Underline != nil || font.Size != nil || font.Strike != nil || font.Color != nil || font.VertAlign != nil {
			result.Font = font
		}
	}

	// Fill
	if style.Fill.Type != "" || style.Fill.Pattern != 0 || len(style.Fill.Color) > 0 {
		fill := &FillStyle{}
		if style.Fill.Type != "" {
			fill.Type = FillType(style.Fill.Type)
		}
		if style.Fill.Pattern != 0 {
			fill.Pattern = intToFillPatternName(style.Fill.Pattern)
		}
		if len(style.Fill.Color) > 0 {
			var colors []string
			for _, color := range style.Fill.Color {
				if color != "" {
					colors = append(colors, "#"+strings.ToUpper(color))
				}
			}
			if len(colors) > 0 {
				fill.Color = colors
			}
		}
		if style.Fill.Shading != 0 {
			shading := intToFillShadingName(style.Fill.Shading)
			fill.Shading = &shading
		}
		if fill.Type != "" || fill.Pattern != FillPatternNone || len(fill.Color) > 0 || fill.Shading != nil {
			result.Fill = fill
		}
	}

	// NumFmt
	if style.CustomNumFmt != nil && *style.CustomNumFmt != "" {
		result.NumFmt = style.CustomNumFmt
	}

	// DecimalPlaces
	if style.DecimalPlaces != nil && *style.DecimalPlaces != 0 {
		result.DecimalPlaces = style.DecimalPlaces
	}

	return result
}

func intToBorderStyleName(style int) BorderStyle {
	styles := map[int]BorderStyle{
		0:  BorderStyleNone,
		1:  BorderStyleContinuous,
		2:  BorderStyleContinuous,
		3:  BorderStyleDash,
		4:  BorderStyleDot,
		5:  BorderStyleContinuous,
		6:  BorderStyleDouble,
		7:  BorderStyleContinuous,
		8:  BorderStyleDashDot,
		9:  BorderStyleDashDotDot,
		10: BorderStyleSlantDashDot,
		11: BorderStyleContinuous,
		12: BorderStyleMediumDashDot,
		13: BorderStyleMediumDashDotDot,
	}
	if name, exists := styles[style]; exists {
		return name
	}
	return BorderStyleContinuous
}

func intToFillPatternName(pattern int) FillPattern {
	patterns := map[int]FillPattern{
		0:  FillPatternNone,
		1:  FillPatternSolid,
		2:  FillPatternMediumGray,
		3:  FillPatternDarkGray,
		4:  FillPatternLightGray,
		5:  FillPatternDarkHorizontal,
		6:  FillPatternDarkVertical,
		7:  FillPatternDarkDown,
		8:  FillPatternDarkUp,
		9:  FillPatternDarkGrid,
		10: FillPatternDarkTrellis,
		11: FillPatternLightHorizontal,
		12: FillPatternLightVertical,
		13: FillPatternLightDown,
		14: FillPatternLightUp,
		15: FillPatternLightGrid,
		16: FillPatternLightTrellis,
		17: FillPatternGray125,
		18: FillPatternGray0625,
	}
	if name, exists := patterns[pattern]; exists {
		return name
	}
	return FillPatternNone
}

func intToFillShadingName(shading int) FillShading {
	shadings := map[int]FillShading{
		0: FillShadingHorizontal,
		1: FillShadingVertical,
		2: FillShadingDiagonalDown,
		3: FillShadingDiagonalUp,
		4: FillShadingFromCenter,
		5: FillShadingFromCorner,
	}
	if name, exists := shadings[shading]; exists {
		return name
	}
	return FillShadingHorizontal
}

func borderStyleNameToInt(style BorderStyle) int {
	styles := map[BorderStyle]int{
		BorderStyleNone:             0,
		BorderStyleContinuous:       1,
		BorderStyleDash:             3,
		BorderStyleDot:              4,
		BorderStyleDouble:           6,
		BorderStyleDashDot:          8,
		BorderStyleDashDotDot:       9,
		BorderStyleSlantDashDot:     10,
		BorderStyleMediumDashDot:    12,
		BorderStyleMediumDashDotDot: 13,
	}
	if value, exists := styles[style]; exists {
		return value
	}
	return 1
}

func fillPatternNameToInt(pattern FillPattern) int {
	patterns := map[FillPattern]int{
		FillPatternNone:            0,
		FillPatternSolid:           1,
		FillPatternMediumGray:      2,
		FillPatternDarkGray:        3,
		FillPatternLightGray:       4,
		FillPatternDarkHorizontal:  5,
		FillPatternDarkVertical:    6,
		FillPatternDarkDown:        7,
		FillPatternDarkUp:          8,
		FillPatternDarkGrid:        9,
		FillPatternDarkTrellis:     10,
		FillPatternLightHorizontal: 11,
		FillPatternLightVertical:   12,
		FillPatternLightDown:       13,
		FillPatternLightUp:         14,
		FillPatternLightGrid:       15,
		FillPatternLightTrellis:    16,
		FillPatternGray125:         17,
		FillPatternGray0625:        18,
	}
	if value, exists := patterns[pattern]; exists {
		return value
	}
	return 0
}

func fillShadingNameToInt(shading FillShading) int {
	shadings := map[FillShading]int{
		FillShadingHorizontal:   0,
		FillShadingVertical:     1,
		FillShadingDiagonalDown: 2,
		FillShadingDiagonalUp:   3,
		FillShadingFromCenter:   4,
		FillShadingFromCorner:   5,
	}
	if value, exists := shadings[shading]; exists {
		return value
	}
	return 0
}

// updateDimension updates the dimension of the worksheet after a cell is updated.
func (w *ExcelizeWorksheet) updateDimension(updatedCell string) error {
	dimension, err := w.file.GetSheetDimension(w.sheetName)
	if err != nil {
		return err
	}
	startCol, startRow, endCol, endRow, err := ParseRange(dimension)
	if err != nil {
		return err
	}
	updatedCol, updatedRow, err := excelize.CellNameToCoordinates(updatedCell)
	if err != nil {
		return err
	}
	if startCol > updatedCol {
		startCol = updatedCol
	}
	if endCol < updatedCol {
		endCol = updatedCol
	}
	if startRow > updatedRow {
		startRow = updatedRow
	}
	if endRow < updatedRow {
		endRow = updatedRow
	}
	startRange, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		return err
	}
	endRange, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return err
	}
	updatedDimension := fmt.Sprintf("%s:%s", startRange, endRange)
	return w.file.SetSheetDimension(w.sheetName, updatedDimension)
}

func (e *ExcelizeExcel) DeleteSheet(sheetName string) error {
	if err := e.file.DeleteSheet(sheetName); err != nil {
		return fmt.Errorf("failed to delete sheet: %w", err)
	}
	return nil
}

func (e *ExcelizeExcel) RenameSheet(oldName, newName string) error {
	if err := e.file.SetSheetName(oldName, newName); err != nil {
		return fmt.Errorf("failed to rename sheet: %w", err)
	}
	return nil
}

func (w *ExcelizeWorksheet) MergeCells(mergeRange string) error {
	startCol, startRow, endCol, endRow, err := ParseRange(mergeRange)
	if err != nil {
		return err
	}
	topLeft, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		return err
	}
	bottomRight, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return err
	}
	return w.file.MergeCell(w.sheetName, topLeft, bottomRight)
}

func (w *ExcelizeWorksheet) UnmergeCells(mergeRange string) error {
	startCol, startRow, endCol, endRow, err := ParseRange(mergeRange)
	if err != nil {
		return err
	}
	topLeft, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		return err
	}
	bottomRight, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return err
	}
	return w.file.UnmergeCell(w.sheetName, topLeft, bottomRight)
}

func (w *ExcelizeWorksheet) SetColumnWidth(startCol, endCol string, width float64) error {
	return w.file.SetColWidth(w.sheetName, startCol, endCol, width)
}

func (w *ExcelizeWorksheet) SetRowHeight(row int, height float64) error {
	return w.file.SetRowHeight(w.sheetName, row, height)
}

func (w *ExcelizeWorksheet) InsertRows(row int, count int) error {
	return w.file.InsertRows(w.sheetName, row, count)
}

func (w *ExcelizeWorksheet) DeleteRows(row int, count int) error {
	for i := 0; i < count; i++ {
		if err := w.file.RemoveRow(w.sheetName, row); err != nil {
			return fmt.Errorf("failed to delete row %d: %w", row, err)
		}
	}
	return nil
}

func (w *ExcelizeWorksheet) InsertColumns(column string, count int) error {
	return w.file.InsertCols(w.sheetName, column, count)
}

func (w *ExcelizeWorksheet) DeleteColumns(column string, count int) error {
	for i := 0; i < count; i++ {
		if err := w.file.RemoveCol(w.sheetName, column); err != nil {
			return fmt.Errorf("failed to delete column %s (iteration %d): %w", column, i+1, err)
		}
	}
	return nil
}

func (w *ExcelizeWorksheet) AddChart(position string, chartType string, dataRange string, title string) error {
	ct := mapExcelizeChartType(chartType)
	chart := &excelize.Chart{
		Type: ct,
		Series: []excelize.ChartSeries{
			{
				Name:   "",
				Values: dataRange,
			},
		},
		Title: []excelize.RichTextRun{{Text: title}},
	}
	return w.file.AddChart(w.sheetName, position, chart)
}

func mapExcelizeChartType(chartType string) excelize.ChartType {
	switch chartType {
	case "col":
		return excelize.Col
	case "bar":
		return excelize.Bar
	case "line":
		return excelize.Line
	case "pie":
		return excelize.Pie
	case "area":
		return excelize.Area
	case "scatter":
		return excelize.Scatter
	default:
		return excelize.Col
	}
}

func (w *ExcelizeWorksheet) FreezePanes(cell string) error {
	col, row, err := excelize.CellNameToCoordinates(cell)
	if err != nil {
		return err
	}
	return w.file.SetPanes(w.sheetName, &excelize.Panes{
		Freeze:      true,
		Split:       false,
		XSplit:      col - 1,
		YSplit:      row - 1,
		TopLeftCell: cell,
		ActivePane:  "bottomRight",
	})
}

func (w *ExcelizeWorksheet) AddDataValidation(validationRange string, validationType string, formula1 string, formula2 string, allowBlank bool) error {
	dv := excelize.NewDataValidation(allowBlank)
	dv.Sqref = validationRange
	switch validationType {
	case "list":
		dv.SetDropList(strings.Split(formula1, ","))
	case "whole":
		dv.SetRange(formula1, formula2, excelize.DataValidationTypeWhole, excelize.DataValidationOperatorBetween)
	case "decimal":
		dv.SetRange(formula1, formula2, excelize.DataValidationTypeDecimal, excelize.DataValidationOperatorBetween)
	default:
		return fmt.Errorf("unsupported validation type: %s", validationType)
	}
	return w.file.AddDataValidation(w.sheetName, dv)
}

func (w *ExcelizeWorksheet) FindReplace(searchRange string, find string, replace string, matchCase bool, matchEntireCell bool) (int, error) {
	rangeStr := searchRange
	if rangeStr == "" {
		dim, err := w.GetDimension()
		if err != nil {
			return 0, err
		}
		rangeStr = dim
	}
	startCol, startRow, endCol, endRow, err := ParseRange(rangeStr)
	if err != nil {
		return 0, err
	}
	count := 0
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				continue
			}
			value, err := w.file.GetCellValue(w.sheetName, cell)
			if err != nil || value == "" {
				continue
			}
			var newValue string
			var matched bool
			if matchEntireCell {
				if matchCase {
					matched = value == find
				} else {
					matched = strings.EqualFold(value, find)
				}
				if matched {
					newValue = replace
				}
			} else {
				if matchCase {
					if strings.Contains(value, find) {
						newValue = strings.ReplaceAll(value, find, replace)
						matched = true
					}
				} else {
					lower := strings.ToLower(value)
					lowerFind := strings.ToLower(find)
					if strings.Contains(lower, lowerFind) {
						// Case-insensitive replace
						result := value
						idx := 0
						for {
							pos := strings.Index(strings.ToLower(result[idx:]), lowerFind)
							if pos < 0 {
								break
							}
							result = result[:idx+pos] + replace + result[idx+pos+len(find):]
							idx = idx + pos + len(replace)
						}
						newValue = result
						matched = true
					}
				}
			}
			if matched {
				if err := w.file.SetCellValue(w.sheetName, cell, newValue); err != nil {
					return count, err
				}
				count++
			}
		}
	}
	return count, nil
}

func (w *ExcelizeWorksheet) AddComment(cell string, author string, text string) error {
	return w.file.AddComment(w.sheetName, excelize.Comment{
		Cell:   cell,
		Author: author,
		Text:   text,
	})
}

func (w *ExcelizeWorksheet) GetComments() ([]Comment, error) {
	comments, err := w.file.GetComments(w.sheetName)
	if err != nil {
		return nil, err
	}
	result := make([]Comment, len(comments))
	for i, c := range comments {
		result[i] = Comment{Cell: c.Cell, Author: c.Author, Text: c.Text}
	}
	return result, nil
}

func (w *ExcelizeWorksheet) AddHyperlink(cell string, url string, display string) error {
	linkType := "External"
	if strings.Contains(url, "!") && !strings.HasPrefix(url, "http") {
		linkType = "Location"
	}
	opts := excelize.HyperlinkOpts{}
	if display != "" {
		opts.Display = &display
	}
	if err := w.file.SetCellHyperLink(w.sheetName, cell, url, linkType, opts); err != nil {
		return err
	}
	if display != "" {
		return w.file.SetCellValue(w.sheetName, cell, display)
	}
	return nil
}

func (w *ExcelizeWorksheet) SetConditionalFormat(formatRange string, ruleType string, criteria string, value string, value2 string, fontColor string, bgColor string) error {
	opt := excelize.ConditionalFormatOptions{}

	switch ruleType {
	case "cell":
		opt.Type = "cell"
		opt.Criteria = criteria
		opt.Value = value
		if value2 != "" {
			opt.MinValue = value
			opt.MaxValue = value2
		}
	case "top":
		opt.Type = "top"
		opt.Criteria = criteria
		opt.Value = value
	case "duplicate":
		opt.Type = "duplicate"
	case "colorScale":
		opt.Type = "2_color_scale"
		if value != "" {
			opt.MinValue = value
		}
		if value2 != "" {
			opt.MaxValue = value2
		}
	case "dataBar":
		opt.Type = "data_bar"
		if value != "" {
			opt.MinValue = value
		}
		if value2 != "" {
			opt.MaxValue = value2
		}
	default:
		return fmt.Errorf("unsupported conditional format type: %s", ruleType)
	}

	if fontColor != "" || bgColor != "" {
		style := &excelize.Style{}
		if fontColor != "" {
			style.Font = &excelize.Font{Color: fontColor}
		}
		if bgColor != "" {
			style.Fill = excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{bgColor}}
		}
		styleID, err := w.file.NewConditionalStyle(style)
		if err != nil {
			return fmt.Errorf("failed to create conditional style: %w", err)
		}
		opt.Format = &styleID
	}

	return w.file.SetConditionalFormat(w.sheetName, formatRange, []excelize.ConditionalFormatOptions{opt})
}

func (e *ExcelizeExcel) SetDefinedName(name string, refersTo string, scope string) error {
	return e.file.SetDefinedName(&excelize.DefinedName{
		Name:     name,
		RefersTo: refersTo,
		Scope:    scope,
	})
}

func (e *ExcelizeExcel) GetDefinedNames() ([]DefinedName, error) {
	names := e.file.GetDefinedName()
	result := make([]DefinedName, len(names))
	for i, n := range names {
		result[i] = DefinedName{Name: n.Name, RefersTo: n.RefersTo, Scope: n.Scope}
	}
	return result, nil
}
