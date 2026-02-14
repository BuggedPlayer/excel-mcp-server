package excel

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// PagingStrategy defines the interface for calculating paging ranges.
type PagingStrategy interface {
	// CalculatePagingRanges returns a list of available paging ranges.
	CalculatePagingRanges() []string
}

// calculateFixedSizeRanges computes paging ranges for a given dimension and page size.
// This is the shared implementation used by both Excelize and OLE fixed-size strategies.
func calculateFixedSizeRanges(dimension string, pageSize int) []string {
	startCol, startRow, endCol, endRow, err := ParseRange(dimension)
	if err != nil {
		return []string{}
	}

	totalCols := endCol - startCol + 1
	rowsPerPage := pageSize / totalCols
	if rowsPerPage < 1 {
		rowsPerPage = 1
	}

	var ranges []string
	currentRow := startRow
	for currentRow <= endRow {
		pageEndRow := currentRow + rowsPerPage - 1
		if pageEndRow > endRow {
			pageEndRow = endRow
		}

		startRange, err := excelize.CoordinatesToCellName(startCol, currentRow)
		if err != nil {
			return ranges
		}
		endRange, err := excelize.CoordinatesToCellName(endCol, pageEndRow)
		if err != nil {
			return ranges
		}
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))

		currentRow = pageEndRow + 1
	}

	return ranges
}

// ExcelizeFixedSizePagingStrategy calculates paging ranges with a fixed cell count per page.
type ExcelizeFixedSizePagingStrategy struct {
	pageSize  int
	worksheet *ExcelizeWorksheet
	dimension string
}

// NewExcelizeFixedSizePagingStrategy creates a new ExcelizeFixedSizePagingStrategy instance.
func NewExcelizeFixedSizePagingStrategy(pageSize int, worksheet *ExcelizeWorksheet) (*ExcelizeFixedSizePagingStrategy, error) {
	if pageSize <= 0 {
		pageSize = 5000
	}

	dimension, err := worksheet.GetDimension()
	if err != nil {
		return nil, err
	}

	return &ExcelizeFixedSizePagingStrategy{
		pageSize:  pageSize,
		worksheet: worksheet,
		dimension: dimension,
	}, nil
}

// CalculatePagingRanges generates paging ranges based on fixed cell count.
func (s *ExcelizeFixedSizePagingStrategy) CalculatePagingRanges() []string {
	return calculateFixedSizeRanges(s.dimension, s.pageSize)
}

func NewOlePagingStrategy(pageSize int, worksheet *OleWorksheet) (PagingStrategy, error) {
	if worksheet == nil {
		return nil, fmt.Errorf("worksheet is nil")
	}

	printAreaPagingStrategy, err := NewPrintAreaPagingStrategy(worksheet)
	if err != nil {
		return nil, err
	}
	printArea, err := printAreaPagingStrategy.getPrintArea()
	if err != nil {
		return nil, err
	}
	if printArea == "" {
		return NewOleFixedSizePagingStrategy(pageSize, worksheet)
	}
	return printAreaPagingStrategy, nil
}

// OleFixedSizePagingStrategy calculates paging ranges with a fixed cell count per page using OLE.
type OleFixedSizePagingStrategy struct {
	pageSize  int
	worksheet *OleWorksheet
	dimension string
}

// NewOleFixedSizePagingStrategy creates a new OleFixedSizePagingStrategy instance.
func NewOleFixedSizePagingStrategy(pageSize int, worksheet *OleWorksheet) (*OleFixedSizePagingStrategy, error) {
	if pageSize <= 0 {
		pageSize = 5000
	}

	if worksheet == nil {
		return nil, fmt.Errorf("worksheet is nil")
	}

	dimension, err := worksheet.GetDimension()
	if err != nil {
		return nil, fmt.Errorf("failed to get dimension: %w", err)
	}

	return &OleFixedSizePagingStrategy{
		pageSize:  pageSize,
		worksheet: worksheet,
		dimension: dimension,
	}, nil
}

// CalculatePagingRanges generates paging ranges based on fixed cell count.
func (s *OleFixedSizePagingStrategy) CalculatePagingRanges() []string {
	return calculateFixedSizeRanges(s.dimension, s.pageSize)
}

// PrintAreaPagingStrategy calculates paging ranges based on print area and page breaks.
type PrintAreaPagingStrategy struct {
	worksheet *OleWorksheet
}

// NewPrintAreaPagingStrategy creates a new PrintAreaPagingStrategy instance.
func NewPrintAreaPagingStrategy(worksheet *OleWorksheet) (*PrintAreaPagingStrategy, error) {
	if worksheet == nil {
		return nil, fmt.Errorf("worksheet is nil")
	}
	return &PrintAreaPagingStrategy{
		worksheet: worksheet,
	}, nil
}

func (s *PrintAreaPagingStrategy) getPrintArea() (string, error) {
	return s.worksheet.PrintArea()
}

func (s *PrintAreaPagingStrategy) getHPageBreaksPositions() ([]int, error) {
	pageBreaks, err := s.worksheet.HPageBreaks()
	if err != nil {
		return nil, fmt.Errorf("failed to get HPageBreaks: %w", err)
	}
	return pageBreaks, nil
}

func (s *PrintAreaPagingStrategy) calculateRangesFromBreaks(printArea string, breaks []int) []string {
	if printArea == "" {
		return []string{}
	}

	startCol, startRow, endCol, endRow, err := ParseRange(printArea)
	if err != nil {
		return []string{}
	}

	ranges := make([]string, 0)
	currentRow := startRow

	if len(breaks) == 0 {
		startRange, err := excelize.CoordinatesToCellName(startCol, startRow)
		if err != nil {
			return []string{}
		}
		endRange, err := excelize.CoordinatesToCellName(endCol, endRow)
		if err != nil {
			return []string{}
		}
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))
		return ranges
	}

	for _, breakRow := range breaks {
		if breakRow <= startRow || breakRow > endRow {
			continue
		}

		startRange, err := excelize.CoordinatesToCellName(startCol, currentRow)
		if err != nil {
			return ranges
		}
		endRange, err := excelize.CoordinatesToCellName(endCol, breakRow-1)
		if err != nil {
			return ranges
		}
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))

		currentRow = breakRow
	}

	if currentRow <= endRow {
		startRange, err := excelize.CoordinatesToCellName(startCol, currentRow)
		if err != nil {
			return ranges
		}
		endRange, err := excelize.CoordinatesToCellName(endCol, endRow)
		if err != nil {
			return ranges
		}
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))
	}

	return ranges
}

// CalculatePagingRanges generates paging ranges based on print area and page breaks.
func (s *PrintAreaPagingStrategy) CalculatePagingRanges() []string {
	printArea, err := s.getPrintArea()
	if err != nil {
		return []string{}
	}

	breaks, err := s.getHPageBreaksPositions()
	if err != nil {
		return []string{}
	}

	return s.calculateRangesFromBreaks(printArea, breaks)
}

// PagingRangeService provides paging operations.
type PagingRangeService struct {
	strategy PagingStrategy
}

// NewPagingRangeService creates a new PagingRangeService instance.
func NewPagingRangeService(strategy PagingStrategy) *PagingRangeService {
	return &PagingRangeService{strategy: strategy}
}

// GetPagingRanges returns a list of available paging ranges.
func (s *PagingRangeService) GetPagingRanges() []string {
	return s.strategy.CalculatePagingRanges()
}

// FilterRemainingPagingRanges returns ranges that are not in knownRanges.
func (s *PagingRangeService) FilterRemainingPagingRanges(allRanges []string, knownRanges []string) []string {
	if len(knownRanges) == 0 {
		return allRanges
	}

	knownMap := make(map[string]bool)
	for _, r := range knownRanges {
		knownMap[r] = true
	}

	remaining := make([]string, 0)
	for _, r := range allRanges {
		if !knownMap[r] {
			remaining = append(remaining, r)
		}
	}

	return remaining
}

// FindNextRange returns the next range in the sequence after the current range.
func (s *PagingRangeService) FindNextRange(allRanges []string, currentRange string) string {
	for i, r := range allRanges {
		if r == currentRange && i+1 < len(allRanges) {
			return allRanges[i+1]
		}
	}
	return ""
}
