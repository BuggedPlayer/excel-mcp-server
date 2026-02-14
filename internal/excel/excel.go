package excel

import (
	"os"

	"github.com/xuri/excelize/v2"
)

type Excel interface {
	// GetBackendName returns the backend used to manipulate the Excel file.
	GetBackendName() string
	// GetSheets returns a list of all worksheets in the Excel file.
	GetSheets() ([]Worksheet, error)
	// FindSheet finds a sheet by its name and returns a Worksheet.
	FindSheet(sheetName string) (Worksheet, error)
	// CreateNewSheet creates a new sheet with the specified name.
	CreateNewSheet(sheetName string) error
	// CopySheet copies a sheet from one to another.
	CopySheet(srcSheetName, destSheetName string) error
	// DeleteSheet deletes a sheet by its name.
	DeleteSheet(sheetName string) error
	// RenameSheet renames a sheet from oldName to newName.
	RenameSheet(oldName, newName string) error
	// SetDefinedName creates or updates a named range in the workbook.
	SetDefinedName(name string, refersTo string, scope string) error
	// GetDefinedNames returns all defined names in the workbook.
	GetDefinedNames() ([]DefinedName, error)
	// Save saves the Excel file.
	Save() error
}

type Worksheet interface {
	// Release releases the worksheet resources.
	Release()
	// Name returns the name of the worksheet.
	Name() (string, error)
	// GetTable returns a tables in this worksheet.
	GetTables() ([]Table, error)
	// GetPivotTable returns a pivot tables in this worksheet.
	GetPivotTables() ([]PivotTable, error)
	// SetValue sets a value in the specified cell.
	SetValue(cell string, value any) error
	// SetFormula sets a formula in the specified cell.
	SetFormula(cell string, formula string) error
	// GetValue gets the value from the specified cell.
	GetValue(cell string) (string, error)
	// GetFormula gets the formula from the specified cell.
	GetFormula(cell string) (string, error)
	// GetDimension gets the dimension of the worksheet.
	GetDimension() (string, error)
	// GetPagingStrategy returns the paging strategy for the worksheet.
	// The pageSize parameter is used to determine the max size of each page.
	GetPagingStrategy(pageSize int) (PagingStrategy, error)
	// CapturePicture returns base64 encoded image data of the specified range.
	CapturePicture(captureRange string) (string, error)
	// AddTable adds a table to this worksheet.
	AddTable(tableRange, tableName string) error
	// GetCellStyle gets style information for the specified cell.
	GetCellStyle(cell string) (*CellStyle, error)
	// SetCellStyle sets style for the specified cell.
	SetCellStyle(cell string, style *CellStyle) error
	// MergeCells merges cells in the specified range.
	MergeCells(mergeRange string) error
	// UnmergeCells unmerges cells in the specified range.
	UnmergeCells(mergeRange string) error
	// SetColumnWidth sets the width of columns in the specified range.
	SetColumnWidth(startCol, endCol string, width float64) error
	// SetRowHeight sets the height of the specified row.
	SetRowHeight(row int, height float64) error
	// InsertRows inserts empty rows at the specified position.
	InsertRows(row int, count int) error
	// DeleteRows deletes rows at the specified position.
	DeleteRows(row int, count int) error
	// InsertColumns inserts empty columns at the specified position.
	InsertColumns(column string, count int) error
	// DeleteColumns deletes columns at the specified position.
	DeleteColumns(column string, count int) error
	// AddChart adds a chart to the worksheet.
	AddChart(position string, chartType string, dataRange string, title string) error
	// FreezePanes freezes rows and columns at the specified cell.
	FreezePanes(cell string) error
	// AddDataValidation adds a data validation rule to the specified range.
	AddDataValidation(validationRange string, validationType string, formula1 string, formula2 string, allowBlank bool) error
	// FindReplace finds and replaces values in the worksheet. Returns the number of replacements made.
	FindReplace(searchRange string, find string, replace string, matchCase bool, matchEntireCell bool) (int, error)
	// AddComment adds a comment to the specified cell.
	AddComment(cell string, author string, text string) error
	// GetComments returns all comments in the worksheet.
	GetComments() ([]Comment, error)
	// AddHyperlink adds a hyperlink to the specified cell.
	AddHyperlink(cell string, url string, display string) error
	// SetConditionalFormat applies conditional formatting to the specified range.
	SetConditionalFormat(formatRange string, ruleType string, criteria string, value string, value2 string, fontColor string, bgColor string) error
}

type Table struct {
	Name  string
	Range string
}

type PivotTable struct {
	Name  string
	Range string
}

type Comment struct {
	Cell   string
	Author string
	Text   string
}

type DefinedName struct {
	Name     string
	RefersTo string
	Scope    string
}

type CellStyle struct {
	Border        []Border   `yaml:"border,omitempty"`
	Font          *FontStyle `yaml:"font,omitempty"`
	Fill          *FillStyle `yaml:"fill,omitempty"`
	NumFmt        *string    `yaml:"numFmt,omitempty"`
	DecimalPlaces *int       `yaml:"decimalPlaces,omitempty"`
}

type Border struct {
	Type  BorderType  `yaml:"type"`
	Style BorderStyle `yaml:"style,omitempty"`
	Color string      `yaml:"color,omitempty"`
}

type FontStyle struct {
	Bold      *bool          `yaml:"bold,omitempty"`
	Italic    *bool          `yaml:"italic,omitempty"`
	Underline *FontUnderline `yaml:"underline,omitempty"`
	Size      *int           `yaml:"size,omitempty"`
	Strike    *bool          `yaml:"strike,omitempty"`
	Color     *string        `yaml:"color,omitempty"`
	VertAlign *FontVertAlign `yaml:"vertAlign,omitempty"`
}

type FillStyle struct {
	Type    FillType     `yaml:"type,omitempty"`
	Pattern FillPattern  `yaml:"pattern,omitempty"`
	Color   []string     `yaml:"color,omitempty"`
	Shading *FillShading `yaml:"shading,omitempty"`
}

// OpenFile opens an Excel file and returns an Excel interface.
// It first tries to open the file using OLE automation, and if that fails,
// it tries to using the excelize library.
// If the file does not exist, it creates a new file using excelize.
func OpenFile(absoluteFilePath string) (Excel, func(), error) {
	ole, releaseFn, err := NewExcelOle(absoluteFilePath)
	if err == nil {
		return ole, releaseFn, nil
	}
	// If OLE fails, try Excelize
	workbook, err := excelize.OpenFile(absoluteFilePath)
	if err != nil {
		// If the file does not exist, create a new one
		if os.IsNotExist(err) {
			workbook = excelize.NewFile()
			workbook.Path = absoluteFilePath
			e := NewExcelizeExcel(workbook)
			return e, func() {
				workbook.Close()
			}, nil
		}
		return nil, func() {}, err
	}
	e := NewExcelizeExcel(workbook)
	return e, func() {
		workbook.Close()
	}, nil
}

// BorderType represents border direction
type BorderType string

const (
	BorderTypeLeft         BorderType = "left"
	BorderTypeRight        BorderType = "right"
	BorderTypeTop          BorderType = "top"
	BorderTypeBottom       BorderType = "bottom"
	BorderTypeDiagonalDown BorderType = "diagonalDown"
	BorderTypeDiagonalUp   BorderType = "diagonalUp"
)

func (b BorderType) String() string {
	return string(b)
}

func (b BorderType) MarshalText() ([]byte, error) {
	return []byte(b.String()), nil
}

func BorderTypeValues() []BorderType {
	return []BorderType{
		BorderTypeLeft,
		BorderTypeRight,
		BorderTypeTop,
		BorderTypeBottom,
		BorderTypeDiagonalDown,
		BorderTypeDiagonalUp,
	}
}

// BorderStyle represents border style constants
type BorderStyle string

const (
	BorderStyleNone             BorderStyle = "none"
	BorderStyleContinuous       BorderStyle = "continuous"
	BorderStyleDash             BorderStyle = "dash"
	BorderStyleDot              BorderStyle = "dot"
	BorderStyleDouble           BorderStyle = "double"
	BorderStyleDashDot          BorderStyle = "dashDot"
	BorderStyleDashDotDot       BorderStyle = "dashDotDot"
	BorderStyleSlantDashDot     BorderStyle = "slantDashDot"
	BorderStyleMediumDashDot    BorderStyle = "mediumDashDot"
	BorderStyleMediumDashDotDot BorderStyle = "mediumDashDotDot"
)

func (b BorderStyle) String() string {
	return string(b)
}

func (b BorderStyle) MarshalText() ([]byte, error) {
	return []byte(b.String()), nil
}

func BorderStyleValues() []BorderStyle {
	return []BorderStyle{
		BorderStyleNone,
		BorderStyleContinuous,
		BorderStyleDash,
		BorderStyleDot,
		BorderStyleDouble,
		BorderStyleDashDot,
		BorderStyleDashDotDot,
		BorderStyleSlantDashDot,
		BorderStyleMediumDashDot,
		BorderStyleMediumDashDotDot,
	}
}

// FontUnderline represents underline styles for font
type FontUnderline string

const (
	FontUnderlineNone             FontUnderline = "none"
	FontUnderlineSingle           FontUnderline = "single"
	FontUnderlineDouble           FontUnderline = "double"
	FontUnderlineSingleAccounting FontUnderline = "singleAccounting"
	FontUnderlineDoubleAccounting FontUnderline = "doubleAccounting"
)

func (f FontUnderline) String() string {
	return string(f)
}
func (f FontUnderline) MarshalText() ([]byte, error) {
	return []byte(f.String()), nil
}

func FontUnderlineValues() []FontUnderline {
	return []FontUnderline{
		FontUnderlineNone,
		FontUnderlineSingle,
		FontUnderlineDouble,
		FontUnderlineSingleAccounting,
		FontUnderlineDoubleAccounting,
	}
}

// FontVertAlign represents vertical alignment options for font styles
type FontVertAlign string

const (
	FontVertAlignBaseline    FontVertAlign = "baseline"
	FontVertAlignSuperscript FontVertAlign = "superscript"
	FontVertAlignSubscript   FontVertAlign = "subscript"
)

func (v FontVertAlign) String() string {
	return string(v)
}

func (v FontVertAlign) MarshalText() ([]byte, error) {
	return []byte(v.String()), nil
}

func FontVertAlignValues() []FontVertAlign {
	return []FontVertAlign{
		FontVertAlignBaseline,
		FontVertAlignSuperscript,
		FontVertAlignSubscript,
	}
}

// FillType represents fill types for cell styles
type FillType string

const (
	FillTypeGradient FillType = "gradient"
	FillTypePattern  FillType = "pattern"
)

func (f FillType) String() string {
	return string(f)
}

func (f FillType) MarshalText() ([]byte, error) {
	return []byte(f.String()), nil
}

func FillTypeValues() []FillType {
	return []FillType{
		FillTypeGradient,
		FillTypePattern,
	}
}

// FillPattern represents fill pattern constants
type FillPattern string

const (
	FillPatternNone            FillPattern = "none"
	FillPatternSolid           FillPattern = "solid"
	FillPatternMediumGray      FillPattern = "mediumGray"
	FillPatternDarkGray        FillPattern = "darkGray"
	FillPatternLightGray       FillPattern = "lightGray"
	FillPatternDarkHorizontal  FillPattern = "darkHorizontal"
	FillPatternDarkVertical    FillPattern = "darkVertical"
	FillPatternDarkDown        FillPattern = "darkDown"
	FillPatternDarkUp          FillPattern = "darkUp"
	FillPatternDarkGrid        FillPattern = "darkGrid"
	FillPatternDarkTrellis     FillPattern = "darkTrellis"
	FillPatternLightHorizontal FillPattern = "lightHorizontal"
	FillPatternLightVertical   FillPattern = "lightVertical"
	FillPatternLightDown       FillPattern = "lightDown"
	FillPatternLightUp         FillPattern = "lightUp"
	FillPatternLightGrid       FillPattern = "lightGrid"
	FillPatternLightTrellis    FillPattern = "lightTrellis"
	FillPatternGray125         FillPattern = "gray125"
	FillPatternGray0625        FillPattern = "gray0625"
)

func (f FillPattern) String() string {
	return string(f)
}

func (f FillPattern) MarshalText() ([]byte, error) {
	return []byte(f.String()), nil
}

func FillPatternValues() []FillPattern {
	return []FillPattern{
		FillPatternNone,
		FillPatternSolid,
		FillPatternMediumGray,
		FillPatternDarkGray,
		FillPatternLightGray,
		FillPatternDarkHorizontal,
		FillPatternDarkVertical,
		FillPatternDarkDown,
		FillPatternDarkUp,
		FillPatternDarkGrid,
		FillPatternDarkTrellis,
		FillPatternLightHorizontal,
		FillPatternLightVertical,
		FillPatternLightDown,
		FillPatternLightUp,
		FillPatternLightGrid,
		FillPatternLightTrellis,
		FillPatternGray125,
		FillPatternGray0625,
	}
}

// FillShading represents fill shading constants
type FillShading string

const (
	FillShadingHorizontal   FillShading = "horizontal"
	FillShadingVertical     FillShading = "vertical"
	FillShadingDiagonalDown FillShading = "diagonalDown"
	FillShadingDiagonalUp   FillShading = "diagonalUp"
	FillShadingFromCenter   FillShading = "fromCenter"
	FillShadingFromCorner   FillShading = "fromCorner"
)

func (f FillShading) String() string {
	return string(f)
}

func (f FillShading) MarshalText() ([]byte, error) {
	return []byte(f.String()), nil
}

func FillShadingValues() []FillShading {
	return []FillShading{
		FillShadingHorizontal,
		FillShadingVertical,
		FillShadingDiagonalDown,
		FillShadingDiagonalUp,
		FillShadingFromCenter,
		FillShadingFromCorner,
	}
}
