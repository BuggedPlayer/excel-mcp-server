package excel

// ExcelApp represents a connection to the Excel application itself,
// as opposed to a specific workbook file. This is used for live Excel
// interaction features like listing open workbooks, opening files in
// Excel, and running macros.
type ExcelApp interface {
	// ListWorkbooks returns information about all currently open workbooks.
	ListWorkbooks() ([]WorkbookInfo, error)
	// OpenWorkbook opens a workbook file in Excel and makes it visible.
	OpenWorkbook(absolutePath string) error
	// CreateWorkbook creates a new workbook in Excel and saves it to the specified path.
	CreateWorkbook(absolutePath string) error
	// RunMacro runs a VBA macro in the specified workbook.
	RunMacro(macroName string, args []string) (string, error)
	// Release releases the COM resources associated with this ExcelApp.
	Release()
}

// WorkbookInfo contains information about an open workbook.
type WorkbookInfo struct {
	Name     string `json:"name"`
	FullPath string `json:"fullPath"`
	Saved    bool   `json:"saved"`
}

// NewExcelApp creates a new ExcelApp instance connected to the running
// Excel application. On non-Windows platforms, this will return an error.
func NewExcelApp() (ExcelApp, error) {
	return newExcelAppPlatform()
}
