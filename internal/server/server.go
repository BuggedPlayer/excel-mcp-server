package server

import (
	"runtime"

	"github.com/mark3labs/mcp-go/server"
	"github.com/negokaz/excel-mcp-server/internal/tools"
)

type ExcelServer struct {
	server *server.MCPServer
}

func New(version string) *ExcelServer {
	s := &ExcelServer{}
	s.server = server.NewMCPServer(
		"excel-mcp-server",
		version,
	)
	tools.AddExcelDescribeSheetsTool(s.server)
	tools.AddExcelReadSheetTool(s.server)
	if runtime.GOOS == "windows" {
		tools.AddExcelScreenCaptureTool(s.server)
	}
	tools.AddExcelWriteToSheetTool(s.server)
	tools.AddExcelCreateTableTool(s.server)
	tools.AddExcelCopySheetTool(s.server)
	tools.AddExcelFormatRangeTool(s.server)
	tools.AddExcelDeleteSheetTool(s.server)
	tools.AddExcelRenameSheetTool(s.server)
	tools.AddExcelMergeCellsTool(s.server)
	tools.AddExcelUnmergeCellsTool(s.server)
	tools.AddExcelSetColumnWidthTool(s.server)
	tools.AddExcelSetRowHeightTool(s.server)
	tools.AddExcelInsertRowsTool(s.server)
	tools.AddExcelDeleteRowsTool(s.server)
	tools.AddExcelInsertColumnsTool(s.server)
	tools.AddExcelDeleteColumnsTool(s.server)
	tools.AddExcelAddChartTool(s.server)
	tools.AddExcelFreezePanesTool(s.server)
	tools.AddExcelAddDataValidationTool(s.server)
	tools.AddExcelFindReplaceTool(s.server)
	// Phase 1: Formula Engine tools
	tools.AddExcelAddCommentTool(s.server)
	tools.AddExcelGetCommentsTool(s.server)
	tools.AddExcelAddHyperlinkTool(s.server)
	tools.AddExcelSetNamedRangeTool(s.server)
	tools.AddExcelSetConditionalFormatTool(s.server)
	// Phase 2: Live Excel Control tools (Windows only)
	if runtime.GOOS == "windows" {
		tools.AddExcelListWorkbooksTool(s.server)
		tools.AddExcelOpenWorkbookTool(s.server)
		tools.AddExcelCreateWorkbookTool(s.server)
		tools.AddExcelRunMacroTool(s.server)
	}
	// Phase 3: Import/Export tools
	tools.AddExcelExportCsvTool(s.server)
	tools.AddExcelImportCsvTool(s.server)
	tools.AddExcelExportJsonTool(s.server)
	tools.AddExcelImportJsonTool(s.server)
	return s
}

func (s *ExcelServer) Start() error {
	return server.ServeStdio(s.server)
}
