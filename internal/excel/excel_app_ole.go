package excel

import (
	"fmt"
	"runtime"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type OleExcelApp struct {
	application *ole.IDispatch
}

func newExcelAppPlatform() (ExcelApp, error) {
	runtime.LockOSThread()
	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)

	// Try to connect to an already-running Excel instance first
	unknown, err := oleutil.GetActiveObject("Excel.Application")
	if err == nil {
		excel, err := unknown.QueryInterface(ole.IID_IDispatch)
		if err == nil {
			return &OleExcelApp{application: excel}, nil
		}
	}

	// Excel is not running — launch a new instance
	unknown2, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		ole.CoUninitialize()
		runtime.UnlockOSThread()
		return nil, fmt.Errorf("failed to launch Excel application: %w", err)
	}
	excel, err := unknown2.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		ole.CoUninitialize()
		runtime.UnlockOSThread()
		return nil, fmt.Errorf("failed to query Excel interface: %w", err)
	}
	// Make the newly launched Excel visible
	if _, err := oleutil.PutProperty(excel, "Visible", true); err != nil {
		excel.Release()
		ole.CoUninitialize()
		runtime.UnlockOSThread()
		return nil, fmt.Errorf("failed to set Visible: %w", err)
	}
	return &OleExcelApp{application: excel}, nil
}

func (a *OleExcelApp) Release() {
	if a.application != nil {
		a.application.Release()
	}
	ole.CoUninitialize()
	runtime.UnlockOSThread()
}

func (a *OleExcelApp) ListWorkbooks() ([]WorkbookInfo, error) {
	workbooksProp, err := oleutil.GetProperty(a.application, "Workbooks")
	if err != nil {
		return nil, fmt.Errorf("failed to get Workbooks: %w", err)
	}
	workbooks := workbooksProp.ToIDispatch()
	defer workbooks.Release()

	countProp, err := oleutil.GetProperty(workbooks, "Count")
	if err != nil {
		return nil, fmt.Errorf("failed to get Workbooks.Count: %w", err)
	}
	count := int(countProp.Val)

	var result []WorkbookInfo
	for i := 1; i <= count; i++ {
		wbProp, err := oleutil.GetProperty(workbooks, "Item", i)
		if err != nil {
			continue
		}
		wb := wbProp.ToIDispatch()

		nameProp, err := oleutil.GetProperty(wb, "Name")
		if err != nil {
			wb.Release()
			continue
		}
		fullNameProp, err := oleutil.GetProperty(wb, "FullName")
		if err != nil {
			wb.Release()
			continue
		}
		savedProp, err := oleutil.GetProperty(wb, "Saved")
		if err != nil {
			wb.Release()
			continue
		}

		result = append(result, WorkbookInfo{
			Name:     nameProp.ToString(),
			FullPath: fullNameProp.ToString(),
			Saved:    savedProp.Val != 0,
		})
		wb.Release()
	}
	return result, nil
}

func (a *OleExcelApp) OpenWorkbook(absolutePath string) error {
	workbooksProp, err := oleutil.GetProperty(a.application, "Workbooks")
	if err != nil {
		return fmt.Errorf("failed to get Workbooks: %w", err)
	}
	workbooks := workbooksProp.ToIDispatch()
	defer workbooks.Release()

	_, err = oleutil.CallMethod(workbooks, "Open", absolutePath)
	if err != nil {
		return fmt.Errorf("failed to open workbook: %w", err)
	}

	// Make Excel visible
	if _, err := oleutil.PutProperty(a.application, "Visible", true); err != nil {
		return fmt.Errorf("failed to set Visible: %w", err)
	}

	return nil
}

func (a *OleExcelApp) CreateWorkbook(absolutePath string) error {
	workbooksProp, err := oleutil.GetProperty(a.application, "Workbooks")
	if err != nil {
		return fmt.Errorf("failed to get Workbooks: %w", err)
	}
	workbooks := workbooksProp.ToIDispatch()
	defer workbooks.Release()

	wbResult, err := oleutil.CallMethod(workbooks, "Add")
	if err != nil {
		return fmt.Errorf("failed to create workbook: %w", err)
	}
	wb := wbResult.ToIDispatch()
	defer wb.Release()

	// SaveAs to the specified path
	_, err = oleutil.CallMethod(wb, "SaveAs", absolutePath)
	if err != nil {
		return fmt.Errorf("failed to save workbook: %w", err)
	}

	// Make Excel visible
	if _, err := oleutil.PutProperty(a.application, "Visible", true); err != nil {
		return fmt.Errorf("failed to set Visible: %w", err)
	}

	return nil
}

func (a *OleExcelApp) RunMacro(macroName string, args []string) (string, error) {
	// Build arguments for Application.Run
	// Application.Run(macroName, arg1, arg2, ..., arg10)
	callArgs := make([]interface{}, 0, 1+len(args))
	callArgs = append(callArgs, macroName)
	for _, arg := range args {
		callArgs = append(callArgs, arg)
	}

	result, err := oleutil.CallMethod(a.application, "Run", callArgs...)
	if err != nil {
		return "", fmt.Errorf("failed to run macro '%s': %w", macroName, err)
	}

	if result.Val == 0 {
		return "", nil
	}
	return fmt.Sprintf("%v", result.Value()), nil
}
