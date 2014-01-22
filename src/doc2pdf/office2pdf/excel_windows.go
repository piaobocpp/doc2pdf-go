// +build windows

// excel_windows
package office2pdf

import (
    "github.com/mattn/go-ole"
    "github.com/mattn/go-ole/oleutil"
    "path/filepath"
)

type Excel struct {
    app       *ole.IDispatch
    workbooks *ole.VARIANT
    xls       *ole.VARIANT
}

func (el *Excel) open(inFile string) (err error) {

    ole.CoInitialize(0)

    var unknown *ole.IUnknown

    unknown, err = oleutil.CreateObject("Excel.Application")
    if err != nil {
        return
    }

    el.app, err = unknown.QueryInterface(ole.IID_IDispatch)
    if err != nil {
        return
    }

    _, err = oleutil.PutProperty(el.app, "Visible", false)
    if err != nil {
        return
    }

    _, err = oleutil.PutProperty(el.app, "DisplayAlerts", false)
    if err != nil {
        return
    }

    el.workbooks, err = oleutil.GetProperty(el.app, "Workbooks")
    if err != nil {
        return
    }

    el.xls, err = oleutil.CallMethod(el.workbooks.ToIDispatch(), "Open", inFile)
    if err != nil {
        return
    }

    return
}

func (el *Excel) Export(inFile, outDir string) (outFile string, err error) {

    outFile = filepath.Join(outDir, filepath.Base(inFile+".pdf"))

    defer func() {
        if err != nil {
            outFile = ""
        }
        el.close()
    }()

    err = el.open(inFile)
    if err != nil {
        return
    }

    _, err = oleutil.CallMethod(el.xls.ToIDispatch(), "ExportAsFixedFormat", 0, outFile)
    if err != nil {
        return
    }

    return
}

func (el *Excel) close() {

    if el.xls != nil {
        oleutil.MustPutProperty(el.xls.ToIDispatch(), "Saved", true)
    }

    if el.workbooks != nil {
        oleutil.MustCallMethod(el.workbooks.ToIDispatch(), "Close")
    }

    if el.app != nil {
        oleutil.MustCallMethod(el.app, "Quit")
        el.app.Release()
    }

    ole.CoUninitialize()
}
