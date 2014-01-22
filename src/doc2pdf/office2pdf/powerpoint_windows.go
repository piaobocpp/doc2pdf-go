// +build windows

// powerpoint_windows
package office2pdf

import (
    "github.com/mattn/go-ole"
    "github.com/mattn/go-ole/oleutil"
    "path/filepath"
)

type PowerPoint struct {
    app           *ole.IDispatch
    presentations *ole.VARIANT
    ppt           *ole.VARIANT
}

func (pt *PowerPoint) open(inFile string) (err error) {

    ole.CoInitialize(0)

    var unknown *ole.IUnknown

    unknown, err = oleutil.CreateObject("PowerPoint.Application")
    if err != nil {
        return
    }

    pt.app, err = unknown.QueryInterface(ole.IID_IDispatch)
    if err != nil {
        return
    }

    _, err = oleutil.PutProperty(pt.app, "DisplayAlerts", 1)
    if err != nil {
        return
    }

    pt.presentations, err = oleutil.GetProperty(pt.app, "Presentations")
    if err != nil {
        return
    }

    pt.ppt, err = oleutil.CallMethod(pt.presentations.ToIDispatch(), "Open", inFile, -1, 0, 0)
    if err != nil {
        return
    }

    return
}

func (pt *PowerPoint) Export(inFile, outDir string) (outFile string, err error) {

    outFile = filepath.Join(outDir, filepath.Base(inFile+".pdf"))

    defer func() {
        if err != nil {
            outFile = ""
        }
        pt.close()
    }()

    err = pt.open(inFile)
    if err != nil {
        return
    }

    _, err = oleutil.CallMethod(pt.ppt.ToIDispatch(), "SaveAs", outFile, 32)
    if err != nil {
        return
    }

    return
}

func (pt *PowerPoint) close() {

    if pt.ppt != nil {
        oleutil.MustPutProperty(pt.ppt.ToIDispatch(), "Saved", -1)
        oleutil.MustCallMethod(pt.ppt.ToIDispatch(), "Close")
    }

    if pt.app != nil {
        oleutil.MustCallMethod(pt.app, "Quit")
        pt.app.Release()
    }

    ole.CoUninitialize()
}
