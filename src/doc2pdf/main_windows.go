// +build windows

// main_windows
package main

import (
    "doc2pdf/office2pdf"
    "fmt"
    "log"
    "net/http"
    "os"
    "path/filepath"
    "strings"
)

func fileIsExist(path string) bool {
    if _, err := os.Stat(path); os.IsNotExist(err) {
        return false
    }
    return true
}

func exporterMap() (m map[string]interface{}) {
    m = map[string]interface{}{
        ".doc":  new(office2pdf.Word),
        ".docx": new(office2pdf.Word),
        ".xls":  new(office2pdf.Excel),
        ".xlsx": new(office2pdf.Excel),
        ".ppt":  new(office2pdf.PowerPoint),
        ".pptx": new(office2pdf.PowerPoint),
    }
    return
}

func export(w http.ResponseWriter, r *http.Request) {

    r.ParseForm()

    inFile, outDir := strings.Join(r.Form["infile"], ""), strings.Join(r.Form["outdir"], "")

    log.Println("input file: " + inFile + "\noutput dir: " + outDir)

    if fileIsExist(inFile) && fileIsExist(outDir) {
        exporter := exporterMap()[filepath.Ext(inFile)]
        if _, ok := exporter.(office2pdf.Exporter); ok {
            outFile, err := exporter.(office2pdf.Exporter).Export(inFile, outDir)
            if err != nil {
                log.Fatal(err)
            }
            log.Println("output file: " + outFile)
            fmt.Fprintf(w, "%v", outFile)
        }
    }
}

func main() {
    port := "9000"
    if len(os.Args) > 1 {
        port = os.Args[1]
    }
    http.HandleFunc("/", export)
    log.Println("Server is listening on port " + port)
    err := http.ListenAndServe(":"+port, nil)
    if err != nil {
        log.Fatal("ListenAndServe: ", err)
    }
}
