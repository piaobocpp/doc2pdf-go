// +build windows

// interface_windows
package office2pdf

type Exporter interface {
    Export(inFile, outDir string) (outFile string, err error)
}
