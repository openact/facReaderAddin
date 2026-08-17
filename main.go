// facReaderAddin is a command-line helper for one-cell lookup from .fac files.
//
// Usage example:
//
//	facReaderAddin -file testData/rbc2512_4/E_1/PHKL_0.fac -period 202412 -sp 0 -type PHKL_0 -var DISC_A_PC
package main

import (
	"flag"
	"fmt"
	"os"

	"github.com/openact/facReaderAddin/internal/lookup"
	"github.com/openact/kit/cache/v3"
)

func main() {
	file := flag.String("file", "", "path to .fac file")
	period := flag.String("period", "", "time-period column, e.g. 202512")
	spFlag := flag.String("sp", "", "SP_CODE exact value")
	typeFlag := flag.String("type", "", "reserved (kept for compatibility; not used)")
	varFlag := flag.String("var", "", "VAR_NAME exact value")
	simFlag := flag.String("sim", "0", "SIM_ID value (optional, default 0)")
	flag.Parse()

	if *file == "" || *period == "" {
		fmt.Fprintln(os.Stderr, "error: -file and -period are required")
		flag.Usage()
		os.Exit(1)
	}

	tbl := cache.LoadTable(*file)
	product := lookup.ProductFromFilePath(*file)
	v, ok := lookup.EProj_Result(tbl, product, *spFlag, *typeFlag, *varFlag, *period, *simFlag)
	if !ok {
		fmt.Fprintln(os.Stderr, "not found")
		os.Exit(1)
	}
	fmt.Printf("%.6f\n", v)
}
