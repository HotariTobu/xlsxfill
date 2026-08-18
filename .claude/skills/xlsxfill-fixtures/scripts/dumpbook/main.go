// dumpbook prints every cell's value and type, formulas, column widths,
// hyperlinks, and picture counts for each sheet of an xlsx file.
//
// Usage: go run . <book.xlsx>
package main

import (
	"fmt"
	_ "image/png"
	"os"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile(os.Args[1])
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	defer f.Close()
	for _, sheet := range f.GetSheetList() {
		fmt.Printf("== %s ==\n", sheet)
		cols, err := f.GetCols(sheet)
		if err != nil {
			// e.g. sheet names over 31 characters; verify those via workbook.xml
			fmt.Println("  (skip:", err, ")")
			continue
		}
		ncols := len(cols)
		rows, err := f.GetRows(sheet)
		if err != nil {
			fmt.Println("  (skip:", err, ")")
			continue
		}
		for ri, row := range rows {
			for ci, v := range row {
				if v == "" {
					continue
				}
				ref, _ := excelize.CoordinatesToCellName(ci+1, ri+1)
				typ, _ := f.GetCellType(sheet, ref)
				fmt.Printf("  %s [%v] %q\n", ref, typ, v)
				if ok, target, err := f.GetCellHyperLink(sheet, ref); err == nil && ok {
					fmt.Printf("    link -> %s\n", target)
				}
			}
		}
		for ri := 1; ri <= len(rows)+1; ri++ {
			for ci := 1; ci <= ncols+2; ci++ {
				ref, _ := excelize.CoordinatesToCellName(ci, ri)
				if fx, err := f.GetCellFormula(sheet, ref); err == nil && fx != "" {
					fmt.Printf("  %s formula =%s\n", ref, fx)
				}
				if pics, err := f.GetPictures(sheet, ref); err == nil && len(pics) > 0 {
					fmt.Printf("  %s pictures = %d\n", ref, len(pics))
				}
			}
		}
		for c := 1; c <= ncols+1; c++ {
			name, _ := excelize.ColumnNumberToName(c)
			w, err := f.GetColWidth(sheet, name)
			if err != nil {
				continue
			}
			fmt.Printf("  width %s = %v\n", name, w)
		}
	}
}
