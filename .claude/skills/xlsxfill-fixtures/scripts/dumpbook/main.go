package main

import (
	"archive/zip"
	"crypto/sha256"
	"encoding/json"
	"encoding/xml"
	"fmt"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"os"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	if len(os.Args) != 2 {
		fmt.Fprintln(os.Stderr, "usage: dumpbook <book.xlsx>")
		os.Exit(2)
	}
	path := os.Args[1]

	pkg, err := readPackage(path)
	if err != nil {
		fmt.Fprintln(os.Stderr, "cannot read package:", err)
		os.Exit(1)
	}

	var out []string
	emit := func(format string, a ...any) {
		out = append(out, fmt.Sprintf(format, a...))
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		emit("book\topen-failed\t%v", err)
		emit("book\tfallback\twhole package read as XML")
		dumpRaw(pkg, emit, nil)
		dumpParts(pkg, emit)
		dumpModelled(pkg, emit)
		print(out)
		return
	}
	defer f.Close()

	sheets := f.GetSheetList()
	for i, name := range sheets {
		emit("sheet[%d]\t%s", i, name)
	}

	names := f.GetDefinedName()
	sort.Slice(names, func(i, j int) bool {
		if names[i].Name != names[j].Name {
			return names[i].Name < names[j].Name
		}
		return names[i].Scope < names[j].Scope
	})
	for _, n := range names {
		emit("definedName\t%s\tscope=%s\t%s", n.Name, n.Scope, n.RefersTo)
	}

	refs := cellRefs(pkg, emit)
	unreachable := map[string]bool{}
	for _, sheet := range sheets {
		if !dumpSheet(f, sheet, refs[sheet], emit) {
			unreachable[sheet] = true
		}
	}
	if len(unreachable) > 0 {
		emit("book\tfallback\t%d sheet(s) read as XML instead", len(unreachable))
	}
	dumpRaw(pkg, emit, unreachable)
	dumpParts(pkg, emit)
	dumpModelled(pkg, emit)

	print(out)
}

func print(out []string) {
	sort.Strings(out)
	for _, line := range out {
		fmt.Println(line)
	}
}

func jsonOf(v any) string {
	b, err := json.Marshal(v)
	if err != nil {
		return fmt.Sprintf("<unmarshalable: %v>", err)
	}
	return string(b)
}

func dumpSheet(f *excelize.File, sheet string, refs []string, emit func(string, ...any)) bool {
	if _, err := f.GetRows(sheet); err != nil {
		emit("sheet-unreachable\t%s\t%v", sheet, err)
		return false
	}

	for _, ref := range refs {
		dumpCell(f, sheet, ref, emit)
	}

	report := func(what string, err error) bool {
		if err != nil {
			emit("%s-error\t%s\t%v", what, sheet, err)
			return false
		}
		return true
	}

	if merges, err := f.GetMergeCells(sheet); report("merge", err) {
		var lines []string
		for _, m := range merges {
			lines = append(lines, m.GetStartAxis()+":"+m.GetEndAxis())
		}
		sort.Strings(lines)
		for _, l := range lines {
			emit("merge\t%s\t%s", sheet, l)
		}
	}

	if dvs, err := f.GetDataValidations(sheet); report("validation", err) {
		var lines []string
		for _, d := range dvs {
			lines = append(lines, jsonOf(d))
		}
		sort.Strings(lines)
		for _, l := range lines {
			emit("validation\t%s\t%s", sheet, l)
		}
	}

	if cfs, err := f.GetConditionalFormats(sheet); report("condfmt", err) {
		var lines []string
		for ref, opts := range cfs {
			for _, o := range opts {
				lines = append(lines, ref+"\t"+jsonOf(o))
			}
		}
		sort.Strings(lines)
		for _, l := range lines {
			emit("condfmt\t%s\t%s", sheet, l)
		}
	}

	if tables, err := f.GetTables(sheet); report("table", err) {
		var lines []string
		for _, t := range tables {
			lines = append(lines, jsonOf(t))
		}
		sort.Strings(lines)
		for _, l := range lines {
			emit("table\t%s\t%s", sheet, l)
		}
	}

	if cells, err := f.GetPictureCells(sheet); report("picture", err) {
		var lines []string
		for _, ref := range cells {
			pics, err := f.GetPictures(sheet, ref)
			if err != nil {
				emit("picture-error\t%s\t%s\t%v", sheet, ref, err)
				continue
			}
			for _, p := range pics {
				sum := sha256.Sum256(p.File)
				lines = append(lines, fmt.Sprintf("%s\text=%s\tsha256=%x\tbytes=%d\t%s",
					ref, p.Extension, sum, len(p.File), jsonOf(p.Format)))
			}
		}
		sort.Strings(lines)
		for _, l := range lines {
			emit("picture\t%s\t%s", sheet, l)
		}
	}

	if comments, err := f.GetComments(sheet); report("comment", err) {
		var lines []string
		for _, c := range comments {
			lines = append(lines, jsonOf(c))
		}
		sort.Strings(lines)
		for _, l := range lines {
			emit("comment\t%s\t%s", sheet, l)
		}
	}

	if visible, err := f.GetSheetVisible(sheet); report("visible", err) {
		emit("visible\t%s\t%v", sheet, visible)
	}
	if view, err := f.GetSheetView(sheet, 0); report("view", err) {
		emit("view\t%s\t%s", sheet, jsonOf(view))
	}
	if props, err := f.GetSheetProps(sheet); report("sheetprops", err) {
		emit("sheetprops\t%s\t%s", sheet, jsonOf(props))
	}
	if layout, err := f.GetPageLayout(sheet); report("pagelayout", err) {
		emit("pagelayout\t%s\t%s", sheet, jsonOf(layout))
	}
	if margins, err := f.GetPageMargins(sheet); report("pagemargins", err) {
		emit("pagemargins\t%s\t%s", sheet, jsonOf(margins))
	}

	return true
}

func dumpCell(f *excelize.File, sheet, ref string, emit func(string, ...any)) {
	raw, err := f.GetCellValue(sheet, ref, excelize.Options{RawCellValue: true})
	if err != nil {
		emit("cell-error\t%s!%s\traw\t%v", sheet, ref, err)
		return
	}
	value, err := f.GetCellValue(sheet, ref)
	if err != nil {
		emit("cell-error\t%s!%s\tvalue\t%v", sheet, ref, err)
		return
	}
	kind, err := f.GetCellType(sheet, ref)
	if err != nil {
		emit("cell-error\t%s!%s\ttype\t%v", sheet, ref, err)
		return
	}
	formula, err := f.GetCellFormula(sheet, ref)
	if err != nil {
		emit("cell-error\t%s!%s\tformula\t%v", sheet, ref, err)
		return
	}
	styleID, err := f.GetCellStyle(sheet, ref)
	if err != nil {
		emit("cell-error\t%s!%s\tstyle\t%v", sheet, ref, err)
		return
	}
	if raw == "" && value == "" && formula == "" && styleID == 0 {
		return
	}
	style := "default"
	if styleID != 0 {
		s, err := f.GetStyle(styleID)
		if err != nil {
			style = fmt.Sprintf("<error: %v>", err)
		} else {
			style = jsonOf(s)
		}
	}
	emit("cell\t%s!%s\ttype=%d\traw=%q\tvalue=%q\tformula=%q\tstyle=%s",
		sheet, ref, kind, raw, value, formula, style)

	linked, target, err := f.GetCellHyperLink(sheet, ref)
	if err != nil {
		emit("link-error\t%s!%s\t%v", sheet, ref, err)
		return
	}
	if linked {
		emit("link\t%s\t%s -> %s", sheet, ref, target)
	}
}

var nsPrefix = map[string]string{
	"http://schemas.openxmlformats.org/spreadsheetml/2006/main":                        "s",
	"http://schemas.openxmlformats.org/drawingml/2006/main":                            "a",
	"http://schemas.openxmlformats.org/drawingml/2006/chart":                           "c",
	"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing":                    "cdr",
	"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing":              "xdr",
	"http://schemas.openxmlformats.org/package/2006/relationships":                     "pr",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships":              "r",
	"http://schemas.openxmlformats.org/package/2006/content-types":                     "ct",
	"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties":        "ep",
	"http://schemas.openxmlformats.org/package/2006/metadata/core-properties":          "cp",
	"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes":             "vt",
	"http://schemas.openxmlformats.org/markup-compatibility/2006":                      "mc",
	"http://purl.org/dc/elements/1.1/":                                                 "dc",
	"http://purl.org/dc/terms/":                                                        "dcterms",
	"http://www.w3.org/2001/XMLSchema-instance":                                        "xsi",
	"http://www.w3.org/XML/1998/namespace":                                             "xml",
	"urn:schemas-microsoft-com:vml":                                                    "v",
	"urn:schemas-microsoft-com:office:office":                                          "o",
	"urn:schemas-microsoft-com:office:excel":                                           "x",
	"http://schemas.microsoft.com/office/drawing/2014/main":                            "a14",
	"http://schemas.openxmlformats.org/officeDocument/2006/customXml":                  "cx",
	"http://schemas.openxmlformats.org/spreadsheetml/2006/main/x14ac":                  "x14ac",
	"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac":                      "x14ac",
	"http://schemas.openxmlformats.org/officeDocument/2006/math":                       "m",
	"http://schemas.microsoft.com/office/drawing/2010/main":                            "a10",
	"http://schemas.openxmlformats.org/officeDocument/2006/bibliography":               "b",
	"http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas":                    "lc",
	"http://schemas.openxmlformats.org/drawingml/2006/picture":                         "pic",
	"http://schemas.openxmlformats.org/drawingml/2006/compatibility":                   "comp",
	"http://schemas.openxmlformats.org/officeDocument/2006/customProperties":           "custom",
	"http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes":                "st",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":    "hl",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":        "img",
	"http://schemas.openxmlformats.org/spreadsheetml/2006/main/revision":               "rev",
	"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main":                   "x14",
	"http://schemas.microsoft.com/office/spreadsheetml/2014/revision":                  "xr",
	"http://schemas.microsoft.com/office/spreadsheetml/2016/revision":                  "xr16",
	"http://schemas.microsoft.com/office/spreadsheetml/2017/revision":                  "xr17",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject":    "ole",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart":        "chartrel",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing":   "vmlrel",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table":        "tablerel",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments":     "cmtrel",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing":      "drawrel",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":    "wsrel",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedString": "sstrel",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":       "stylerel",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":        "themerel",
}

func short(space, local string) string {
	if space == "" {
		return local
	}
	if p, ok := nsPrefix[space]; ok {
		return p + ":" + local
	}
	return "{" + space + "}" + local
}

func modelled(name string) bool {
	switch name {
	case "xl/workbook.xml", "xl/sharedStrings.xml":
		return true
	}
	switch {
	case strings.HasPrefix(name, "xl/worksheets/") && strings.HasSuffix(name, ".xml"):
		return true
	case strings.HasPrefix(name, "xl/tables/"):
		return true
	case strings.HasPrefix(name, "xl/comments"):
		return true
	}
	return false
}

func flattenPart(part string, blob []byte, emit func(string, ...any)) {
	dec := xml.NewDecoder(strings.NewReader(string(blob)))
	var path []string
	counts := map[string]int{}
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return
		}
		if err != nil {
			emit("part-error\t%s\t%v", part, err)
			return
		}
		switch t := tok.(type) {
		case xml.StartElement:
			name := short(t.Name.Space, t.Name.Local)
			key := strings.Join(path, "/") + "/" + name
			i := counts[key]
			counts[key] = i + 1
			path = append(path, fmt.Sprintf("%s[%d]", name, i))
			var attrs []string
			for _, a := range t.Attr {
				attrs = append(attrs, fmt.Sprintf("%s=%q", short(a.Name.Space, a.Name.Local), a.Value))
			}
			sort.Strings(attrs)
			emit("part\t%s\t%s\t%s", part, strings.Join(path, "/"), strings.Join(attrs, " "))
		case xml.EndElement:
			if len(path) > 0 {
				path = path[:len(path)-1]
			}
		case xml.CharData:
			if strings.TrimSpace(string(t)) == "" {
				continue
			}
			emit("part\t%s\t%s\ttext=%q", part, strings.Join(path, "/"), string(t))
		}
	}
}

func dumpParts(pkg *pkgFile, emit func(string, ...any)) {
	for _, name := range pkg.order {
		if modelled(name) {
			continue
		}
		blob := pkg.parts[name]
		lower := strings.ToLower(name)
		if strings.HasSuffix(lower, ".xml") || strings.HasSuffix(lower, ".rels") ||
			strings.HasSuffix(lower, ".vml") {
			flattenPart(name, blob, emit)
			continue
		}
		sum := sha256.Sum256(blob)
		emit("blob\t%s\tsha256=%x\tbytes=%d", name, sum, len(blob))
	}
}

const (
	nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	nsRel  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)

type pkgFile struct {
	parts map[string][]byte
	order []string
}

func readPackage(path string) (*pkgFile, error) {
	r, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer r.Close()
	pkg := &pkgFile{parts: map[string][]byte{}}
	for _, file := range r.File {
		if strings.HasSuffix(file.Name, "/") {
			continue
		}
		rc, err := file.Open()
		if err != nil {
			return nil, err
		}
		blob, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return nil, err
		}
		pkg.parts[file.Name] = blob
		pkg.order = append(pkg.order, file.Name)
	}
	return pkg, nil
}

type xlWorkbook struct {
	Ignorable   string `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 Ignorable,attr"`
	FileVersion struct {
		AppName      string `xml:"appName,attr"`
		LastEdited   string `xml:"lastEdited,attr"`
		LowestEdited string `xml:"lowestEdited,attr"`
		RupBuild     string `xml:"rupBuild,attr"`
	} `xml:"fileVersion"`
	WorkbookPr struct {
		Date1904            string `xml:"date1904,attr"`
		DefaultThemeVersion string `xml:"defaultThemeVersion,attr"`
		FilterPrivacy       string `xml:"filterPrivacy,attr"`
		BackupFile          string `xml:"backupFile,attr"`
		ShowObjects         string `xml:"showObjects,attr"`
	} `xml:"workbookPr"`
	WorkbookProtection struct {
		LockStructure string `xml:"lockStructure,attr"`
		LockWindows   string `xml:"lockWindows,attr"`
	} `xml:"workbookProtection"`
	BookViews struct {
		WorkbookView []struct {
			WindowHeight string `xml:"windowHeight,attr"`
			WindowWidth  string `xml:"windowWidth,attr"`
			XWindow      string `xml:"xWindow,attr"`
			YWindow      string `xml:"yWindow,attr"`
			ActiveTab    string `xml:"activeTab,attr"`
			Visibility   string `xml:"visibility,attr"`
		} `xml:"workbookView"`
	} `xml:"bookViews"`
	Sheets struct {
		Sheet []struct {
			Name    string `xml:"name,attr"`
			ID      string `xml:"id,attr"`
			SheetID string `xml:"sheetId,attr"`
			State   string `xml:"state,attr"`
		} `xml:"sheet"`
	} `xml:"sheets"`
	DefinedNames struct {
		DefinedName []struct {
			Name         string `xml:"name,attr"`
			LocalSheetID string `xml:"localSheetId,attr"`
			Hidden       string `xml:"hidden,attr"`
			Value        string `xml:",chardata"`
		} `xml:"definedName"`
	} `xml:"definedNames"`
	CalcPr struct {
		CalcID         string `xml:"calcId,attr"`
		CalcMode       string `xml:"calcMode,attr"`
		FullCalcOnLoad string `xml:"fullCalcOnLoad,attr"`
	} `xml:"calcPr"`
}

type xlTable struct {
	ID          string `xml:"id,attr"`
	Name        string `xml:"name,attr"`
	DisplayName string `xml:"displayName,attr"`
	Ref         string `xml:"ref,attr"`
	HeaderRow   string `xml:"headerRowCount,attr"`
	TotalsRow   string `xml:"totalsRowCount,attr"`
	AutoFilter  struct {
		Ref string `xml:"ref,attr"`
	} `xml:"autoFilter"`
	TableColumns struct {
		Count  string `xml:"count,attr"`
		Column []struct {
			ID   string `xml:"id,attr"`
			Name string `xml:"name,attr"`
		} `xml:"tableColumn"`
	} `xml:"tableColumns"`
	TableStyleInfo struct {
		Name              string `xml:"name,attr"`
		ShowColumnStripes string `xml:"showColumnStripes,attr"`
		ShowFirstColumn   string `xml:"showFirstColumn,attr"`
		ShowLastColumn    string `xml:"showLastColumn,attr"`
		ShowRowStripes    string `xml:"showRowStripes,attr"`
	} `xml:"tableStyleInfo"`
}

type xlComments struct {
	Authors struct {
		Author []string `xml:"author"`
	} `xml:"authors"`
	CommentList struct {
		Comment []struct {
			Ref      string   `xml:"ref,attr"`
			AuthorID string   `xml:"authorId,attr"`
			Text     []string `xml:"text>t"`
			RichText []string `xml:"text>r>t"`
		} `xml:"comment"`
	} `xml:"commentList"`
}

type xlRels struct {
	Rel []struct {
		ID     string `xml:"Id,attr"`
		Target string `xml:"Target,attr"`
	} `xml:"Relationship"`
}

type xlText struct {
	Space string `xml:"http://www.w3.org/XML/1998/namespace space,attr"`
	Text  string `xml:",chardata"`
}

type xlVal struct {
	Val string `xml:"val,attr"`
}

type xlRPr struct {
	B         *xlVal `xml:"b"`
	I         *xlVal `xml:"i"`
	U         *xlVal `xml:"u"`
	Strike    *xlVal `xml:"strike"`
	Sz        *xlVal `xml:"sz"`
	Family    *xlVal `xml:"family"`
	Charset   *xlVal `xml:"charset"`
	Scheme    *xlVal `xml:"scheme"`
	RFont     *xlVal `xml:"rFont"`
	VertAlign *xlVal `xml:"vertAlign"`
	Color     *struct {
		RGB     string `xml:"rgb,attr"`
		Theme   string `xml:"theme,attr"`
		Tint    string `xml:"tint,attr"`
		Indexed string `xml:"indexed,attr"`
		Auto    string `xml:"auto,attr"`
	} `xml:"color"`
}

type xlSST struct {
	Count       string `xml:"count,attr"`
	UniqueCount string `xml:"uniqueCount,attr"`
	SI          []struct {
		T []xlText `xml:"t"`
		R []struct {
			RPr *xlRPr `xml:"rPr"`
			T   xlText `xml:"t"`
		} `xml:"r"`
	} `xml:"si"`
}

type xlRelID struct {
	ID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

type xlSheet struct {
	Dimension struct {
		Ref string `xml:"ref,attr"`
	} `xml:"dimension"`
	SheetPr struct {
		FilterMode string `xml:"filterMode,attr"`
		TabColor   struct {
			RGB   string `xml:"rgb,attr"`
			Theme string `xml:"theme,attr"`
		} `xml:"tabColor"`
	} `xml:"sheetPr"`
	SheetViews struct {
		SheetView []struct {
			ShowGridLines  string `xml:"showGridLines,attr"`
			TabSelected    string `xml:"tabSelected,attr"`
			WorkbookViewID string `xml:"workbookViewId,attr"`
			ZoomScale      string `xml:"zoomScale,attr"`
			Pane           struct {
				ActivePane  string `xml:"activePane,attr"`
				State       string `xml:"state,attr"`
				TopLeftCell string `xml:"topLeftCell,attr"`
				XSplit      string `xml:"xSplit,attr"`
				YSplit      string `xml:"ySplit,attr"`
			} `xml:"pane"`
		} `xml:"sheetView"`
	} `xml:"sheetViews"`
	AutoFilter struct {
		Ref string `xml:"ref,attr"`
	} `xml:"autoFilter"`
	Hyperlinks struct {
		Hyperlink []struct {
			Ref      string `xml:"ref,attr"`
			ID       string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
			Tooltip  string `xml:"tooltip,attr"`
			Display  string `xml:"display,attr"`
			Location string `xml:"location,attr"`
		} `xml:"hyperlink"`
	} `xml:"hyperlinks"`
	SheetProtection struct {
		Sheet   string `xml:"sheet,attr"`
		Objects string `xml:"objects,attr"`
		Content string `xml:"content,attr"`
	} `xml:"sheetProtection"`
	PrintOptions struct {
		HorizontalCentered string `xml:"horizontalCentered,attr"`
		VerticalCentered   string `xml:"verticalCentered,attr"`
		GridLines          string `xml:"gridLines,attr"`
		Headings           string `xml:"headings,attr"`
	} `xml:"printOptions"`
	PageSetup struct {
		Orientation string `xml:"orientation,attr"`
		PaperSize   string `xml:"paperSize,attr"`
		Scale       string `xml:"scale,attr"`
		FitToWidth  string `xml:"fitToWidth,attr"`
		FitToHeight string `xml:"fitToHeight,attr"`
	} `xml:"pageSetup"`
	HeaderFooter struct {
		OddHeader  string `xml:"oddHeader"`
		OddFooter  string `xml:"oddFooter"`
		EvenHeader string `xml:"evenHeader"`
		EvenFooter string `xml:"evenFooter"`
	} `xml:"headerFooter"`
	Drawing       xlRelID `xml:"drawing"`
	LegacyDrawing xlRelID `xml:"legacyDrawing"`
	TableParts    struct {
		Count     string    `xml:"count,attr"`
		TablePart []xlRelID `xml:"tablePart"`
	} `xml:"tableParts"`
	SheetFormatPr struct {
		DefaultRowHeight string `xml:"defaultRowHeight,attr"`
		DefaultColWidth  string `xml:"defaultColWidth,attr"`
		BaseColWidth     string `xml:"baseColWidth,attr"`
		CustomHeight     string `xml:"customHeight,attr"`
	} `xml:"sheetFormatPr"`
	Cols struct {
		Col []struct {
			Min         string `xml:"min,attr"`
			Max         string `xml:"max,attr"`
			Width       string `xml:"width,attr"`
			CustomWidth string `xml:"customWidth,attr"`
			Hidden      string `xml:"hidden,attr"`
			Style       string `xml:"style,attr"`
		} `xml:"col"`
	} `xml:"cols"`
	SheetData struct {
		Row []struct {
			R            string `xml:"r,attr"`
			Ht           string `xml:"ht,attr"`
			CustomHeight string `xml:"customHeight,attr"`
			Hidden       string `xml:"hidden,attr"`
			S            string `xml:"s,attr"`
			CustomFormat string `xml:"customFormat,attr"`
			C            []struct {
				R  string `xml:"r,attr"`
				T  string `xml:"t,attr"`
				S  string `xml:"s,attr"`
				V  string `xml:"v"`
				F  string `xml:"f"`
				Is struct {
					T []string `xml:"t"`
				} `xml:"is"`
			} `xml:"c"`
		} `xml:"row"`
	} `xml:"sheetData"`
	MergeCells struct {
		MergeCell []struct {
			Ref string `xml:"ref,attr"`
		} `xml:"mergeCell"`
	} `xml:"mergeCells"`
}

func dumpRaw(pkg *pkgFile, emit func(string, ...any), only map[string]bool) {
	strings_ := rawStrings(pkg, emit)

	for i, sheet := range readSheets(pkg, emit) {
		if only == nil {
			emit("sheet[%d]\t%s", i, sheet.Name)
		}
		ws := sheet.WS
		dumpRawGeometry(sheet.Name, &ws, emit)
		dumpSheetPart(sheet.Name, &ws, emit)
		for _, row := range ws.SheetData.Row {
			for _, c := range row.C {
				kind, value := c.T, c.V
				switch c.T {
				case "s":
					n, err := strconv.Atoi(c.V)
					if err != nil || n < 0 || n >= len(strings_) {
						emit("raw-error\t%s!%s\tbad shared string index %q", sheet.Name, c.R, c.V)
						continue
					}
					kind, value = "str", strings_[n]
				case "inlineStr":
					kind, value = "str", strings.Join(c.Is.T, "")
				case "":
					kind = "n"
				}
				emit("raw-cell\t%s!%s\ttype=%s\tvalue=%q\tformula=%q\tstyle=%s",
					sheet.Name, c.R, kind, value, c.F, orDefault(c.S))
			}
		}
		var merges []string
		for _, m := range ws.MergeCells.MergeCell {
			merges = append(merges, m.Ref)
		}
		sort.Strings(merges)
		for _, m := range merges {
			emit("raw-merge\t%s\t%s", sheet.Name, m)
		}
	}
}

type rawSheet struct {
	Name string
	WS   xlSheet
}

func readSheets(pkg *pkgFile, emit func(string, ...any)) []rawSheet {
	raw, ok := pkg.parts["xl/workbook.xml"]
	if !ok {
		emit("raw-error\tno xl/workbook.xml")
		return nil
	}
	var wb xlWorkbook
	if err := xml.Unmarshal(raw, &wb); err != nil {
		emit("raw-error\tworkbook.xml\t%v", err)
		return nil
	}
	var rels xlRels
	if blob, ok := pkg.parts["xl/_rels/workbook.xml.rels"]; ok {
		if err := xml.Unmarshal(blob, &rels); err != nil {
			emit("raw-error\tworkbook.xml.rels\t%v", err)
			return nil
		}
	}
	target := map[string]string{}
	for _, r := range rels.Rel {
		target[r.ID] = r.Target
	}

	var out []rawSheet
	for _, sheet := range wb.Sheets.Sheet {
		part := target[sheet.ID]
		if part == "" {
			emit("raw-error\t%s\tno relationship for %s", sheet.Name, sheet.ID)
			continue
		}
		part = "xl/" + strings.TrimPrefix(strings.TrimPrefix(part, "/"), "xl/")
		blob, ok := pkg.parts[part]
		if !ok {
			emit("raw-error\t%s\tmissing part %s", sheet.Name, part)
			continue
		}
		var ws xlSheet
		if err := xml.Unmarshal(blob, &ws); err != nil {
			emit("raw-error\t%s\t%v", sheet.Name, err)
			continue
		}
		out = append(out, rawSheet{Name: sheet.Name, WS: ws})
	}
	return out
}

func cellRefs(pkg *pkgFile, emit func(string, ...any)) map[string][]string {
	out := map[string][]string{}
	for _, sheet := range readSheets(pkg, emit) {
		var refs []string
		for _, row := range sheet.WS.SheetData.Row {
			for _, c := range row.C {
				if c.R == "" {
					continue
				}
				refs = append(refs, c.R)
			}
		}
		out[sheet.Name] = refs
	}
	return out
}

func dumpSheetPart(sheet string, ws *xlSheet, emit func(string, ...any)) {
	emit("dimension\t%s\t%s", sheet, orDash(ws.Dimension.Ref))
	emit("sheetpr\t%s\tfilterMode=%s\ttabColor=%s/%s", sheet,
		orDash(ws.SheetPr.FilterMode), orDash(ws.SheetPr.TabColor.RGB),
		orDash(ws.SheetPr.TabColor.Theme))
	emit("autofilter\t%s\t%s", sheet, orDash(ws.AutoFilter.Ref))
	var links []string
	for _, h := range ws.Hyperlinks.Hyperlink {
		links = append(links, fmt.Sprintf("%s\trel=%s\ttooltip=%q\tdisplay=%q\tlocation=%q",
			h.Ref, orDash(h.ID), h.Tooltip, h.Display, h.Location))
	}
	sort.Strings(links)
	for _, l := range links {
		emit("hyperlink-raw\t%s\t%s", sheet, l)
	}
	emit("protection\t%s\tsheet=%s\tobjects=%s\tcontent=%s", sheet,
		orDash(ws.SheetProtection.Sheet), orDash(ws.SheetProtection.Objects),
		orDash(ws.SheetProtection.Content))
	emit("printoptions\t%s\thCentered=%s\tvCentered=%s\tgridLines=%s\theadings=%s", sheet,
		orDash(ws.PrintOptions.HorizontalCentered), orDash(ws.PrintOptions.VerticalCentered),
		orDash(ws.PrintOptions.GridLines), orDash(ws.PrintOptions.Headings))
	emit("pagesetup\t%s\torientation=%s\tpaperSize=%s\tscale=%s\tfitTo=%sx%s", sheet,
		orDash(ws.PageSetup.Orientation), orDash(ws.PageSetup.PaperSize),
		orDash(ws.PageSetup.Scale), orDash(ws.PageSetup.FitToWidth),
		orDash(ws.PageSetup.FitToHeight))
	emit("headerfooter\t%s\toddHeader=%q\toddFooter=%q\tevenHeader=%q\tevenFooter=%q", sheet,
		ws.HeaderFooter.OddHeader, ws.HeaderFooter.OddFooter,
		ws.HeaderFooter.EvenHeader, ws.HeaderFooter.EvenFooter)
	emit("sheetrels\t%s\tdrawing=%s\tlegacyDrawing=%s\ttableParts=%s", sheet,
		orDash(ws.Drawing.ID), orDash(ws.LegacyDrawing.ID), orDash(ws.TableParts.Count))
	for i, tp := range ws.TableParts.TablePart {
		emit("sheettablepart\t%s\t%03d\t%s", sheet, i, orDash(tp.ID))
	}
	for i, v := range ws.SheetViews.SheetView {
		emit("sheetview-raw\t%s\t%d\tgridLines=%s\ttabSelected=%s\tviewId=%s\tzoom=%s\tpane=%s/%s/%s/%s/%s",
			sheet, i, orDash(v.ShowGridLines), orDash(v.TabSelected),
			orDash(v.WorkbookViewID), orDash(v.ZoomScale), orDash(v.Pane.ActivePane),
			orDash(v.Pane.State), orDash(v.Pane.TopLeftCell), orDash(v.Pane.XSplit),
			orDash(v.Pane.YSplit))
	}
}

func dumpRawGeometry(sheet string, ws *xlSheet, emit func(string, ...any)) {
	fp := ws.SheetFormatPr
	if fp.DefaultRowHeight != "" || fp.DefaultColWidth != "" || fp.BaseColWidth != "" {
		emit("format\t%s\tdefaultRowHeight=%s\tdefaultColWidth=%s\tbaseColWidth=%s\tcustomHeight=%s",
			sheet, orDash(fp.DefaultRowHeight), orDash(fp.DefaultColWidth),
			orDash(fp.BaseColWidth), orDash(fp.CustomHeight))
	}
	var lines []string
	for _, c := range ws.Cols.Col {
		lines = append(lines, fmt.Sprintf("%s\tmin=%s\tmax=%s\twidth=%s\tcustomWidth=%s\thidden=%s\tstyle=%s",
			pad(c.Min), c.Min, c.Max, orDash(c.Width), orDash(c.CustomWidth),
			orDash(c.Hidden), orDash(c.Style)))
	}
	sort.Strings(lines)
	for _, l := range lines {
		emit("col\t%s\t%s", sheet, strings.SplitN(l, "\t", 2)[1])
	}
	lines = nil
	for _, r := range ws.SheetData.Row {
		if r.Ht == "" && r.CustomHeight == "" && r.Hidden == "" && r.CustomFormat == "" {
			continue
		}
		lines = append(lines, fmt.Sprintf("%s\tr=%s\tht=%s\tcustomHeight=%s\thidden=%s\tcustomFormat=%s\ts=%s",
			pad(r.R), r.R, orDash(r.Ht), orDash(r.CustomHeight), orDash(r.Hidden),
			orDash(r.CustomFormat), orDash(r.S)))
	}
	sort.Strings(lines)
	for _, l := range lines {
		emit("row\t%s\t%s", sheet, strings.SplitN(l, "\t", 2)[1])
	}
}

func pad(n string) string {
	v, err := strconv.Atoi(n)
	if err != nil {
		return n
	}
	return fmt.Sprintf("%08d", v)
}

func orDash(s string) string {
	if s == "" {
		return "-"
	}
	return s
}

func readSST(pkg *pkgFile, emit func(string, ...any)) (xlSST, bool) {
	raw, ok := pkg.parts["xl/sharedStrings.xml"]
	if !ok {
		return xlSST{}, false
	}
	var sst xlSST
	if err := xml.Unmarshal(raw, &sst); err != nil {
		emit("raw-error\tsharedStrings.xml\t%v", err)
		return xlSST{}, false
	}
	return sst, true
}

func rawStrings(pkg *pkgFile, emit func(string, ...any)) []string {
	sst, ok := readSST(pkg, emit)
	if !ok {
		return nil
	}
	out := make([]string, 0, len(sst.SI))
	for _, si := range sst.SI {
		var b strings.Builder
		for _, t := range si.T {
			b.WriteString(t.Text)
		}
		for _, r := range si.R {
			b.WriteString(r.T.Text)
		}
		out = append(out, b.String())
	}
	return out
}

func dumpModelled(pkg *pkgFile, emit func(string, ...any)) {
	dumpWorkbookPart(pkg, emit)
	dumpSSTPart(pkg, emit)
	dumpTableParts(pkg, emit)
	dumpCommentParts(pkg, emit)
}

func dumpWorkbookPart(pkg *pkgFile, emit func(string, ...any)) {
	raw, ok := pkg.parts["xl/workbook.xml"]
	if !ok {
		emit("raw-error\tno xl/workbook.xml")
		return
	}
	var wb xlWorkbook
	if err := xml.Unmarshal(raw, &wb); err != nil {
		emit("raw-error\tworkbook.xml\t%v", err)
		return
	}
	emit("workbookroot\tIgnorable=%s", orDash(wb.Ignorable))
	emit("workbookpr\tdate1904=%s\tdefaultThemeVersion=%s\tfilterPrivacy=%s\tbackupFile=%s\tshowObjects=%s",
		orDash(wb.WorkbookPr.Date1904), orDash(wb.WorkbookPr.DefaultThemeVersion),
		orDash(wb.WorkbookPr.FilterPrivacy), orDash(wb.WorkbookPr.BackupFile),
		orDash(wb.WorkbookPr.ShowObjects))
	emit("fileversion\tappName=%s\tlastEdited=%s\tlowestEdited=%s\trupBuild=%s",
		orDash(wb.FileVersion.AppName), orDash(wb.FileVersion.LastEdited),
		orDash(wb.FileVersion.LowestEdited), orDash(wb.FileVersion.RupBuild))
	emit("calcpr\tcalcId=%s\tcalcMode=%s\tfullCalcOnLoad=%s",
		orDash(wb.CalcPr.CalcID), orDash(wb.CalcPr.CalcMode), orDash(wb.CalcPr.FullCalcOnLoad))
	emit("workbookprotection\tlockStructure=%s\tlockWindows=%s",
		orDash(wb.WorkbookProtection.LockStructure), orDash(wb.WorkbookProtection.LockWindows))
	for i, v := range wb.BookViews.WorkbookView {
		emit("workbookview\t%d\twindow=%sx%s at %s,%s\tactiveTab=%s\tvisibility=%s",
			i, orDash(v.WindowWidth), orDash(v.WindowHeight), orDash(v.XWindow),
			orDash(v.YWindow), orDash(v.ActiveTab), orDash(v.Visibility))
	}
	for i, s := range wb.Sheets.Sheet {
		emit("sheetentry\t%03d\tname=%q\tsheetId=%s\tstate=%s\trel=%s",
			i, s.Name, orDash(s.SheetID), orDash(s.State), orDash(s.ID))
	}
	var lines []string
	for _, n := range wb.DefinedNames.DefinedName {
		lines = append(lines, fmt.Sprintf("%s\tlocalSheetId=%s\thidden=%s\t%s",
			n.Name, orDash(n.LocalSheetID), orDash(n.Hidden), n.Value))
	}
	sort.Strings(lines)
	for _, l := range lines {
		emit("definedname-raw\t%s", l)
	}
}

func dumpSSTPart(pkg *pkgFile, emit func(string, ...any)) {
	sst, ok := readSST(pkg, emit)
	if !ok {
		return
	}
	emit("sst\tcount=%s\tuniqueCount=%s\tentries=%d",
		orDash(sst.Count), orDash(sst.UniqueCount), len(sst.SI))
	for i, si := range sst.SI {
		var parts []string
		for _, t := range si.T {
			parts = append(parts, fmt.Sprintf("%q space=%s", t.Text, orDash(t.Space)))
		}
		for _, r := range si.R {
			format := "-"
			if r.RPr != nil {
				format = jsonOf(r.RPr)
			}
			parts = append(parts, fmt.Sprintf("run %q space=%s rPr=%s",
				r.T.Text, orDash(r.T.Space), format))
		}
		emit("si\t%05d\t%s", i, strings.Join(parts, " + "))
	}
}

func dumpTableParts(pkg *pkgFile, emit func(string, ...any)) {
	for _, name := range pkg.order {
		if !strings.HasPrefix(name, "xl/tables/") || !strings.HasSuffix(name, ".xml") {
			continue
		}
		var t xlTable
		if err := xml.Unmarshal(pkg.parts[name], &t); err != nil {
			emit("raw-error\t%s\t%v", name, err)
			continue
		}
		emit("tablepart\t%s\tid=%s\tname=%s\tdisplayName=%s\tref=%s\theaderRows=%s\ttotalsRows=%s\tautoFilter=%s",
			name, orDash(t.ID), orDash(t.Name), orDash(t.DisplayName), orDash(t.Ref),
			orDash(t.HeaderRow), orDash(t.TotalsRow), orDash(t.AutoFilter.Ref))
		emit("tablecolumns\t%s\tcount=%s\tactual=%d",
			name, orDash(t.TableColumns.Count), len(t.TableColumns.Column))
		for i, c := range t.TableColumns.Column {
			emit("tablecolumn\t%s\t%03d\tid=%s\tname=%q", name, i, orDash(c.ID), c.Name)
		}
		emit("tablestyle\t%s\tname=%s\tcolStripes=%s\tfirstCol=%s\tlastCol=%s\trowStripes=%s",
			name, orDash(t.TableStyleInfo.Name), orDash(t.TableStyleInfo.ShowColumnStripes),
			orDash(t.TableStyleInfo.ShowFirstColumn), orDash(t.TableStyleInfo.ShowLastColumn),
			orDash(t.TableStyleInfo.ShowRowStripes))
	}
}

func dumpCommentParts(pkg *pkgFile, emit func(string, ...any)) {
	for _, name := range pkg.order {
		if !strings.HasPrefix(name, "xl/comments") || !strings.HasSuffix(name, ".xml") {
			continue
		}
		var c xlComments
		if err := xml.Unmarshal(pkg.parts[name], &c); err != nil {
			emit("raw-error\t%s\t%v", name, err)
			continue
		}
		for i, a := range c.Authors.Author {
			emit("commentauthor\t%s\t%03d\t%q", name, i, a)
		}
		var lines []string
		for _, cm := range c.CommentList.Comment {
			text := strings.Join(cm.Text, "") + strings.Join(cm.RichText, "")
			lines = append(lines, fmt.Sprintf("%s\tauthorId=%s\t%q", cm.Ref, orDash(cm.AuthorID), text))
		}
		sort.Strings(lines)
		for _, l := range lines {
			emit("commententry\t%s\t%s", name, l)
		}
	}
}

func orDefault(s string) string {
	if s == "" {
		return "0"
	}
	return s
}
