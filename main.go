package main

import (
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"log"
	"net/http"
)

type ExcelInput struct {
	FileName string  `json:"fileName"`
	Sheets   []Sheet `json:sheets`
}

type Sheet struct {
	Name        string      `json:"name"`
	MergedCells [][]string  `json:"mergedCells"`
	CellData    []CellData  `json:"cellData"`
	TableExport TableExport `json:"tableExport"`
}

type CellData struct {
	Cell       string  `json:"cell"`
	Text       string  `json:"text"`
	FontFamily string  `json:"fontFamily"`
	FontSize   float64 `json:"fontSize"`
	IsBold     bool    `json:"isBold"`
	IsItalic   bool    `json:"isItalic"`
	Color      string  `json:"color"`
}

type TableExport struct {
	TableStarts   string `json:"tableStarts"`
	SerialNumbers struct {
		AutoAdd bool   `json:"autoAdd"`
		Title   string `json:"title"`
	} `json:"serialNumbers"`
	TableHeading struct {
		FirstRowOfTableData bool     `json:"firstRowOfTableData"`
		HeadingTitles       []string `json:"headingTitles"`
		IsBold              bool     `json:"isBold"`
	} `json:"tableHeading"`
	TableData [][]string `json:"tableData"`
}

func ping(w http.ResponseWriter, r *http.Request) {
	fmt.Fprintf(w, "Pong!")
}

func exportExcel(w http.ResponseWriter, r *http.Request) {
	var e ExcelInput

	if r.Method == "GET" {
		e = getSampleData()
	} else {
		// Try to decode the request body into the struct. If there is an error,
		// respond to the client with the error message and a 400 status code.
		err := json.NewDecoder(r.Body).Decode(&e)
		if err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		} else {
		}
	}
	// Get the Excel file with the user input data
	file := ProcessExcelInput(e)

	// Set the headers necessary to get browsers to interpret the downloadable file
	w.Header().Set("Content-Type", "application/octet-stream")
	w.Header().Set("Content-Disposition", fmt.Sprintf("attachment;filename=\"%s\"", e.FileName))
	w.Header().Set("File-Name", e.FileName)
	w.Header().Set("Content-Transfer-Encoding", "binary")
	w.Header().Set("Expires", "0")
	file.Write(w)
}

func ProcessExcelInput(e ExcelInput) *excelize.File {
	f := excelize.NewFile()
	isFirstSheet := true
	for iii := 0; iii < len(e.Sheets); iii++ {
		ProcessSheetInput(f, e.Sheets[iii], isFirstSheet)
		isFirstSheet = false
	}

	return f
}

func ProcessSheetInput(f *excelize.File, s Sheet, isFirstSheet bool) {
	if isFirstSheet {
		f.SetSheetName("Sheet1", s.Name)
	} else {
		f.NewSheet(s.Name)
	}
	ProcessMergedCells(f, s.MergedCells)
	for iii := 0; iii < len(s.CellData); iii++ {
		ProcessCellData(f, s.CellData[iii])
	}
	ProcessTableExport(f, s.TableExport)
}

func ProcessMergedCells(f *excelize.File, mergedCells [][]string) {
	for iii := 0; iii < len(mergedCells); iii++ {
		f.MergeCell(GetActiveSheetName(f), mergedCells[iii][0], mergedCells[iii][1])
	}
}

func ProcessCellData(f *excelize.File, cd CellData) {
	f.SetCellRichText(GetActiveSheetName(f), cd.Cell, []excelize.RichTextRun{
		{
			Text: cd.Text,
			Font: &excelize.Font{
				Bold:   cd.IsBold,
				Italic: cd.IsItalic,
				Family: cd.FontFamily,
				Size:   cd.FontSize,
				Color:  cd.Color,
			},
		},
	})
}

func ProcessTableExport(f *excelize.File, te TableExport) {
	x, y, _ := excelize.CellNameToCoordinates(te.TableStarts)
	startX := x
	//startY := y
	isTableHeading := true
	serialNoCounter := 1

	if !te.TableHeading.FirstRowOfTableData {
		isTableHeading = false
		if te.SerialNumbers.AutoAdd {
			currCell, _ := excelize.CoordinatesToCellName(x, y)
			f.SetCellRichText(GetActiveSheetName(f), currCell, []excelize.RichTextRun{
				{
					Text: te.SerialNumbers.Title,
					Font: &excelize.Font{
						Bold: te.TableHeading.IsBold,
					},
				},
			})
			x++
		}

		for iii := 0; iii < len(te.TableHeading.HeadingTitles); iii++ {
			currCell, _ := excelize.CoordinatesToCellName(x, y)

			f.SetCellRichText(GetActiveSheetName(f), currCell, []excelize.RichTextRun{
				{
					Text: te.TableHeading.HeadingTitles[iii],
					Font: &excelize.Font{
						Bold: te.TableHeading.IsBold,
					},
				},
			})
			x++
		}
		y++
	}

	for iii := 0; iii < len(te.TableData); iii++ {
		x = startX
		if isTableHeading {
			if te.SerialNumbers.AutoAdd {
				currCell, _ := excelize.CoordinatesToCellName(x, y)
				f.SetCellRichText(GetActiveSheetName(f), currCell, []excelize.RichTextRun{
					{
						Text: te.SerialNumbers.Title,
						Font: &excelize.Font{
							Bold: te.TableHeading.IsBold,
						},
					},
				})
				x++
			}
			for jjj := 0; jjj < len(te.TableData[iii]); jjj++ {
				currCell, _ := excelize.CoordinatesToCellName(x, y)
				f.SetCellRichText(GetActiveSheetName(f), currCell, []excelize.RichTextRun{
					{
						Text: te.TableData[iii][jjj],
						Font: &excelize.Font{
							Bold: te.TableHeading.IsBold,
						},
					},
				})
				x++
			}
			isTableHeading = false
		} else {
			currCell, _ := excelize.CoordinatesToCellName(x, y)
			if te.SerialNumbers.AutoAdd {
				f.SetCellValue(GetActiveSheetName(f), currCell, serialNoCounter)
				serialNoCounter++
				x++
			}
			for jjj := 0; jjj < len(te.TableData[iii]); jjj++ {
				currCell, _ := excelize.CoordinatesToCellName(x, y)
				f.SetCellValue(GetActiveSheetName(f), currCell, te.TableData[iii][jjj])
				x++
			}
		}
		y++
	}
}

func GetActiveSheetName(f *excelize.File) string {
	return f.GetSheetName(f.GetActiveSheetIndex())
}

func importExcel(w http.ResponseWriter, r *http.Request) {
	fmt.Fprintf(w, "Importing Excel!")
}

func getSampleData() ExcelInput {
	testJson := `{
		"fileName" : "TestFile.xlsx",
		"sheets" : [
		{
			"name": "Users",
			"mergedCells": [
			["A1", "A5"],
			["B1", "B5"]
			],
			"cellData": [
			{
				"cell" : "C1",
				"text" : "Exporting User Data",
				"fontFamily" : "Times New Roman",
				"fontSize" : 16,
				"isBold" : true
			},
			{
				"cell" : "C2",
				"text" : "Report Generated on 2021-01-01 14:37:22",
				"fontFamily" : "Times New Roman",
				"fontSize" : 16,
				"isBold" : true,
				"isItalic" : true,
				"color" : "FF0000"
			}
			],
			"tableExport": {
				"tableStarts": "A6",
				"serialNumbers": {
					"autoAdd": true,
					"title": "Sr. No"
				},
				"tableHeading": {
					"firstRowOfTableData": false,
					"headingTitles": [
						"Name",
						"Email",
						"City"
					],
					"isBold": true
				},
				"tableData": [
					["Kishan Gor", "me@kishan.co", "Pune"],
					["Kishan Gor Second", "me@kishan.co", "Pune"]
				]
			}
		}
		]
	}`
	var excelInput ExcelInput
	json.Unmarshal([]byte(testJson), &excelInput)
	return excelInput
}

func handleRequests() {
	http.HandleFunc("/", ping)
	http.HandleFunc("/export", exportExcel)
	http.HandleFunc("/import", importExcel)
	log.Fatal(http.ListenAndServe(":10000", nil))
}

func main() {
	handleRequests()
}
