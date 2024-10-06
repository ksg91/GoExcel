package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

// ExcelInput is Input format for generating Excel
type ExcelInput struct {
	FileName string  `json:"fileName"`
	Sheets   []Sheet `json:sheets`
}

// Sheet is a representation of Sheet data in Excel
type Sheet struct {
	Name        string      `json:"name"`
	MergedCells [][]string  `json:"mergedCells"`
	CellData    []CellData  `json:"cellData"`
	TableExport TableExport `json:"tableExport"`
}

// CellData is a Data of a particular Cell that needs to be written in Excel
type CellData struct {
	Cell       string  `json:"cell"`
	Text       string  `json:"text"`
	FontFamily string  `json:"fontFamily"`
	FontSize   float64 `json:"fontSize"`
	IsBold     bool    `json:"isBold"`
	IsItalic   bool    `json:"isItalic"`
	Color      string  `json:"color"`
}

// TableExport is a tabular data that needs to be written from a given starting position
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
		}
	}
	// Get the Excel file with the user input data
	file := processExcelInput(e)

	// Set the headers necessary to get browsers to interpret the downloadable file
	w.Header().Set("Content-Type", "application/octet-stream")
	w.Header().Set("Content-Disposition", fmt.Sprintf("attachment;filename=\"%s\"", e.FileName))
	w.Header().Set("File-Name", e.FileName)
	w.Header().Set("Content-Transfer-Encoding", "binary")
	w.Header().Set("Expires", "0")
	file.Write(w)
}

func processExcelInput(e ExcelInput) *excelize.File {
	f := excelize.NewFile()
	isFirstSheet := true
	for iii := 0; iii < len(e.Sheets); iii++ {
		processSheetInput(f, e.Sheets[iii], isFirstSheet, iii)
		isFirstSheet = false
	}

	return f
}

func processSheetInput(f *excelize.File, s Sheet, isFirstSheet bool, sheetIndex int) {
	sheetName := "Sheet1"
	if s.Name == "" {
		sheetName = "Sheet" + strconv.Itoa(sheetIndex+1)
	}
	if isFirstSheet && sheetName != "Sheet1" {
		if sheetName != "Sheet1" {
			f.SetSheetName("Sheet1", sheetName)
		}
	} else {
		iSheetIndex := f.NewSheet(sheetName)
		f.SetActiveSheet(iSheetIndex)
	}
	processMergedCells(f, s.MergedCells)
	for iii := 0; iii < len(s.CellData); iii++ {
		processCellData(f, s.CellData[iii])
	}
	processTableExport(f, s.TableExport)
}

func processMergedCells(f *excelize.File, mergedCells [][]string) {
	for iii := 0; iii < len(mergedCells); iii++ {
		f.MergeCell(getActiveSheetName(f), mergedCells[iii][0], mergedCells[iii][1])
	}
}

func processCellData(f *excelize.File, cd CellData) {
	f.SetCellRichText(getActiveSheetName(f), cd.Cell, []excelize.RichTextRun{
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

func processTableExport(f *excelize.File, te TableExport) {
	x, y, _ := excelize.CellNameToCoordinates(te.TableStarts)
	startX := x
	//startY := y
	isTableHeading := true
	serialNoCounter := 1

	if !te.TableHeading.FirstRowOfTableData {
		isTableHeading = false
		if te.SerialNumbers.AutoAdd {
			currCell, _ := excelize.CoordinatesToCellName(x, y)
			f.SetCellRichText(getActiveSheetName(f), currCell, []excelize.RichTextRun{
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

			f.SetCellRichText(getActiveSheetName(f), currCell, []excelize.RichTextRun{
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
				f.SetCellRichText(getActiveSheetName(f), currCell, []excelize.RichTextRun{
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
				f.SetCellRichText(getActiveSheetName(f), currCell, []excelize.RichTextRun{
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
				f.SetCellValue(getActiveSheetName(f), currCell, serialNoCounter)
				serialNoCounter++
				x++
			}
			for jjj := 0; jjj < len(te.TableData[iii]); jjj++ {
				currCell, _ := excelize.CoordinatesToCellName(x, y)
				// Try to parse the string value as a number
				if num, err := strconv.Atoi(te.TableData[iii][jjj]); err == nil {
					f.SetCellValue(getActiveSheetName(f), currCell, num)
				} else if num, err := strconv.ParseFloat(te.TableData[iii][jjj], 64); err == nil {
					f.SetCellValue(getActiveSheetName(f), currCell, num)
				} else {
					// If not a number, set it as text
					f.SetCellValue(getActiveSheetName(f), currCell, te.TableData[iii][jjj])
				}
				x++
			}
		}
		y++
	}
}

func getActiveSheetName(f *excelize.File) string {
	return f.GetSheetName(f.GetActiveSheetIndex())
}

func importExcel(w http.ResponseWriter, r *http.Request) {
	fmt.Fprintf(w, "Importing Excel!")
}

func getSampleData() ExcelInput {
	exampleInput, _ := ioutil.ReadFile("exampleExcelInput.json")
	var excelInput ExcelInput
	json.Unmarshal([]byte(exampleInput), &excelInput)
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
