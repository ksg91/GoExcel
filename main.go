package main

import (
	"context"
	b64 "encoding/base64"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/aws/aws-lambda-go/events"
	"github.com/aws/aws-lambda-go/lambda"
)

// ExcelInput is Input format for generating Excel
type ExcelInput struct {
	FileName string  `json:"fileName"`
	Sheets   []Sheet `json:sheets`
}

type HTTPHeaders struct {
	ContentType             string `json:"Content-Type"`
	ContentDisposition      string `json:"Content-Disposition"`
	FileName                string `json:"File-Name"`
	ContentTransferEncoding string `json:"Content-Transfer-Encoding"`
}

type ExcelResponse struct {
	Headers         HTTPHeaders `json:"headers"`
	StatusCode      int         `json:"statusCode"`
	Body            string      `json:"body"`
	IsBase64Encoded bool        `json:"isBase64Encoded"`
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
	// w.Header().Set("Content-Type", "application/octet-stream")
	// w.Header().Set("Content-Disposition", fmt.Sprintf("attachment;filename=\"%s\"", e.FileName))
	// w.Header().Set("File-Name", e.FileName)
	// w.Header().Set("Content-Transfer-Encoding", "binary")
	// w.Header().Set("Expires", "0")
	buff, err := file.WriteToBuffer()

	if err != nil {
		http.Error(w, err.Error(), http.StatusBadRequest)
		return
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(b64.StdEncoding.EncodeToString(buff.Bytes()))
}

func processExcelInput(e ExcelInput) *excelize.File {
	f := excelize.NewFile()
	isFirstSheet := true
	for iii := 0; iii < len(e.Sheets); iii++ {
		processSheetInput(f, e.Sheets[iii], isFirstSheet)
		isFirstSheet = false
	}

	return f
}

func processSheetInput(f *excelize.File, s Sheet, isFirstSheet bool) {
	if isFirstSheet {
		f.SetSheetName("Sheet1", s.Name)
	} else {
		iSheetIndex := f.NewSheet(s.Name)
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
				f.SetCellValue(getActiveSheetName(f), currCell, te.TableData[iii][jjj])
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

func handleRequests(ctx context.Context, request events.APIGatewayProxyRequest) (events.APIGatewayProxyResponse, error) {

	fmt.Printf("Processing request data for request %s.\n", request.RequestContext.RequestID)
	fmt.Printf("Body size = %d.\n", len(request.Body))

	fmt.Println("Headers:")
	for key, value := range request.Headers {
		fmt.Printf("    %s: %s\n", key, value)
	}

	var excelInput ExcelInput

	err := json.Unmarshal([]byte(request.Body), &excelInput)

	if err != nil {
		fmt.Printf("There was an error decoding the json. err = %s", err)
	}

	fmt.Printf("ExcelInput: %#v", request.Body)
	fmt.Printf("ExcelInput: %#v", excelInput)

	file := processExcelInput(excelInput)

	buff, err := file.WriteToBuffer()

	if err != nil {
		return events.APIGatewayProxyResponse{StatusCode: 500}, err
	}

	bodyContent := b64.StdEncoding.EncodeToString(buff.Bytes())

	// headers := HTTPHeaders{ContentType: "application/octet-stream", ContentDisposition: fmt.Sprintf("attachment;filename=%s", excelInput.FileName),
	// 	FileName:                excelInput.FileName,
	// 	ContentTransferEncoding: "binary"}

	return events.APIGatewayProxyResponse{
		StatusCode:        200,
		Headers:           map[string]string{"content-type": "application/octet-stream", "content-disposition": fmt.Sprintf("attachment;filename=%s", excelInput.FileName), "file-name": excelInput.FileName, "Content-Transfer-Encoding": "binary"},
		MultiValueHeaders: map[string][]string{},
		Body:              bodyContent,
		IsBase64Encoded:   true,
	}, nil

}

func main() {
	lambda.Start(handleRequests)
}
