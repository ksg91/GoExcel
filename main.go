package main

import (
    "fmt"
    "log"
    "net/http"
    "github.com/360EntSecGroup-Skylar/excelize/v2"
)

func ping(w http.ResponseWriter, r *http.Request){
    fmt.Fprintf(w, "Pong!");
}

func exportExcel(w http.ResponseWriter, r *http.Request){
    // Get the Excel file with the user input data
  file := PrepareDummyAndReturnExcel()

  // Set the headers necessary to get browsers to interpret the downloadable file
  w.Header().Set("Content-Type", "application/octet-stream")
  w.Header().Set("Content-Disposition", "attachment;filename=\"demoExcel.xlsx\"")
  w.Header().Set("File-Name", "demoExcel.xlsx")
  w.Header().Set("Content-Transfer-Encoding", "binary")
  w.Header().Set("Expires", "0")
  file.Write(w)
}

func PrepareDummyAndReturnExcel() *excelize.File {
   f := excelize.NewFile()
   f.SetCellValue("Sheet1", "A1", "Kishan")
   f.SetCellValue("Sheet1", "A2", "Gor")
   return f
}


func importExcel(w http.ResponseWriter, r *http.Request){
    fmt.Fprintf(w, "Importing Excel!");
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
