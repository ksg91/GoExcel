package main

import (
    "fmt"
    "log"
    "net/http"
)

func ping(w http.ResponseWriter, r *http.Request){
    fmt.Fprintf(w, "Pong!");
}

func exportExcel(w http.ResponseWriter, r *http.Request){
    fmt.Fprintf(w, "Exporting Excel!");
}

func exportExcel(w http.ResponseWriter, r *http.Request){
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
