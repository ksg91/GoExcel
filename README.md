
# GoExcel

## Background
GoExcel is a small web service that can convert your data into an excel. 
This is being made an as exercise for me to get familiar with GoLang. But this will also be useful for projects where programming languages like PHP are slow and heavy on memory to create large Excel sheet from data. Plan is to use this as a microservice which will help other apps to create Excel sheets using over HTTP requests. 

Remember, I am just learning the Go so code will be very messy.

## Usage
Run the script using `go run main.go`
This will start the server at port 10000. You can access the service at `http:://localhost:10000`
`GET /export` will export an Excel sheet with Sample Data hardcoded in the code

`POST /export` will export your ExcelInput JSON data into an Excel sheet. 
Example Payload:

    {
      "fileName": "TestFile.xlsx",
      "sheets": [
        {
          "name": "Users",
          "mergedCells": [
            [
              "A1",
              "A5"
            ],
            [
              "B1",
              "B5"
            ]
          ],
          "cellData": [
            {
              "cell": "C1",
              "text": "Exporting User Data",
              "fontFamily": "Times New Roman",
              "fontSize": 16,
              "isBold": true
            },
            {
              "cell": "C2",
              "text": "Report Generated on 2021-01-01 14:37:22",
              "fontFamily": "Times New Roman",
              "fontSize": 16,
              "isBold": true,
              "isItalic": true,
              "color": "FF0000"
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
              [
                "Kishan Gor",
                "me@kishan.co",
                "Pune"
              ],
              [
                "Kishan Gor Second",
                "me@kishan.co",
                "Pune"
              ]
            ]
          }
        }
      ]
    }

### To-Do
- Error Handling
- `/import` function that can read an excel and give output as json structure similar to `ExcelInput`
- Config Management