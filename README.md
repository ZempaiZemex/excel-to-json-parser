# excel-to-json-parser

Excel-to-json-parse is a tool for convert relation database style excel table to json.

## dependencies
- install **GO** >> https://go.dev/doc/install

## How to use it
- clone repository
- go to the project folder
- run `go run . -file-path <path to file>`
- select sheet
- **optional:** hide some columns 

## Excel table
| id | name  | age |
|----|-------|-----|
| 1  | John  | 18  |
| 2  | James | 25  |
| 3  | Jane  | 22  |

## Final JSON file
```
[
  { "age": "18", "id": "1", "name": "John" },
  { "age": "25", "id": "2", "name": "James" },
  { "age": "22", "id": "3", "name": "Jane" }
]
```
