package exporter

import (
	"bufio"
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"
	"regexp"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func getSheetNames(xlsxFile *excelize.File) []string {
	sheets := xlsxFile.GetSheetMap()
	var sheetNames []string
	for _, sheet := range sheets {
		sheetNames = append(sheetNames, sheet)
	}
	return sheetNames
}

func getDefaultJsonObjectKeys(tableRows [][]string) []string {
	return tableRows[0]
}

func getTableValuesWithoutHeader(tableRows [][]string) [][]string {
	tableRowsWithoutHeader := tableRows[1:]
	return tableRowsWithoutHeader
}

func intSliceContains(array []int, value int) bool {
	for _, arrayItem := range array {
		if arrayItem == value {
			return true
		}
	}

	return false
}

func init() {
	log.SetPrefix("Error:")
	log.SetFlags(0)
}

var reader = bufio.NewReader(os.Stdin)

func Exporter() {
	// input params
	filePathValue := flag.String("file-path", "", "specify excel file folder")
	flag.Parse()

	path := *filePathValue
	xlsxFile, xlsxError := excelize.OpenFile(path)

	if xlsxError != nil {
		log.Fatalf("Error: %s", xlsxError.Error())
	}

	sheetNames := getSheetNames(xlsxFile)
	fmt.Println("<<Sheets>>")
	for index, sheetName := range sheetNames {
		fmt.Printf("(%v) => %v \n", index, sheetName)
	}
	fmt.Print("Select sheet index:")
	selectedSheetIndexString, _ := reader.ReadString('\n')
	parsedSelectedSheetIndex, parsedSelectedSheetIndexError := strconv.Atoi(strings.TrimSuffix(selectedSheetIndexString, "\n"))

	if parsedSelectedSheetIndexError != nil {
		log.Fatal(parsedSelectedSheetIndexError)
	}
	if parsedSelectedSheetIndex < 0 || parsedSelectedSheetIndex > len(sheetNames)-1 {
		log.Fatalf("Value (%v) is not in range [0 - %v]", parsedSelectedSheetIndex, len(sheetNames)-1)
	}

	selectedSheetName := sheetNames[parsedSelectedSheetIndex]
	tableRows := xlsxFile.GetRows(selectedSheetName)

	tableHeaderValues := getDefaultJsonObjectKeys(tableRows)
	var disabledTableHeaderValuesIndexes []int

	fmt.Println("<<Table header values>>")
	for index, tableHeaderValue := range tableHeaderValues {
		fmt.Printf("(%v) => %v \n", index, tableHeaderValue)
	}
	fmt.Print("hide some columns (use ',' as separator):")
	disabledTableHeaderIndexesInput, _ := reader.ReadString('\n')
	disabledTableHeadersIndexesWithoutNewLine := strings.TrimSuffix(disabledTableHeaderIndexesInput, "\n")

	if disabledTableHeadersIndexesWithoutNewLine != "" {
		if disabledTableHeaderIsInCorrectFormat, _ := regexp.MatchString("^[0-9,]*$", disabledTableHeadersIndexesWithoutNewLine); disabledTableHeaderIsInCorrectFormat {
			disabledTableHeadersIndexes := strings.Split(disabledTableHeadersIndexesWithoutNewLine, ",")
			for _, disabledTableHeadersIndex := range disabledTableHeadersIndexes {
				if parsedIndex, parsedIndexError := strconv.Atoi(disabledTableHeadersIndex); disabledTableHeadersIndex != "" && parsedIndexError == nil {
					disabledTableHeaderValuesIndexes = append(disabledTableHeaderValuesIndexes, parsedIndex)
				} else {
					fmt.Printf("Value ('%v') is not a number! >> Skipped \n", disabledTableHeadersIndex)
				}
			}
		} else {
			log.Fatal("Incorrect format of value!")
		}
	}

	tableRowsWithValues := getTableValuesWithoutHeader(tableRows)

	var dataForJson []map[string]string
	for _, tableRow := range tableRowsWithValues {
		jsonObjectFromRowValues := make(map[string]string)
		for tableRowValueIndex, tableRowValue := range tableRow {
			if !intSliceContains(disabledTableHeaderValuesIndexes, tableRowValueIndex) {
				jsonObjectFromRowValues[tableHeaderValues[tableRowValueIndex]] = tableRowValue
			}
		}
		dataForJson = append(dataForJson, jsonObjectFromRowValues)
	}

	jsonFile, jsonError := json.Marshal(dataForJson)
	if jsonError != nil {
		log.Fatalf("Error: %s", jsonError.Error())
	} else {
		jsonFileName := selectedSheetName + ".json"

		os.WriteFile(jsonFileName, jsonFile, 0600)
		fmt.Printf("%v exported!\n", jsonFileName)
	}

}
