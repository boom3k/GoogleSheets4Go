package googlesheets4go

import (
	"context"
	"fmt"
	"log"
	"net/http"
	"strings"
	"time"

	"google.golang.org/api/googleapi"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

var ctx = context.Background()

func Initialize(option *option.ClientOption, subject string) *GoogleSheets {
	service, err := sheets.NewService(ctx, *option)
	if err != nil {
		log.Println(err.Error())
		panic(err)
	}

	log.Printf("Initialized GoogleSheets4go as (%s)\n", subject)
	return &GoogleSheets{Service: service, Subject: subject}
}

type GoogleSheets struct {
	Service *sheets.Service
	Subject string
}

func (receiver *GoogleSheets) PrintToSheet(spreadsheetId, a1Notation, majorDimension string, values [][]interface{}, overwrite bool) interface{} {
	var valueRange sheets.ValueRange
	valueRange.MajorDimension = strings.ToUpper(majorDimension)
	valueRange.Values = values
	log.Println("Spreadsheet Write Request --> SpreadsheetID:[" + spreadsheetId + "], A1Notation:[" + a1Notation + "], TotalInserts[" + fmt.Sprint(len(values)) + "], overwrite[" + fmt.Sprint(overwrite) + "]")
	if overwrite == true {
		response, err := receiver.Service.Spreadsheets.Values.Update(spreadsheetId, a1Notation, &valueRange).ValueInputOption("RAW").Do()
		if err != nil {
			log.Println(err.Error())
			panic(err)
		}
		return response
	}
	response, err := receiver.Service.Spreadsheets.Values.Append(spreadsheetId, a1Notation, &valueRange).ValueInputOption("USER_ENTERED").Fields("*").Do()
	if err != nil {
		log.Println(err.Error())
		if strings.Contains(err.Error(), "Quota exceeded") {
			log.Println("Backing off for 2.5 seconds...")
			time.Sleep(time.Millisecond * 2500)
			return receiver.PrintToSheet(spreadsheetId, a1Notation, majorDimension, values, overwrite)
		}
		panic(err)
	}
	log.Println("Spreadsheet write request was successful...")
	return response
}

func (receiver *GoogleSheets) CreateSpreadsheet(spreadsheetName string) *sheets.Spreadsheet {
	ss := &sheets.Spreadsheet{}
	ss.Properties = &sheets.SpreadsheetProperties{Title: spreadsheetName}
	response, err := receiver.Service.Spreadsheets.Create(ss).Fields("*").Do()
	if err != nil {
		log.Println(err.Error())
		panic(err)
	}
	log.Println("Created spreadsheet -> ", spreadsheetName, " [", response.SpreadsheetId, "] @ "+response.SpreadsheetUrl)
	return response
}

func (receiver *GoogleSheets) CreateSheet(spreadsheetId, newSheetName string) *sheets.BatchUpdateSpreadsheetResponse {
	properties := &sheets.SheetProperties{Title: newSheetName}
	addSheetsRequest := &sheets.AddSheetRequest{Properties: properties}
	request := []*sheets.Request{{AddSheet: addSheetsRequest}}
	content := &sheets.BatchUpdateSpreadsheetRequest{Requests: request}
	response, err := receiver.Service.Spreadsheets.BatchUpdate(spreadsheetId, content).Fields("*").Do()
	if err != nil {
		log.Println(err.Error())
		panic(err)
	}
	return response
}

func (receiver *GoogleSheets) RenameSheet(spreadsheet *sheets.Spreadsheet, oldSheetName, newSheetName string) *sheets.BatchUpdateSpreadsheetResponse {
	sheetId := receiver.GetSheetByName(spreadsheet, oldSheetName).Properties.SheetId
	sheetProperties := &sheets.SheetProperties{Title: newSheetName, SheetId: sheetId}
	updateSheetPropertiesRequest := &sheets.UpdateSheetPropertiesRequest{Properties: sheetProperties, Fields: "title"}
	request := []*sheets.Request{{UpdateSheetProperties: updateSheetPropertiesRequest}}
	content := &sheets.BatchUpdateSpreadsheetRequest{Requests: request}
	response, err := receiver.Service.Spreadsheets.BatchUpdate(spreadsheet.SpreadsheetId, content).Fields("*").Do()
	if err != nil {
		log.Println(err.Error())
		panic(err)
	}
	return response
}

func (receiver *GoogleSheets) GetSheetValues(spreadsheetId, a1Notation string) [][]interface{} {
	sheetOutputValues, err := receiver.Service.Spreadsheets.Values.Get(spreadsheetId, a1Notation).Do()
	if err != nil {
		log.Println(err.Error())
		panic(err)
	}
	return sheetOutputValues.Values
}

func (receiver *GoogleSheets) GetColumnValues(spreadsheetId, a1Notation string) []interface{} {
	sheetOutputValues, err := receiver.Service.Spreadsheets.Values.Get(spreadsheetId, a1Notation).Do()
	if err != nil {
		log.Println(err.Error())
		panic(err)
	}
	var columnValues []interface{}

	for _, row := range sheetOutputValues.Values {
		for i := range row {
			columnValues = append(columnValues, row[i])
		}
	}

	return columnValues
}

func (receiver *GoogleSheets) GetColumnValuesAsStringMap(spreadsheetId, a1Notation string, toLower bool) map[string]bool {
	m := make(map[string]bool)
	for _, s := range receiver.GetColumnValuesAsString(spreadsheetId, a1Notation, toLower) {
		if m[s] == false {
			m[s] = true
		}
	}
	return m
}

func (receiver *GoogleSheets) GetColumnValuesAsString(spreadsheetId, a1Notation string, toLower bool) []string {
	sheetOutputValues, err := receiver.Service.Spreadsheets.Values.Get(spreadsheetId, a1Notation).Do()
	if err != nil {
		log.Println(err.Error())
		panic(err)
	}
	var columnValues []string

	for _, row := range sheetOutputValues.Values {
		for i := range row {
			columnValues = append(columnValues, row[i].(string))
		}
	}

	if toLower {
		for i := range columnValues {
			columnValues[i] = strings.ToLower(columnValues[i])
		}
	}
	return columnValues
}

func (receiver *GoogleSheets) GetSheetByName(spreadsheet *sheets.Spreadsheet, sheetName string) *sheets.Sheet {
	for _, sheet := range spreadsheet.Sheets {
		if sheet.Properties.Title == sheetName {
			return sheet
		}
	}
	log.Println(googleapi.Error{Body: "Sheet SendEmail " + sheetName + " not found in SpreadsheetID: " + spreadsheet.SpreadsheetId, Message: "Sheet not found"})
	return nil
}
