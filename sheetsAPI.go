package googlesheets4go

import (
	"context"
	"fmt"
	"google.golang.org/api/option"
	"log"
	"net/http"
	"strings"
	"time"

	"google.golang.org/api/googleapi"
	"google.golang.org/api/sheets/v4"
)

func (receiver *SheetsAPI) Build(client *http.Client, subject string, context *context.Context) *SheetsAPI {
	service, err := sheets.NewService(*context, option.WithHTTPClient(client))
	if err != nil {
		log.Println(err.Error())
		panic(err)
	}
	receiver.Service = service
	receiver.Subject = subject
	log.Printf("SheetAPI --> \n"+
		"\tService: %s\n"+
		"\tUserEmail: %s\n", receiver.Service.BasePath, receiver.Subject,
	)
	return receiver
}

type SheetsAPI struct {
	Service *sheets.Service
	Subject string
}

func (receiver *SheetsAPI) PrintToSheet(spreadsheetId, a1Notation, majorDimension string, values [][]interface{}, overwrite bool) interface{} {
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

func (receiver *SheetsAPI) CreateSpreadsheet(spreadsheetName string) *sheets.Spreadsheet {
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

func (receiver *SheetsAPI) RenameSpreadSheet(spreadsheetId, newTitle string) (*sheets.Spreadsheet, error) {

	spreadsheetProperties := &sheets.SpreadsheetProperties{Title: newTitle}
	updateSpreadsheetPropertiesRequest := &sheets.UpdateSpreadsheetPropertiesRequest{Properties: spreadsheetProperties, Fields: "*"}
	request := &sheets.Request{UpdateSpreadsheetProperties: updateSpreadsheetPropertiesRequest}
	var requests = []*sheets.Request{request}
	batchUpdateSpreadsheetRequest := &sheets.BatchUpdateSpreadsheetRequest{Requests: requests}
	response, err := receiver.Service.Spreadsheets.BatchUpdate(spreadsheetId, batchUpdateSpreadsheetRequest).Do()
	if err != nil {
		log.Println(err, err.Error())
	}
	log.Printf("Renamed SpreadsheetID: [%s] is now \"%s\"\n", spreadsheetId, response.UpdatedSpreadsheet.Properties.Title)
	return response.UpdatedSpreadsheet, err
}

func (receiver *SheetsAPI) CreateSheet(spreadsheetId, newSheetName string) *sheets.BatchUpdateSpreadsheetResponse {
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

func (receiver *SheetsAPI) RenameTab(spreadsheet *sheets.Spreadsheet, oldSheetName, newSheetName string) *sheets.BatchUpdateSpreadsheetResponse {
	sheetId := receiver.GetSheetByTabName(spreadsheet, oldSheetName).Properties.SheetId
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

func (receiver *SheetsAPI) GetSheetValues(spreadsheetId, a1Notation string) [][]interface{} {
	sheetOutputValues, err := receiver.Service.Spreadsheets.Values.Get(spreadsheetId, a1Notation).Do()
	if err != nil {
		log.Println(err.Error())
		panic(err)
	}
	return sheetOutputValues.Values
}

func (receiver *SheetsAPI) GetColumnValues(spreadsheetId, a1Notation string) []interface{} {
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

func (receiver *SheetsAPI) GetColumnValuesAsStringMap(spreadsheetId, a1Notation string, toLower bool) map[string]bool {
	m := make(map[string]bool)
	for _, s := range receiver.GetColumnValuesAsString(spreadsheetId, a1Notation, toLower) {
		if m[s] == false {
			m[s] = true
		}
	}
	return m
}

func (receiver *SheetsAPI) GetColumnValuesAsString(spreadsheetId, a1Notation string, toLower bool) []string {
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

func (receiver *SheetsAPI) GetSheetByTabName(spreadsheet *sheets.Spreadsheet, sheetName string) *sheets.Sheet {
	for _, sheet := range spreadsheet.Sheets {
		if sheet.Properties.Title == sheetName {
			return sheet
		}
	}
	log.Println(googleapi.Error{Body: "Sheet SendEmail " + sheetName + " not found in SpreadsheetID: " + spreadsheet.SpreadsheetId, Message: "Sheet not found"})
	return nil
}

func (receiver *SheetsAPI) ClearValues(spreadsheetID, a1Notation string) *sheets.ClearValuesResponse {
	response, err := receiver.Service.Spreadsheets.Values.Clear(spreadsheetID, a1Notation, &sheets.ClearValuesRequest{}).Fields("*").Do()
	if err != nil {
		log.Println(err.Error())
		return nil
	}
	log.Printf("Cleared %s [%s]\n", spreadsheetID, a1Notation)
	return response
}
