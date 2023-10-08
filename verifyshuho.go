/*
	Go program to verify my Shuho and Invoice excel files for work
	by James Rubingh
	james@wrive.com 2023
*/

package main

import (
	"fmt"
	"os"
	"regexp"

	"github.com/xuri/excelize/v2"
)

type InvoiceEntry struct {
	rowNum     string
	IDate      string
	ICaseNum   string
	IType      string
	IWordCount string
	rate       string
}

func (e InvoiceEntry) String() string {
	return fmt.Sprintf("%s, %s, %s, %s, %s, %s,", e.rowNum, e.IDate, e.ICaseNum, e.IType, e.IWordCount, e.rate)
}

type ShuhoEntry struct {
	SDate       string
	SCaseNum    string
	SType       string
	STWordCount string
	SCWordCount string
	SAuthor     string
}

func (e ShuhoEntry) String() string {
	return fmt.Sprintf("%s, %s, %s, %s", e.SDate, e.SCaseNum, e.SType, e.SAuthor)
}

func greeting() {
	fmt.Println("------------------------")
	fmt.Println("Verify Shuho and Invoice")
	fmt.Println("------------------------")
}

func main() {
	if len(os.Args[1:]) != 2 {
		fmt.Println("ERROR Usage: ./verifyshuho <Shuho.xlsx> <Invoice.xlsx>")
		return
	}

	shuhoFileName := os.Args[1]
	invoiceFileName := os.Args[2]

	var shuhoEntries []ShuhoEntry
	var invoiceEntries []InvoiceEntry

	greeting()

	fshuho, err := excelize.OpenFile(shuhoFileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	finvoice, err := excelize.OpenFile(invoiceFileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the invoice spreadsheet.
		if err := finvoice.Close(); err != nil {
			fmt.Println(err)
		}
		// Close the shuho spreadsheet.
		if err := fshuho.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	if err != nil {
		fmt.Println("ERROR:", err)
		return
	}

	shuhoEntries = parseShuho(fshuho)
	invoiceEntries = parseInvoice(finvoice)

	if shuhoEntries == nil || invoiceEntries == nil {
		fmt.Println("Empty Shuho or Invoice Entries variable")
		return
	}

	for _, entry := range invoiceEntries {
		fmt.Printf("%s\n", entry)
	}

	fmt.Println("fin")
}

func parseInvoice(f *excelize.File) []InvoiceEntry {
	entries := []InvoiceEntry{}
	var sheetName string

	for index, name := range f.GetSheetList() {
		fmt.Println("INVOICE SHEET NAME", index, name)
		sheetName = name
	}

	rows, err := f.Rows(sheetName)
	if err != nil {
		fmt.Println(err)
		return entries
	}

	if rows == nil {
		fmt.Println("No Rows")
		return entries
	}

	dateRe := regexp.MustCompile(`\d+-\d+-\d+$`)
	if err != nil {
		return entries
	}

	for rows.Next() {
		var ie InvoiceEntry
		row, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
			return entries
		}

		//no row
		if row == nil || len(row) < 5 {
			continue
		}
		//not a complete row
		if rowNotComplete(row) {
			continue
		}

		regres := dateRe.Match([]byte(row[3]))
		if err != nil {
			fmt.Println(err)
			return entries
		}
		//first column cell is not a date string e.g. 6/20
		if !regres {
			continue
		}

		if len(row) > 5 {
			ie.rowNum = row[0]
			ie.IDate = row[3]
			ie.ICaseNum = row[1]
			ie.IType = row[2]
			ie.IWordCount = row[4]
			ie.rate = row[5]

			//		fmt.Printf("%s\n", row)
		}

		entries = append(entries, ie)
	}

	return entries
}

// make sure that the row has required fields
func rowNotComplete(row []string) bool {
	//check that each field has a value
	if (row[0] == "") || (row[3] == "") || (row[1] == "") || (row[2] == "") || (row[4] == "") || (row[5] == "") {
		return true
	}

	//check for default casenum "ALP-"
	match, _ := regexp.MatchString(`^(?i)ALP-$`, row[1])

	//must be last
	return match
}

func parseShuho(f *excelize.File) []ShuhoEntry {
	entries := []ShuhoEntry{}

	for index, name := range f.GetSheetList() {
		fmt.Println("SHUHO SHEET NAME", index, name)
	}

	return entries
}
