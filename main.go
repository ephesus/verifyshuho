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

type ShuhoEntry struct {
	SDate       string
	SCaseNum    string
	SType       string
	STWordCount string
	SCWordCount string
	SAuthor     string
}

func greeting() {
	fmt.Println("------------------------")
	fmt.Println("Verify Shuho and Invoice")
	fmt.Println("------------------------\n")
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
	finvoice, err := excelize.OpenFile(invoiceFileName)
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
		fmt.Printf("%s\n", invoiceEntryToString(entry))
	}

	fmt.Println("fin")
	return
}

func invoiceEntryToString(ie InvoiceEntry) string {
	return fmt.Sprintf("%s %s %s %s %s", ie.rowNum, ie.ICaseNum, ie.IType, ie.IWordCount, ie.rate)
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

		regres, err := regexp.Match("\\d+-\\d+-\\d+$", []byte(row[3]))
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

func parseShuho(f *excelize.File) []ShuhoEntry {
	entries := []ShuhoEntry{}

	for index, name := range f.GetSheetList() {
		fmt.Println("SHUHO SHEET NAME", index, name)
	}

	return entries
}
