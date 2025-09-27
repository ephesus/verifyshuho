package main

// by James Rubingh
// Compare Invoices and Shuhos for mismatching case numbers,
// mismatching dates, etc.

import (
	"fmt"
	"os"
	"regexp"

	"github.com/xuri/excelize/v2"
)

type CaseEntry interface {
	//give an ID signature for the entry (casenumber + date + wordcount + type)
	Identify() string

	//give only the casenumber for the entry
	Casenumber() string
}

//create InvoiceEntry for each Invoice row, and methods implementing interfaces
type InvoiceEntry struct {
	rowNum     string
	IDate      string
	ICaseNum   string
	IType      string
	IWordCount string
	rate       string
}

//implement CaseEntry interface
func (i InvoiceEntry) Identify() string {
	return fmt.Sprintf("Case: %s, Type: %s, Date: %s, Word Count: %s", i.ICaseNum, i.IType, i.IDate, i.IWordCount)
}

func (i InvoiceEntry) Casenumber() string {
	return i.ICaseNum
}

//implement Stringer interface
func (i InvoiceEntry) String() string {
	return fmt.Sprintf("Date: %s, Case: %s, Type: %s", i.IDate, i.ICaseNum, i.IType)
}

//create ShuhoEntry for each Invoice row, and methods implementing interfaces
type ShuhoEntry struct {
	SDate      string
	SCaseNum   string
	SType      string
	SWordCount string
	SAuthor    string
}

//implement CaseEntry interface
func (i ShuhoEntry) Identify() string {
	return fmt.Sprintf("Case: %s, Type: %s, Date: %s, Word Count: %s", i.SCaseNum, i.SType, i.SDate, i.SWordCount)
}

func (i ShuhoEntry) Casenumber() string {
	return i.SCaseNum
}

//implement Stringer interface
func (i ShuhoEntry) String() string {
	return i.Identify()
}

func greeting() {
	fmt.Println("------------------------")
	fmt.Println("Verify Shuho and Invoice")
	fmt.Println("------------------------\n")
}

func countEntries(entries []CaseEntry) int {
	return len(entries)
}

func main() {
	if len(os.Args[1:]) != 2 {
		fmt.Println("ERROR Usage: ./verifyshuho <Shuho.xlsx> <Invoice.xlsx>")
		return
	}

	shuhoFileName := os.Args[1]
	invoiceFileName := os.Args[2]

	var shuhoEntries []CaseEntry
	var invoiceEntries []CaseEntry

	greeting()

	fshuho, err_s := excelize.OpenFile(shuhoFileName)
	finvoice, err_i := excelize.OpenFile(invoiceFileName)
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

	if err_s != nil {
		fmt.Println("ERROR Shuho File:", err_s)
		return
	}
	if err_i != nil {
		fmt.Println("ERROR Invoice File::", err_s)
		return
	}

	shuhoEntries = parseShuho(fshuho)
	invoiceEntries = parseInvoice(finvoice)

	if shuhoEntries == nil || invoiceEntries == nil {
		fmt.Println("Empty Shuho or Invoice Entries variable")
		return
	}

	fmt.Println("Shuho Entries:")
	for _, entry := range shuhoEntries {
		fmt.Printf("%s\n", entry.Identify())
	}

	fmt.Println("Invoice Entries:")
	for _, entry := range invoiceEntries {
		fmt.Printf("%s\n", entry.Identify())
	}

	fmt.Printf("Number of Invoices: %d\n", countEntries(invoiceEntries))
	fmt.Printf("Number of Shuhos: %d\n", countEntries(shuhoEntries))

	fmt.Println("fin")
	return
}

func parseInvoice(f *excelize.File) []CaseEntry {
	entries := []CaseEntry{}
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

func parseShuho(f *excelize.File) []CaseEntry {
	entries := []CaseEntry{}

	for index, name := range f.GetSheetList() {
		fmt.Println("INVOICE SHEET NAME", index, name)
		sheetName := name

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
			var se ShuhoEntry
			row, err := rows.Columns()

			fmt.Println(row)

			if err != nil {
				fmt.Println(err)
				return entries
			}

			//no row
			if row == nil || len(row) < 5 {
				continue
			}

			//check for a date, then check that the case doesnt' include x
			regres, err := regexp.Match("\\d+/\\d+$", []byte(row[0]))
			if err != nil {
				fmt.Println(err)
				return entries
			}
			regres2, err2 := regexp.Match("^x", []byte(row[0]))
			if err2 != nil {
				fmt.Println(err2)
				return entries
			}

			//first column cell is not a date string e.g. 6/20
			if !regres || regres2 {
				continue
			}

			if len(row) > 5 {
				se.SDate = row[0]
				se.SCaseNum = row[1]
				se.SType = row[2]
				se.SWordCount = row[4]
				se.SAuthor = row[6]
				//		fmt.Printf("%s\n", row)
			}

			entries = append(entries, se)
		}
	}

	return entries
}
