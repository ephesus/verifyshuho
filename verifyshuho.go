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
	"time"

	"github.com/xuri/excelize/v2"
)

// entry signatures are Date, Casenum, Type, Wordcount
type Entry interface {
	signature() string
	String() string
	Type() string
	Date() string
}

type InvoiceEntry struct {
	rowNum     string
	IDate      string
	ICaseNum   string
	IType      string
	IWordCount string
	rate       string
}

// stuct methods
func (e InvoiceEntry) signature() string {
	return fmt.Sprintf("%s %s %s", e.ICaseNum, e.IType, e.IWordCount)
}

func (e InvoiceEntry) String() string {
	return fmt.Sprintf("%s, %s, %s, %s, %s, %s", e.rowNum, e.IDate, e.ICaseNum, e.IType, e.IWordCount, e.rate)
}

func (e InvoiceEntry) Date() string {
	return e.IDate
}

func (e InvoiceEntry) Type() string {
	return e.IType
}

type ShuhoEntry struct {
	SDate       string
	SCaseNum    string
	SType       string
	SCWordCount string
	STWordCount string
	SAuthor     string
}

func getShuhoEntryWordCount(e ShuhoEntry) string {
	var wordcount string

	switch e.SType {
	case "翻訳":
		wordcount = e.STWordCount
	case "英文チェック":
		wordcount = e.SCWordCount
	default:
		//should never happen, the excel file restricts to the two above values
		fmt.Printf("NOTE: %s - %s, %s\n", e.SType, e.SDate, e.SCaseNum)
		wordcount = "UNKNOWN"
	}

	return wordcount
}

// stuct methods
func (e ShuhoEntry) signature() string {
	wordcount := getShuhoEntryWordCount(e)

	return fmt.Sprintf("%s %s %s", e.SCaseNum, e.SType, wordcount)
}

func (e ShuhoEntry) String() string {
	wordcount := getShuhoEntryWordCount(e)

	return fmt.Sprintf("%s, %s, %s, %s, %s", e.SDate, e.SCaseNum, e.SType, wordcount, e.SAuthor)
}

func (e ShuhoEntry) Date() string {
	return e.SDate
}

func (e ShuhoEntry) Type() string {
	return e.SType
}

// print error for structs satisfying Entry interface
func printEntryError(e Entry) {
	fmt.Printf("Error: %s\n", e.String())
	os.Exit(1)
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

	var shuhoEntries []Entry
	var invoiceEntries []Entry

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

	invoiceEntries = parseInvoice(finvoice)
	shuhoEntries = parseShuho(fshuho)

	if shuhoEntries == nil || invoiceEntries == nil {
		fmt.Println("Empty Shuho or Invoice Entries variable")
		return
	}

	fmt.Printf("Invoice Entries: %d\n", len(invoiceEntries))
	fmt.Printf("Shuho Entries: %d\n", len(shuhoEntries))
	fmt.Println("")
	fmt.Printf("Total Translations: %d\n", sumOfTranslations(invoiceEntries))
	fmt.Printf("Total Checks: %d\n", sumOfChecks(invoiceEntries))

	ensureNoDuplicateInvoiceEntries(invoiceEntries)
	ensureInvoiceEntriesAreInShuho(invoiceEntries)

	//main
}

// FOR DEBUG
// printAllEntries(shuhoEntries)
func printAllEntries(entries []Entry) {
	for index, entry := range entries {
		fmt.Printf("%d: %s\n", index, entry.String())
	}
}

func ensureInvoiceEntriesAreInShuho(entries []Entry) {

	for _, entry := range entries {
		entryDate := getDate(entry.Date())
		fmt.Println(entryDate)
	}
}

func getDate(eDate string) time.Time {
	entryDate, err := time.Parse("01-02-06", eDate)

	if err != nil {
		entryDate, err = time.Parse("1/2", eDate)
		if err != nil {
			fmt.Printf("ERROR: Invalid Date %s", eDate)
			os.Exit(2)
		}
	}

	return entryDate
}

func ensureNoDuplicateInvoiceEntries(entries []Entry) {
	var entry Entry
	var copies int

	for _, entry = range entries {
		copies = 0
		for _, nextentry := range entries {
			if entry.signature() == nextentry.signature() {
				copies++
			}
		}
	}

	if copies > 1 {
		fmt.Printf("ERROR: Duplicate entry (%s)\n", entry.signature())
		os.Exit(copies)
	}
}

func sumOfChecks(entries []Entry) int {
	var total int

	for _, entry := range entries {
		if entry.Type() == "英文チェック" {
			total++
		}
	}

	return total
}

func sumOfTranslations(entries []Entry) int {
	var total int

	for _, entry := range entries {
		if entry.Type() == "翻訳" {
			total++
		}
	}

	return total
}

func parseInvoice(f *excelize.File) []Entry {
	entries := make([]Entry, 0, 40)
	var sheetName string

	for _, name := range f.GetSheetList() {
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
		//not a complete row, placeholder in excel file
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

	return checkForEmptyCase(row[1])
}

func checkForEmptyCase(caseField string) bool {
	//check for default casenum "ALP-" or blank casenum
	match, _ := regexp.MatchString(`^(?i)ALP-$`, caseField)

	return match && (caseField == "")
}

// only words for shuho entires x/x format
func checkForValidDate(dateField string) bool {
	match, _ := regexp.MatchString(`^(?i)\d+/\d+$`, dateField)

	return match
}

func parseShuho(f *excelize.File) []Entry {
	entries := make([]Entry, 0, 500)

	for index, name := range f.GetSheetList() {
		//fmt.Println("SHUHO SHEET NAME", index, name)

		//skip the first "template" sheet in the file
		if index == 0 {
			continue
		}

		rows, err := f.Rows(name)
		if err != nil {
			fmt.Println(err)
			return entries
		}

		if rows == nil {
			fmt.Printf("ERROR: Sheet %s (%d) - No Rows", name, index)
			return entries
		}

		for rows.Next() {
			var se ShuhoEntry

			row, err := rows.Columns()
			if err != nil {
				fmt.Println(err)
				return entries
			}

			//no row
			if row == nil || len(row) < 6 {
				continue
			}

			if !checkForValidDate(row[0]) {
				continue
			}

			//check for default casenum "ALP-"
			if checkForEmptyCase(row[1]) {
				continue
			}

			//check that 0, 1, 2, and 6 have a value, and that 3 OR 4 has a wordcount
			if (row[2] == "") || (row[6] == "") {
				continue
			}

			//one of the two wordcounts needs to be present
			if (row[3] == "") && (row[4] == "") {
				continue
			}

			se.SDate = row[0]
			se.SCaseNum = row[1]
			se.SType = row[2]
			se.SCWordCount = row[3]
			se.STWordCount = row[4]
			se.SAuthor = row[5]

			entries = append(entries, se)
		}
	}

	return entries
}
