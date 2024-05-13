/*
	Go program to verify my Shuho and Invoice excel files for work
	by James Rubingh
	james@wrive.com 2023
*/

package main

import (
	"flag"
	"fmt"
	"math"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
	"golang.org/x/text/language"
	"golang.org/x/text/message"
)

type Color string

const (
	ColorBlack  Color = "\u001b[30m"
	ColorRed    Color = "\u001b[31m"
	ColorGreen  Color = "\u001b[32m"
	ColorYellow Color = "\u001b[33m"
	ColorBlue   Color = "\u001b[34m"
	ColorReset  Color = "\u001b[0m"
)

func colorize(color Color, message string) {
	fmt.Println(string(color), message, string(ColorReset))
}

var invoicesf *bool
var shuhosf *bool
var checksf *bool
var translationsf *bool

// entry signatures are Date, Casenum, Type, Wordcount
type Entry interface {
	signature() string
	String() string
	Type() string
	Date() time.Time
	Rate() string
	WordCount() string
}

type InvoiceEntry struct {
	rowNum     string
	IDate      time.Time
	ICaseNum   string
	IType      string
	IWordCount string
	rate       string
}

// stuct methods

// signature doesn't include dates
func (e InvoiceEntry) signature() string {
	return fmt.Sprintf("%s %s %s", e.ICaseNum, e.IType, e.IWordCount)
}

func (e InvoiceEntry) String() string {
	return fmt.Sprintf("%s, %s, %s, %s, %s, %s", e.rowNum, e.ICaseNum, e.IDate, e.IType, e.IWordCount, e.rate)
}

func (e InvoiceEntry) Rate() string {
	return e.rate
}

func (e InvoiceEntry) Date() time.Time {
	return e.IDate
}

func (e InvoiceEntry) WordCount() string {
	return e.IWordCount
}

func (e InvoiceEntry) Type() string {
	return e.IType
}

type ShuhoEntry struct {
	SDate       time.Time
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
		fmt.Printf("NOTE: %s - %v, %s\n", e.SType, e.SDate, e.SCaseNum)
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

	return fmt.Sprintf("%v, %s, %s, %s, %s", e.SDate, e.SCaseNum, e.SType, wordcount, e.SAuthor)
}

func (e ShuhoEntry) Date() time.Time {
	return e.SDate
}

func (e ShuhoEntry) Rate() string {
	return ""
}

func (e ShuhoEntry) WordCount() string {
	return getShuhoEntryWordCount(e)
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
	invoicesf = flag.Bool("invoices", false, "display every invoice entry")
	shuhosf = flag.Bool("shuhos", false, "display every shuho entry")
	checksf = flag.Bool("checks", false, "display all checks")
	translationsf = flag.Bool("translations", false, "display all translations")

	flag.Parse()

	if flag.NArg() != 2 {
		fmt.Println("\033[1;31mERROR Usage:\033[0m ./verifyshuho [OPTIONS] <Shuho.xlsx> <Invoice.xlsx>")
		fmt.Println("--invoices show all invoice entries")
		fmt.Println("--shuhos show all shuho entries")
		fmt.Println("--translations show all translations")
		fmt.Println("--checks show all checks")
		return
	}

	shuhoFileName := flag.Arg(0)
	invoiceFileName := flag.Arg(1)

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
	fmt.Printf("Total Translations: \033[1;36m%d\033[0m\n", sumOfTranslations(invoiceEntries))
	fmt.Printf("Total Checks: %d\n", sumOfChecks(invoiceEntries))

	fmt.Println("")

	ensureRatesAreCorrect(invoiceEntries)
	ensureNoDuplicateInvoiceEntries(invoiceEntries)
	ensureInvoiceEntriesAreInShuho(shuhoEntries, invoiceEntries)
	ensureShuhoEntriesAreInShuho(shuhoEntries, invoiceEntries)

	p := message.NewPrinter(language.English)

	fmt.Println("")
	ieTotal := roundFloat(sumEntries(invoiceEntries, "翻訳"), 2)
	p.Printf("Total for translations: \t%.2f\n", ieTotal)
	icTotal := roundFloat(sumEntries(invoiceEntries, "英文チェック"), 2)
	p.Printf("Total for Checks:     \t\t%.2f\n", icTotal)
	pretax := icTotal + ieTotal + 81.16
	p.Printf("\033[1;31mPre-T Total: \t\t\t%.2f\033[0m (%.2f /YR)\n", pretax, pretax*12)
	//p.Printf("\033[1;32mAfter-T Total:          \t\t%.0f\033[0m\n", roundFloat((pretax*0.8979)-330, 2))

	if *invoicesf {
		printAllInvoices(invoiceEntries)
	}

	if *shuhosf {
		printAllShuhos(getScopedShuho(shuhoEntries, invoiceEntries))
	}

	if *translationsf {
		printAllTranslations(getScopedShuho(shuhoEntries, invoiceEntries))
	}

	if *checksf {
		printAllChecks(getScopedShuho(shuhoEntries, invoiceEntries))
	}

	//main
}

func printAllChecks(entries []Entry) {
	colorize(ColorGreen, "\n** All Checks: ")
	for index, entry := range entries {
		if entry.Type() == "英文チェック" {
			fmt.Printf("%d: %s\n", index, entry.String())
		}
	}
}

func printAllTranslations(entries []Entry) {
	colorize(ColorGreen, "\n** All Translations: ")
	for index, entry := range entries {
		if entry.Type() == "翻訳" {
			fmt.Printf("%d: %s\n", index, entry.String())
		}
	}
}

func printAllInvoices(entries []Entry) {
	colorize(ColorGreen, "\n** All Invoices: ")
	for index, entry := range entries {
		fmt.Printf("%d: %s\n", index, entry.String())
	}
}

func printAllShuhos(entries []Entry) {
	colorize(ColorGreen, "\n** All Shuhos: ")
	for index, entry := range entries {
		fmt.Printf("%d: %s\n", index, entry.String())
	}
}

func roundFloat(val float64, precision uint) float64 {
	ratio := math.Pow(10, float64(precision))
	return math.Round(val*ratio) / ratio
}

// sum screening by Type() (translation or check)
func sumEntries(ientries []Entry, eType string) float64 {
	var total float64

	for _, ie := range ientries {
		if ie.Type() == eType {
			rate, _ := strconv.ParseFloat(ie.Rate(), 64)
			wordc, err := strconv.ParseFloat(ie.WordCount(), 64)
			if err != nil {
				fmt.Println(err)
			}

			total += wordc * rate
		}
	}

	return total
}

// FOR DEBUG
// printAllEntries(shuhoEntries)
func printAllEntries(entries []Entry) {
	for index, entry := range entries {
		fmt.Printf("%d: %s\n", index, entry.String())
	}
}

func getDate(txtDate string) time.Time {
	entryDate, err := time.Parse("01-02-06", txtDate)

	if err != nil {
		entryDate, err = time.Parse("1/2", txtDate)
		if err != nil {
			fmt.Printf("ERROR: Invalid Date %s", txtDate)
			os.Exit(2)
		}
		entryDate = thisYearOrLastYear(entryDate)
	}

	return entryDate
}

func thisYearOrLastYear(theDate time.Time) time.Time {
	var MyYear int

	//if the month/day is earlier than a week from now, assume it's last year
	if theDate.YearDay() <= time.Now().AddDate(0, 0, 7).YearDay() {
		MyYear = time.Now().Year()
	} else {
		MyYear = time.Now().AddDate(0, 0, 7).Year() - 1 //take the year -1 (use date after add date in case it's late Dec)
	}

	return time.Date(MyYear, theDate.Month(), theDate.Day(), 0, 0, 0, theDate.Nanosecond(), theDate.Location())
}

func ensureRatesAreCorrect(entries []Entry) {
	var entry Entry
	var errors int

	for _, entry = range entries {
		if entry.Rate() == "18" {
			if entry.Type() != "翻訳" {
				errors++
			}
		} else if entry.Rate() == "1.4" {
			if entry.Type() != "英文チェック" {
				errors++
			}
		}
	}

	if errors != 0 {
		fmt.Printf("\033[1;31mERROR:\033[0m Rate is incorrect (Row %s)\n", entry.String())
	} else {
		showCheckSuccess("Invoice rates are correct")
	}
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

	if copies != 1 {
		fmt.Printf("\033[1;31mERROR:\033[0m Duplicate entry (Row %s)\n", entry.String())
	} else {
		showCheckSuccess("No Duplicate Invoice Entries")
	}
}

func ensureInvoiceEntriesAreInShuho(sentries []Entry, ientries []Entry) {
	scopedShuhoEntries := getScopedShuho(sentries, ientries)
	var totalerrors, copies int

	for _, ientry := range ientries {
		copies = 0
		for _, sentry := range scopedShuhoEntries {
			if sentry.signature() == ientry.signature() {
				copies++
			}
		}

		if copies < 1 {
			fmt.Printf("\033[1;31mERROR:\033[0m Invoice Entry Not in Shuho: Row %s\n", ientry.String())
			totalerrors++
		}
	}

	if totalerrors == 0 {
		showCheckSuccess("All Invoice Entries are in the Shuho")
	}
}

func ensureShuhoEntriesAreInShuho(sentries []Entry, ientries []Entry) {
	scopedShuhoEntries := getScopedShuho(sentries, ientries)
	var totalerrors, copies int

	for _, sentry := range scopedShuhoEntries {
		copies = 0

		for _, ientry := range ientries {
			if ientry.signature() == sentry.signature() {
				copies++
			}
		}

		if copies != 1 {
			fmt.Printf("\033[1;31mERROR:\033[0m Shuho Entry Not in Invoice: %s\n", sentry.String())
			totalerrors++
		}
	}

	if totalerrors == 0 {
		showCheckSuccess("All Shuho Entries are in the Invoice")
	}
}

func getScopedShuho(sentries []Entry, ientries []Entry) []Entry {
	var sse []Entry //scoped shuho entries
	startDate := ientries[0].Date()
	endDate := ientries[len(ientries)-1].Date()

	for _, entry := range sentries {
		//if the date is between the start and end dates,
		//but also check for equal to the start end date
		if (entry.Date().After(startDate.AddDate(0, 0, -1))) && (entry.Date().Before(endDate.AddDate(0, 0, 1))) {
			sse = append(sse, entry)
		}
	}

	//fmt.Printf("Length of sse: %d\n", len(sse))

	return sse
}

func showCheckSuccess(message string) {
	fmt.Printf("OKAY... %s\n", message)
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
			ie.IDate = getDate(row[3])
			ie.ICaseNum = strings.ReplaceAll(row[1], ",", "")
			ie.IType = row[2]
			tmp := strings.ReplaceAll(row[4], ",", "")
			ie.IWordCount = strings.ReplaceAll(tmp, " ", "")
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
			fmt.Printf("\033[1;31mERROR:\033[0m Sheet %s (%d) - No Rows", name, index)
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

			se.SDate = getDate(row[0])
			se.SCaseNum = strings.ReplaceAll(row[1], ",", "")
			se.SType = row[2]
			tmp := strings.ReplaceAll(row[3], ",", "")
			se.SCWordCount = strings.ReplaceAll(tmp, " ", "")
			tmp = strings.ReplaceAll(row[4], ",", "")
			se.STWordCount = strings.ReplaceAll(tmp, " ", "")
			se.SAuthor = row[6]

			entries = append(entries, se)
		}
	}

	return entries
}
