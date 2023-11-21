package main

import (
	"testing"
	"time"
)

func TestThisYearOrLastYear(t *testing.T) {
	// theNow, _ := time.Parse("2006-01-02", "2020-01-29")
	theDate := time.Now().AddDate(0, 0, 8)
	theDate2 := time.Now().AddDate(0, 0, 3)
	theDate3 := time.Now().AddDate(0, 0, -2)
	want := theDate.AddDate(-1, 0, 0) //should change to last year
	want2 := theDate.AddDate(0, 0, 0) //current year
	want3 := theDate.AddDate(0, 0, 0) //current year

	calculatedDate := thisYearOrLastYear(theDate)
	if calculatedDate.Year() != want.Year() {
		t.Fatalf("Date should be last year, got %v, wanted %v", calculatedDate, want)
	}

	calculatedDate = thisYearOrLastYear(theDate2)
	if calculatedDate.Year() != want2.Year() {
		t.Fatalf("Date should be last year, got %v, wanted %v", calculatedDate, want2)
	}

	calculatedDate = thisYearOrLastYear(theDate3)
	if calculatedDate.Year() != want3.Year() {
		t.Fatalf("Date should be last year, got %v, wanted %v", calculatedDate, want3)
	}
}
