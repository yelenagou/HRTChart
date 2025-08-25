package main

import (
	"HrtChart/calendardata"
	"HrtChart/reportutil"
	"flag"
	"fmt"
	"log"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	// 1. Parse the "start-day" flag from the command line
	startDayStr := flag.String("start-day", "2024-01-01", "Specify the start day in YYYY-MM-DD format")
	fileNameStr := flag.String("file-name", "hrtschedule.xlsx", "Enter file name and format")
	flag.Parse()
	var fullFileName = fmt.Sprintf("%s%s.%s", *fileNameStr, *startDayStr, "xlsx")

	// 2. Convert the provided string into a time.Time object
	startDate, err := time.Parse("2006-01-02", *startDayStr)
	if err != nil {
		log.Fatalf("Invalid start day format: %v", err)
	}

	// Create a new Excel file
	f := excelize.NewFile()
	sheetName := "Sheet1"
	f.SetSheetName(f.GetSheetName(f.GetActiveSheetIndex()), sheetName)

	// Create headers
	headers := []string{"Day", "Date", "Hormones", "Amount", "Notes"}
	for colIdx, header := range headers {
		cellRef, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
		if err := f.SetCellValue(sheetName, cellRef, header); err != nil {
			log.Fatalf("failed to set header %q: %v", header, err)
		}
	}

	// Define a wrap-text style for the Hormones column
	amountWrapStyle, errWrap := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			WrapText: true,
			Vertical: "top",
		},
	})
	if errWrap != nil {
		log.Fatalf("failed to create wrap-text style: %v", errWrap)
	}

	// Set column width for "Hormones" so "Testosterone" fits on one line
	if err := f.SetColWidth(sheetName, "C", "C", 15); err != nil {
		log.Fatalf("failed to set column width: %v", err)
	}

	// Style for wrapping text in the Hormones column
	wrapStyle1, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			WrapText: true,
			Vertical: "top",
		},
	})

	// 3. Generate rows for Days 1..28, using the user-provided start date
	for day := 1; day <= 28; day++ {
		row := day + 1 // row 2..29 in the sheet

		// A) Day
		dayCell, _ := excelize.CoordinatesToCellName(1, row)
		if err := f.SetCellValue(sheetName, dayCell, day); err != nil {
			log.Fatalf("failed to set Day: %v", err)
		}

		// B) Date - add (day-1) days to the startDate
		dateValue := startDate.AddDate(0, 0, day-1)
		dateCell, _ := excelize.CoordinatesToCellName(2, row)
		if err := f.SetCellValue(sheetName, dateCell, dateValue.Format("2006-01-02")); err != nil {
			log.Fatalf("failed to set Date: %v", err)
		}

		// C) Hormones (rich text)
		hormoneCell, _ := excelize.CoordinatesToCellName(3, row)
		hormoneRuns := []excelize.RichTextRun{
			{
				Text: "\x20Estrogen",
				Font: &excelize.Font{Color: "#008000"}, // green
			},
			{
				Text: "\x20Progesterone",
				Font: &excelize.Font{Color: "#FFA500"}, // orange
			},
			{
				Text: "\x20STestosterone",
				Font: &excelize.Font{Color: "#A020F0"}, // purple
			},
		}
		if err := f.SetCellRichText(sheetName, hormoneCell, hormoneRuns); err != nil {
			log.Fatalf("failed to set Hormones rich text: %v", err)
		}
		if err := f.SetCellStyle(sheetName, hormoneCell, hormoneCell, wrapStyle1); err != nil {
			log.Fatalf("failed to set style on Hormones cell: %v", err)
		}

		// D) Amount
		amountCell, _ := excelize.CoordinatesToCellName(4, row)

		var amountText string

		amountText = calendardata.GetAmountText(day)

		if err := f.SetCellValue(sheetName, amountCell, amountText); err != nil {
			log.Fatalf("failed to set Amount: %v", err)
		}
		if amountText != "" {
			if err := f.SetCellStyle(sheetName, amountCell, amountCell, amountWrapStyle); err != nil {
				log.Fatalf("failed to set style on Amount cell: %v", err)
			}
		}

		// E) Notes
		notesCell, _ := excelize.CoordinatesToCellName(5, row)
		if err := f.SetCellValue(sheetName, notesCell, ""); err != nil {
			log.Fatalf("failed to set Notes: %v", err)
		}

		// Increase the row height so multi-line text is visible
		if err := f.SetRowHeight(sheetName, row, 50); err != nil {
			log.Fatalf("failed to set row height: %v", err)
		}
	}

	// Optional: auto-filter or set column widths
	f.SetColWidth(sheetName, "B", "B", 40)

	// Set the active sheet
	sheetIndex, _ := f.GetSheetIndex(sheetName)
	f.SetActiveSheet(sheetIndex)

	// Save the file
	if err := f.SaveAs(fullFileName); err != nil {
		log.Fatalf("failed to save file: %v", err)
	}

	log.Printf("Spreadsheet generated successfully: %s", fullFileName)
	reportutil.CreateHormonesDoc(*fileNameStr, startDate)
}
