package main

import (
	"fmt"
	"log"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
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

	// Set column width for "Hormones" to ensure "Testosterone" fits on one line
	// "Testosterone" is 12 characters, so let's set width to a bit more than 12
	if err := f.SetColWidth(sheetName, "C", "C", 15); err != nil {
		log.Fatalf("failed to set column width: %v", err)
	}
	wrapStyle1, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			WrapText: true,
			Vertical: "top",
		},
	})
	// For each day (1..28), populate the rows
	for day := 1; day <= 28; day++ {
		row := day + 1 // row 2..29 in the sheet

		// A) Day
		dayCell, _ := excelize.CoordinatesToCellName(1, row)
		if err := f.SetCellValue(sheetName, dayCell, day); err != nil {
			log.Fatalf("failed to set Day: %v", err)
		}

		// B) Date (for example, using January 2024, day = 1..28)
		dateValue := time.Date(2024, time.January, day, 0, 0, 0, 0, time.UTC)
		dateCell, _ := excelize.CoordinatesToCellName(2, row)
		if err := f.SetCellValue(sheetName, dateCell, dateValue.Format("2006-01-02")); err != nil {
			log.Fatalf("failed to set Date: %v", err)
		}

		// C) Hormones (multiline cell)
		//hormonesText := "Estrogen\nProgesterone\nTestosterone\n"
		hormoneCell, _ := excelize.CoordinatesToCellName(3, row)
		hormoneRuns := []excelize.RichTextRun{
			{
				Text: "Estrogen",
				Font: &excelize.Font{
					Color: "#008000", // green
				},
			},
			{
				Text: "\nProgesterone",
				Font: &excelize.Font{
					Color: "#FFA500", // orange
				},
			},
			{
				Text: "\nTestosterone",
				Font: &excelize.Font{
					Color: "#A020F0", // purple
				},
			},
		}
		// if err := f.SetCellValue(sheetName, hormoneCell, hormonesText); err != nil {
		// 	log.Fatalf("failed to set Hormones: %v", err)
		// }
		if err := f.SetCellRichText(sheetName, hormoneCell, hormoneRuns); err != nil {
			log.Fatalf("failed to set Hormones rich text: %v", err)
		}

		if err := f.SetCellStyle(sheetName, hormoneCell, hormoneCell, wrapStyle1); err != nil {
			log.Fatalf("failed to set style on Hormones cell: %v", err)
		}

		// D) Amount (just an example placeholder)
		amountCell, _ := excelize.CoordinatesToCellName(4, row)

		// If day is from 1 through 5, place "6" and "1" on separate lines
		var amountText string
		switch {
		case day >= 1 && day <= 5:
			// Day 1..5: "6\n1"
			amountText = "6\n\n1"
		case day >= 6 && day <= 8:
			// Day 6..8: "8\n1"
			amountText = "8\n\n1"
		case day >= 9 && day <= 11:
			// Day 6..8: "8\n1"
			amountText = "9\n\n1"
		case day == 12:
			// Day 6..8: "8\n1"
			amountText = "10\n\n1"
		case day == 13:
			// Day 6..8: "8\n1"
			amountText = "4\n\n2"
		case day == 14:
			// Day 6..8: "8\n1"
			amountText = "4\n6\n3"
		case day == 15:
			// Day 6..8: "8\n1"
			amountText = "5\n6\n4"
		case day == 15:
			// Day 6..8: "8\n1"
			amountText = "5\n10\n3"
		default:
			// Day 9..28: blank
			amountText = ""
		}
		if err := f.SetCellValue(sheetName, amountCell, amountText); err != nil {
			log.Fatalf("failed to set Amount: %v", err)
		}
		// Apply wrap text style so multi-line amounts are visible
		if amountText != "" {
			if err := f.SetCellStyle(sheetName, amountCell, amountCell, amountWrapStyle); err != nil {
				log.Fatalf("failed to set style on Amount cell: %v", err)
			}
		}

		// E) Notes (another placeholder)
		notesCell, _ := excelize.CoordinatesToCellName(5, row)
		if err := f.SetCellValue(sheetName, notesCell, ""); err != nil {
			log.Fatalf("failed to set Notes: %v", err)
		}

		// Increase the row height to ensure the multi-line text is visible
		if err := f.SetRowHeight(sheetName, row, 50); err != nil {
			log.Fatalf("failed to set row height: %v", err)
		}
	}
	// Auto-fit columns A..E (optional, though doesn't auto-adjust row height)
	// For a nicer layout, you might prefer setting fixed widths.
	for col := 1; col <= 5; col++ {
		colName, _ := excelize.ColumnNumberToName(col)
		if err := f.AutoFilter(sheetName, fmt.Sprintf("%s1:%s29", colName, colName), nil); err != nil {
			// Using AutoFilter with no criteria won't filter,
			// but sometimes helps with column sizing in some viewers.
			// Another approach is to set column widths manually:
			//f.SetColWidth(sheetName, colName, colName, 45)
		}
	}

	// (Optional) Auto-fit or set other column widths A, B, D, E as needed
	f.SetColWidth(sheetName, "B", "B", 40)
	// e.g., f.SetColWidth(sheetName, "B", "B", 12)

	// Set the active sheet to our sheet
	sheetNameVal, _ := f.GetSheetIndex(sheetName)
	f.SetActiveSheet(sheetNameVal)

	// Save the file
	if err := f.SaveAs("HormonesSchedule.xlsx"); err != nil {
		log.Fatalf("failed to save file: %v", err)
	}

	log.Println("Spreadsheet generated successfully: HormonesSchedule.xlsx")
}
