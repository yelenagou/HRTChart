package reportutil

import (
	"HrtChart/calendardata"
	"fmt"
	"time"

	"baliance.com/gooxml/color"
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/schema/soo/wml"
	"github.com/xuri/excelize/v2"
)

// CreateHormonesSpreadsheet generates an Excel file with 5 columns:
// Day, Date, Hormones (multiline), Amount, and Notes.
// The days go from 1 to 28, and the "Date" is calculated from startDate for each day.
func CreateHormonesSpreadsheet(filename string, startDate time.Time) error {
	f := excelize.NewFile()
	sheetName := "Sheet1"

	// Create headers in the first row
	headers := []string{"Day", "Date", "Hormones", "Amount", "Notes"}
	for colIndex, header := range headers {
		cell := fmt.Sprintf("%s1", columnName(colIndex))
		if err := f.SetCellValue(sheetName, cell, header); err != nil {
			return fmt.Errorf("failed to set header %q: %w", header, err)
		}
	}

	// Set column width for "Hormones" (3rd column -> "C") so multiline text is visible
	if err := f.SetColWidth(sheetName, "C", "C", 25); err != nil {
		return fmt.Errorf("failed to set column width on 'C': %w", err)
	}

	// Define a style with wrapping turned on
	wrapStyle, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			WrapText: true,
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create wrap style: %w", err)
	}

	// Populate rows: Day 1 to Day 28
	for day := 1; day <= 28; day++ {
		rowNum := day + 1 // data starts on row 2 (row 1 is headers)
		dayCell := fmt.Sprintf("A%d", rowNum)
		dateCell := fmt.Sprintf("B%d", rowNum)
		hormonesCell := fmt.Sprintf("C%d", rowNum)
		amountCell := fmt.Sprintf("D%d", rowNum)
		notesCell := fmt.Sprintf("E%d", rowNum)

		// Day
		if err := f.SetCellValue(sheetName, dayCell, day); err != nil {
			return fmt.Errorf("failed to set day cell (row=%d): %w", rowNum, err)
		}

		// Date: startDate + (day-1)
		dateValue := startDate.AddDate(0, 0, day-1).Format("2006-01-02")
		if err := f.SetCellValue(sheetName, dateCell, dateValue); err != nil {
			return fmt.Errorf("failed to set date cell (row=%d): %w", rowNum, err)
		}

		// Hormones (multiline)
		multiLineHormones := "Estrogen\nProgesterone\nTestosterone"
		if err := f.SetCellValue(sheetName, hormonesCell, multiLineHormones); err != nil {
			return fmt.Errorf("failed to set hormones cell (row=%d): %w", rowNum, err)
		}
		if err := f.SetCellStyle(sheetName, hormonesCell, hormonesCell, wrapStyle); err != nil {
			return fmt.Errorf("failed to set wrap style for cell (row=%d): %w", rowNum, err)
		}

		// Amount (placeholder or empty)
		if err := f.SetCellValue(sheetName, amountCell, ""); err != nil {
			return fmt.Errorf("failed to set amount cell (row=%d): %w", rowNum, err)
		}

		// Notes (placeholder or empty)
		if err := f.SetCellValue(sheetName, notesCell, ""); err != nil {
			return fmt.Errorf("failed to set notes cell (row=%d): %w", rowNum, err)
		}
	}

	// Save the spreadsheet
	if err := f.SaveAs(filename); err != nil {
		return fmt.Errorf("failed to save spreadsheet %q: %w", filename, err)
	}

	return nil
}

// CreateHormonesDoc generates a Word document (DOCX) with 5 columns:
// Day, Date, Hormones (multiline), Amount, and Notes.
// The days go from 1 to 28, and the "Date" is calculated from startDate for each day.
func CreateHormonesDoc(filename string, startDate time.Time) error {
	// Create a new document
	doc := document.New()

	// Add a paragraph (title)
	titlePara := doc.AddParagraph()
	titleRun := titlePara.AddRun()
	titleRun.AddText("Hormone Tracking")
	titlePara.SetStyle("Title")

	// Add a table and create a row for headers
	table := doc.AddTable()

	table.Properties().Borders().SetBottom(
		wml.ST_BorderSingle, // border style
		color.Black,         // border color
		1,                   // border size
	)
	table.Properties().Borders().SetInsideHorizontal(
		wml.ST_BorderBasicBlackDots, // border style
		color.Green,                 // border color
		1,                           // border size
	)
	table.Properties().Borders().SetInsideVertical(
		wml.ST_BorderSingle, // border style
		color.Black,         // border color
		1,                   // border size
	)
	table.Properties().SetCellSpacingAuto()
	headerRow := table.AddRow()

	// Our headers
	headerCells := []string{"Day", "Date", "Hormones", "Amount", "Notes"}
	for _, header := range headerCells {
		cell := headerRow.AddCell()
		cell.AddParagraph().AddRun().AddText(header)
	}

	// Generate 28 rows
	for day := 1; day <= 28; day++ {
		row := table.AddRow()
		
		amountText := calendardata.GetAmountText(day)
		

		// Day
		cellDay := row.AddCell()
		cellDay.AddParagraph().AddRun().AddText(fmt.Sprintf("%d", day))

		// Date: startDate + (day-1)
		dateValue := startDate.AddDate(0, 0, day-1).Format("2006-01-02")
		cellDate := row.AddCell()
		cellDate.AddParagraph().AddRun().AddText(dateValue)

		// Hormones
		cellHormones := row.AddCell()
		cellHormones.AddParagraph().AddRun().AddText("Estrogen")
		cellHormones.AddParagraph().AddRun().AddText("Progesterone")
		cellHormones.AddParagraph().AddRun().AddText("Testosterone")

		// Amount

		cellAmount := row.AddCell()
		cellAmount.AddParagraph().AddRun().AddText(amountText)

		// Notes
		cellNotes := row.AddCell()
		cellNotes.AddParagraph().AddRun().AddText("")
	}

	// Save the document
	if err := doc.SaveToFile(filename); err != nil {
		return fmt.Errorf("failed to save Word document %q: %w", filename, err)
	}
	return nil
}

// columnName returns the Excel column name (A, B, C, etc.) for a zero-based index.
// This simple version covers columns A–Z only.
func columnName(index int) string {
	return string('A' + rune(index))
}
