package reportutil

import (
	"HrtChart/calendardata"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"

	"baliance.com/gooxml/color"
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/schema/soo/wml"
	"github.com/joho/godotenv"
	"github.com/xuri/excelize/v2"
	"gopkg.in/gomail.v2"
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
	// Ensure the filename ends with .doc
	if !strings.HasSuffix(filename, ".doc") {
		filename += ".doc"
	}

	// Create a new document
	doc := document.New()

	// Add a paragraph (title)
	titlePara := doc.AddParagraph()
	titleRun := titlePara.AddRun()
	titleRun.AddText("Hormone Tracking")
	titlePara.SetStyle("Title")

	// Add a table and create a row for headers
	table := doc.AddTable()

	table.Properties().Borders().SetBottom(wml.ST_BorderSingle, color.Black, 2)
	table.Properties().Borders().SetInsideHorizontal(wml.ST_BorderBasicBlackDots, color.Green, 1)
	table.Properties().Borders().SetInsideVertical(wml.ST_BorderSingle, color.Black, 1)
	table.Properties().SetCellSpacingAuto()

	headerRow := table.AddRow()

	// Headers
	headerCells := []string{"Day", "Date", "Hormones", "Amount", "Notes"}
	for _, header := range headerCells {
		cell := headerRow.AddCell()
		cell.AddParagraph().AddRun().AddText(header)
	}

	// ✅ Add actual data rows
	for day := 1; day <= 28; day++ {
		row := table.AddRow()
		amountText := calendardata.GetAmountTextDoc(day)

		// Day
		row.AddCell().AddParagraph().AddRun().AddText(fmt.Sprintf("%d", day))

		// Date
		dateValue := startDate.AddDate(0, 0, day-1).Format("2006-01-02")
		row.AddCell().AddParagraph().AddRun().AddText("\t" + dateValue)

		// Hormones
		hormoneCell := row.AddCell()
		hormoneCell.AddParagraph().AddRun().AddText("\x20\x20Estrogen")
		hormoneCell.AddParagraph().AddRun().AddText("\x20\x20Progesterone")
		hormoneCell.AddParagraph().AddRun().AddText("\x20\x20Testosterone")

		// Amount
		row.AddCell().AddParagraph().AddRun().AddText(amountText)

		// Notes
		row.AddCell().AddParagraph().AddRun().AddText("")
	}

	if err := doc.SaveToFile(filename); err != nil {
		return fmt.Errorf("failed to save Word document %q: %w", filename, err)
	}

	// ✅ Get the absolute file path
	absPath, err := filepath.Abs(filename)
	if err != nil {
		return fmt.Errorf("failed to determine absolute path for %q: %w", filename, err)
	}

	// ✅ Ensure the file is fully written and is not empty
	if err := verifyFileIsReady(absPath); err != nil {
		return err
	}

	// ✅ Send the document via email after ensuring it's fully saved
	if err := sendEmailWithAttachment(absPath, "lenags@gmail.com"); err != nil {
		return fmt.Errorf("failed to send email: %w", err)
	}

	fmt.Println("File successfully created and sent:", absPath)
	return nil
}

// columnName returns the Excel column name (A, B, C, etc.) for a zero-based index.
// This simple version covers columns A–Z only.
func columnName(index int) string {
	return string('A' + rune(index))
}

// ✅ Helper function to verify the file is ready before sending
func verifyFileIsReady(filename string) error {
	// Ensure file exists
	fileInfo, err := os.Stat(filename)
	if os.IsNotExist(err) {
		return fmt.Errorf("file %q does not exist, document saving failed", filename)
	}

	// Ensure file is not empty
	if fileInfo.Size() == 0 {
		return fmt.Errorf("file %q is empty, document writing failed", filename)
	}

	// Ensure file is accessible (try opening it)
	file, err := os.Open(filename)
	if err != nil {
		return fmt.Errorf("file %q is still locked or not closed properly: %w", filename, err)
	}
	defer file.Close()

	return nil
}
func sendEmailWithAttachment(filename, recipient string) error {
	err := godotenv.Load()
	if err != nil {
		log.Fatal("Error loading .env file")
	}
	// Email settings
	smtpHost := "smtp.gmail.com"                   // Replace with your SMTP server
	smtpPort := 587                                // Replace with your SMTP port
	senderEmail := "lenags@gmail.com"              // Your email
	senderPassword := os.Getenv("SENDER_PASSWORD") // Your email password

	m := gomail.NewMessage()
	m.SetHeader("From", senderEmail)
	m.SetHeader("To", recipient)
	m.SetHeader("Subject", "Hormone Tracking Document")
	m.SetBody("text/plain", "Attached is your hormone tracking document.")
	m.Attach(filename)

	d := gomail.NewDialer(smtpHost, smtpPort, senderEmail, senderPassword)

	// Send email
	if err := d.DialAndSend(m); err != nil {
		return fmt.Errorf("failed to send email: %w", err)
	}

	fmt.Println("Email sent successfully to", recipient)
	return nil
}
