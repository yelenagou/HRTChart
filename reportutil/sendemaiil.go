package reportutil

import (
	"fmt"
	"log"
	"os"

	"github.com/joho/godotenv"
	"gopkg.in/gomail.v2"
)

func SendEmailWithAttachment(filename, recipient string) error {
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
