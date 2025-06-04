# InvoiceAutomation

# Google Calendar + Google Sheets Invoice Automation

This Google Apps Script automates invoicing for service-based professionals using Google Calendar and Google Sheets.

## Features

- Automatically generates invoices based on past month's calendar sessions
- Reads client rates and emails from Google Sheets
- Creates professional Google Docs and PDFs
- Emails invoices to clients with attached PDFs (if email is provided in google sheets)
- Logs every invoice in a backup sheet

## Requirements

- A Google Calendar where each event title includes the client name followed by "session"
- ("session" keyword can be changed, but must be a word on your calendar that shows up on every billable service)
  Example: `Jane Smith session`
- A Google Sheet named `Clients` with the following columns:
  1. `Client` – full name
  2. `Email` – client’s email address
  3. `Rate` – hourly rate in euros/dollars/etc

- A trigger set to run this script on the **1st of each month**
- Depending on the frequency of invoices, this can be changed too

## Setup

1. Go to [script.google.com](https://script.google.com/)
2. Paste the script into a new Apps Script project
3. Replace:
   - `"YOUR_GOOGLE_CALENDAR_ID"` with your calendar’s ID
   - `"YOUR_GOOGLE_SHEET_ID"` with your Google Sheet ID
   - `"YOUR NAME"` with your real/business name
4. In the Apps Script menu, go to **Triggers → Add Trigger**
   - Choose `generateInvoices`
   - Run monthly, on the 1st
5. Authorize the script when prompted

## Optional Features

- Group discounts supported automatically (subtracts €10 per person)
- You can customize currency, template style, and event filtering

## Data Privacy

All invoice content stays within your Google account. No third-party APIs are used.

## License

Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
