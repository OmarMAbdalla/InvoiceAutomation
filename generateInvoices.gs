function generateInvoices() {
  const calendarId = "YOUR_GOOGLE_CALENDAR_ID"; //  Replace with your Google Calendar ID
  const sheetId = "YOUR_GOOGLE_SHEET_ID";       //  Replace with your Google Sheet ID

  // Get previous month's date range
  const today = new Date();
  const startDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const endDate = new Date(today.getFullYear(), today.getMonth(), 0);

  const calendar = CalendarApp.getCalendarById(calendarId);
  const events = calendar.getEvents(startDate, endDate);

  // Load client data from the "Clients" sheet
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Clients");
  const data = sheet.getDataRange().getValues();

  const rates = {};
  const emails = {};

  for (let i = 1; i < data.length; i++) {
    const clientKey = data[i][0].toLowerCase().trim();
    emails[clientKey] = data[i][1];
    rates[clientKey] = data[i][2];
  }

  const invoices = {};
  const totals = {};

  // Build invoice records from calendar events
  events.forEach(event => {
    const title = event.getTitle().toLowerCase();
    if (!title.includes("session")) return; // Only match events with "session"

    const namePart = title.split("session")[0].trim();
    const clientNames = namePart.split(/\s+and\s+/i).map(name => name.trim());

    const duration = (event.getEndTime() - event.getStartTime()) / (1000 * 60);
    const hours = duration / 60;

    clientNames.forEach(name => {
      const key = name.toLowerCase();
      const rate = rates[key] || 30; // 30 per hour default if not stated
      let charge = rate * hours;

      // Optional: apply €10 discount per client for group sessions
      if (clientNames.length > 1) {
        charge -= 10;
        if (charge < 0) charge = 0;
      }

      const line = `${event.getStartTime().toDateString()} - ${name}: ${hours.toFixed(2)} hr × €${rate} → €${charge.toFixed(2)}`;

      if (!invoices[key]) {
        invoices[key] = [];
        totals[key] = 0;
      }

      invoices[key].push(line);
      totals[key] += charge;
    });
  });

  const logSheet = getOrCreateLogSheet(sheetId);

  // Create and email invoice for each client
  for (let client in invoices) {
    const displayName = capitalize(client);
    const doc = DocumentApp.create(`Invoice - ${displayName} - ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}`);
    const body = doc.getBody();

    // Build the document
    body.appendParagraph("INVOICE").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`Client: ${displayName}`);
    body.appendParagraph(`Issue Date: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy")}`);
    body.appendParagraph(`\nServices provided between ${startDate.toLocaleDateString()} and ${endDate.toLocaleDateString()}:\n`);



    const table = body.appendTable();
    const header = table.appendTableRow();
    header.appendTableCell("Date").setBold(true);
    header.appendTableCell("Hours").setBold(true);
    header.appendTableCell("Fee (€)").setBold(true);

    // Optional: Set custom column widths
    table.setColumnWidth(0, 220); // Date
    table.setColumnWidth(1, 110); // Hours
    table.setColumnWidth(2, 110); // Fee

invoices[client].forEach(line => {
  const parts = line.split("→ €");
  const date = parts[0].split(" - ")[0].trim();
  const hours = parts[0].match(/([\d.]+)\s*hr/); // e.g. 1.5 hr
  const fee = "€" + Number(parts[1]).toFixed(2);

  const row = table.appendTableRow();
  row.appendTableCell(date).setBold(false);
  row.appendTableCell(hours ? hours[1] : "—").setBold(false);
  row.appendTableCell(fee).setBold(false);
});

const totalRow = table.appendTableRow();
totalRow.appendTableCell("Total:").setBold(true);
totalRow.appendTableCell(""); // Empty cell for hours column
totalRow.appendTableCell(`€${totals[client].toFixed(2)}`).setBold(true);


    body.appendParagraph(`\nBank Details for Transfer:`).setBold(true);
    body.appendParagraph("Bank Holder: NAME").setBold(false);
    body.appendParagraph("BIC: ");
    body.appendParagraph("IBAN: \n");
    body.appendParagraph("").appendText("Or on other payment type (Zelle, Revolut, Paypal) at: ")

    body.appendParagraph("\nPlease make payment within 30 days of this invoice.");
    body.appendParagraph("\nThank you!");
    body.appendParagraph("\nThis invoice was automatically generated.").setFontSize(8).setForegroundColor("#888888");

    // Finalize document
    doc.saveAndClose();
    const pdf = doc.getAs(MimeType.PDF);
    const file = DriveApp.createFile(pdf);
    const fileUrl = file.getUrl();
    const docUrl = doc.getUrl();

    // Send email only if client has a non-zero balance and a valid email
    let emailSent = "No";
    if (emails[client] && totals[client] > 0) {
      const emailBody =
        `Dear ${displayName},\n\n` +
        `Please find attached your invoice for services between ${startDate.toLocaleDateString()} and ${endDate.toLocaleDateString()}.\n\n` +
        `Total due: €${totals[client].toFixed(2)}\n\n` +
        `Thank you!\n\n` +
        `This email was automatically generated.`;

      MailApp.sendEmail({
        to: emails[client],
        subject: `Invoice for ${displayName} – ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}`,
        body: emailBody,
        attachments: [pdf],
        name: "YOUR NAME"
      });

      emailSent = "Yes";
    }

    // Log result
    logSheet.appendRow([
      new Date(),
      displayName,
      totals[client],
      docUrl,
      fileUrl,
      emails[client] || "N/A",
      emailSent
    ]);
  }
}

// Capitalizes the first letter of a name
function capitalize(name) {
  return name.charAt(0).toUpperCase() + name.slice(1).toLowerCase();
}

// Creates a log sheet if it doesn't exist
function getOrCreateLogSheet(sheetId) {
  const ss = SpreadsheetApp.openById(sheetId); 
  let logSheet = ss.getSheetByName("Invoice Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Invoice Log");
    logSheet.appendRow(["Timestamp", "Client", "Total (€)", "Doc URL", "PDF URL", "Email", "Email Sent"]);
  }
  return logSheet;
}
