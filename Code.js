// NeuroLoco Automated Onboarding Email script
// University of Michigan Robotics Department
// EK Klinkman
// Version 1.0 10.15.2024
// Version 1.1 10.16.2024 Moved global variables to beginning
// Version 2.0 04.02.2025 Added functions to send email to HR
// Version 2.1 05.14.2025 Commented out old functionality to send onboarding email to new personnel - replaced w/Slackbot
// Version 2.2 05.14.2025 moved to new Apps Script project connected to original google form

///// New stuff
// Global Variables
const FORM_ID = "1k533tLcjujb0P3mJUiPTG5VuWwlKF0ojhaaRrw7pKP0";  // Form ID
const SHEET_ID = "1d5nM5QMJWwudZskRu_OyLfEBrwZsk75yHI5kKpbcN-o";  // Google Sheet ID
const SHEET_NAME = "Form Responses 1";  // Sheet name
const HR_EMAIL = "pricesam@umich.edu";  // HR's email address
const CC_EMAILS = "emilykk@umich.edu, ejrouse@umich.edu";  // CC recipients
const STATUS_COLUMN = 18;  // Column R (send status: "YES")
const TIMESTAMP_COLUMN = 19;  // Column where "Sent on <date>" is logged

// Apps script code
function sendHRNotification() {
  const form = FormApp.openById(FORM_ID);
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const lastResponse = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Extract form fields (adjust column indices as needed); Column A = [0]
  const name = lastResponse[2];  // Assuming Name is in Column C
  const uniqname = lastResponse[3]; // uniqname in Column D
  const startDate = Utilities.formatDate(new Date(lastResponse[9]), Session.getScriptTimeZone(), "MMMM dd, yyyy"); //;  // Start date in Column J
  const endDate = Utilities.formatDate(new Date(lastResponse[10]), Session.getScriptTimeZone(), "MMMM dd, yyyy"); //; // End Date in column K
  const sendToHR = lastResponse[13];  // Column N (Checkbox) CONFIRMED
  const shortcode = lastResponse[14];  // Column 0
  const rateOfPay = lastResponse[15];  // Column P
  const maxHours = lastResponse[16]; // Column Q
  const statusFlag = lastResponse[STATUS_COLUMN - 1];  // Column R: "YES" or blank
  const notificationSent = lastResponse[TIMESTAMP_COLUMN - 1];  // Column ??? (Tracking email status)

  // Only send email if checkbox is "Yes" and timestamp is empty
  if (sendToHR === "Yes" && notificationSent === "") {
    const subject = "New Personnel Hiring - Neurobionics Lab";
    const body = `Hi Samantha,\n\n`
               + `The Neurobionics Lab is hiring a new personnel: ${name} (${uniqname}). They will be starting on ${startDate}, end date: ${endDate}.\n`
               + `Their pay rate will be ${rateOfPay} per hour and they will be paid from ${shortcode} shortcode. The max hours per week is ${maxHours}.\n\n`
               + `Please email emilykk@umich.edu and ejrouse@umich.edu with any questions. Thank you!`;

    // Send the email
    MailApp.sendEmail({
      to: HR_EMAIL,
      cc: CC_EMAILS,
      subject: subject,
      body: body
    });

    // Log timestamp
    sheet.getRange(lastRow, STATUS_COLUMN).setValue("YES");
    sheet.getRange(lastRow, TIMESTAMP_COLUMN).setValue("Sent on " + new Date().toLocaleString());
  }
}