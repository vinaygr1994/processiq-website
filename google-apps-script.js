// ═══════════════════════════════════════════════════════════════════
// PROCESS IQ AUTOMATION — Google Sheets Form Handler
// ═══════════════════════════════════════════════════════════════════
// HOW TO SET UP (5 minutes):
//
// 1. Go to: https://script.google.com
// 2. Click "New project"
// 3. Delete all existing code, paste THIS entire file
// 4. Change SHEET_ID below to your Google Sheet ID
// 5. Click "Deploy" → "New deployment"
// 6. Type: Web app
// 7. Execute as: Me
// 8. Who has access: Anyone
// 9. Click "Deploy" → copy the Web App URL
// 10. Paste that URL into index.html where it says PASTE_YOUR_SCRIPT_URL_HERE
// ═══════════════════════════════════════════════════════════════════

// ── REPLACE THIS with your Google Sheet ID ──
// Sheet URL looks like: https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit
const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';
const SHEET_TAB = 'Consultations'; // Tab name — will be created if it doesn't exist
const NOTIFY_EMAIL = 'info@processiqautomation.com'; // Gets email on each submission

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    writeToSheet(data);
    sendNotificationEmail(data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // Health check endpoint
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'Process IQ Form Handler Active' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeToSheet(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_TAB);
  
  // Create sheet with headers if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TAB);
    const headers = [
      'Timestamp', 'First Name', 'Last Name', 'Email',
      'Company', 'Primary Challenge', 'Message', 'Source'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    // Style header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0a7c8c');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    // Set column widths
    sheet.setColumnWidth(1, 160);  // Timestamp
    sheet.setColumnWidth(4, 220);  // Email
    sheet.setColumnWidth(6, 220);  // Challenge
    sheet.setColumnWidth(7, 300);  // Message
  }
  
  // Append the new row
  sheet.appendRow([
    data.timestamp || new Date().toLocaleString(),
    data.firstName || '',
    data.lastName  || '',
    data.email     || '',
    data.company   || '',
    data.challenge || '',
    data.message   || '',
    'processiqautomation.com',
  ]);
}

function sendNotificationEmail(data) {
  const subject = `New Consultation Request — ${data.firstName} ${data.lastName} (${data.company})`;
  const body = `
New consultation request submitted on processiqautomation.com

━━━━━━━━━━━━━━━━━━━━━━━━
CONTACT DETAILS
━━━━━━━━━━━━━━━━━━━━━━━━
Name:       ${data.firstName} ${data.lastName}
Email:      ${data.email}
Company:    ${data.company}

━━━━━━━━━━━━━━━━━━━━━━━━
REQUEST DETAILS
━━━━━━━━━━━━━━━━━━━━━━━━
Challenge:  ${data.challenge}
Message:    ${data.message}

━━━━━━━━━━━━━━━━━━━━━━━━
Submitted:  ${data.timestamp}
━━━━━━━━━━━━━━━━━━━━━━━━

View all submissions in Google Sheets:
https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit
  `;
  
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: subject,
    body: body,
  });
}
