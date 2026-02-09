// ============================================================
// IJTA Roll Call - Google Apps Script
// ============================================================
// This script receives attendance data from the IJTA Roll Call
// web app and writes it to the Attendance Google Sheet.
//
// SETUP INSTRUCTIONS:
// 1. Open your Attendance Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Delete any existing code and paste this entire file
// 4. Click the floppy disk icon to save
// 5. Click "Deploy" > "New deployment"
// 6. Click the gear icon next to "Select type" and choose "Web app"
// 7. Set "Execute as" to "Me"
// 8. Set "Who has access" to "Anyone"
// 9. Click "Deploy"
// 10. Authorize the script when prompted
// 11. Copy the Web App URL — you'll need it for the app
// ============================================================

const ATTENDANCE_SHEET_ID = '1ipQEh5KCRywBOin8GM4xjzvGh9iK1YWp8VD9BXGH_YA';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { date, clinic, coaches, players } = data;

    const ss = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID);
    let sheet = ss.getSheetByName('Attendance');

    // Create the Attendance sheet with headers if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Attendance');
      sheet.appendRow(['Date', 'Clinic', 'Coaches', 'Player Name']);

      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#2e7d32');
      headerRange.setFontColor('white');

      // Set column widths
      sheet.setColumnWidth(1, 120);  // Date
      sheet.setColumnWidth(2, 300);  // Clinic
      sheet.setColumnWidth(3, 250);  // Coaches
      sheet.setColumnWidth(4, 200);  // Player Name

      // Freeze header row
      sheet.setFrozenRows(1);
    }

    const coachesStr = coaches.join(', ');

    // Add one row per player
    // Coaches only appear on the first row of each session
    for (let i = 0; i < players.length; i++) {
      const row = [
        date,
        clinic,
        i === 0 ? coachesStr : '',  // Coaches only on first row
        players[i]
      ];
      sheet.appendRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, playersRecorded: players.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// This handles GET requests (for testing)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'IJTA Roll Call API is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}
