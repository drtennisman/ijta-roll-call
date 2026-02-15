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
      sheet.appendRow(['Date', 'Clinic', 'Coaches', 'Player Name', 'Status']);

      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#2e7d32');
      headerRange.setFontColor('white');

      // Set column widths
      sheet.setColumnWidth(1, 120);  // Date
      sheet.setColumnWidth(2, 300);  // Clinic
      sheet.setColumnWidth(3, 250);  // Coaches
      sheet.setColumnWidth(4, 200);  // Player Name
      sheet.setColumnWidth(5, 80);   // Status (M/G)

      // Freeze header row
      sheet.setFrozenRows(1);
    }

    const coachesStr = coaches.join(', ');

    // Players now come as objects: { name: "Last, First", status: "M"|"G" }
    // Add one row per player
    // Coaches only appear on the first row of each session
    for (let i = 0; i < players.length; i++) {
      const player = typeof players[i] === 'string' ? { name: players[i], status: 'M' } : players[i];
      const row = [
        date,
        clinic,
        i === 0 ? coachesStr : '',  // Coaches only on first row
        player.name,
        player.status || 'M'
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

// ============================================================
// MONTHLY BILLING REPORT
// ============================================================
// Generates a billing summary in a separate Google Sheet.
// Run this function manually at the end of each month,
// or set up a monthly trigger (Edit > Triggers).
// ============================================================

const BILLING_SHEET_ID = '1GXysHPQzxIRZnxPPnlnZksL-b7Vc2cIJamcgBR75-oI';

// Pricing lookup tables — total charged for N sessions
// Taken directly from the pricing spreadsheet
const PRICING = {
  'Red Ball (Ages 8 and Under)': {
    M: [0, 15, 30, 45, 60, 75, 90, 90, 105, 120, 135],
    G: [0, 20, 40, 60, 80, 100, 120, 120, 140, 160, 180]
  },
  'Orange Ball (Ages 10 and Under)': {
    M: [0, 15, 30, 45, 60, 75, 90, 90, 105, 120, 135],
    G: [0, 20, 40, 60, 80, 100, 120, 120, 140, 160, 180]
  },
  'Green Ball (Ages 12 and Under)': {
    M: [0, 20, 40, 60, 80, 100, 120, 140, 140, 160, 180],
    G: [0, 25, 50, 75, 100, 125, 150, 175, 175, 200, 225]
  },
  'Middle School Yellow Ball Clinic (Ages 12-14)': {
    M: [0, 25, 50, 75, 100, 125, 150, 175, 175, 200, 225],
    G: [0, 30, 60, 90, 120, 150, 180, 210, 210, 240, 270]
  },
  'High School Yellow Ball Clinic (Ages 14 and Over)': {
    M: [0, 25, 50, 75, 100, 125, 150, 175, 200, 200, 225, 250, 275, 300, 325, 350],
    G: [0, 30, 60, 90, 120, 150, 180, 210, 240, 240, 270, 300, 330, 360, 390, 420]
  },
  'Bruno': {
    M: [0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200],
    G: [0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200]
  }
};

// Per-session rates for sessions beyond lookup table range
const PER_SESSION_RATE = {
  'Red Ball (Ages 8 and Under)':                      { M: 15, G: 20 },
  'Orange Ball (Ages 10 and Under)':                   { M: 15, G: 20 },
  'Green Ball (Ages 12 and Under)':                    { M: 20, G: 25 },
  'Middle School Yellow Ball Clinic (Ages 12-14)':     { M: 25, G: 30 },
  'High School Yellow Ball Clinic (Ages 14 and Over)': { M: 25, G: 30 },
  'Bruno':                                              { M: 20, G: 20 }
};

function getTotalCharge(clinic, status, sessions) {
  const table = PRICING[clinic];
  const rate = PER_SESSION_RATE[clinic];
  if (!table || !rate) return 0;

  const s = status === 'G' ? 'G' : 'M';
  const lookup = table[s];

  if (sessions <= 0) return 0;
  if (sessions < lookup.length) return lookup[sessions];

  // Beyond the table — use last table value + extra sessions at per-session rate
  const lastIndex = lookup.length - 1;
  const extraSessions = sessions - lastIndex;
  return lookup[lastIndex] + (extraSessions * rate[s]);
}

function generateMonthlyBilling(monthOverride, yearOverride) {
  // Determine which month to bill for
  const now = new Date();
  const billingMonth = monthOverride || now.getMonth() + 1; // 1-12
  const billingYear = yearOverride || now.getFullYear();

  const monthName = new Date(billingYear, billingMonth - 1, 1)
    .toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  // Read attendance data
  const ss = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID);
  const sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    Logger.log('No Attendance sheet found');
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('No attendance data found');
    return;
  }

  // Tally sessions per player per clinic for the billing month
  // data[0] is the header row: [Date, Clinic, Coaches, Player Name, Status]
  const playerData = {}; // key: "PlayerName|||Clinic" -> { name, clinic, status, sessions }

  for (let i = 1; i < data.length; i++) {
    const dateVal = data[i][0];
    const clinic = data[i][1];
    const playerName = data[i][3];
    const status = data[i][4] || 'M';

    if (!playerName || !clinic) continue;

    // Parse date — could be string "MM/DD/YYYY" or Date object
    let rowDate;
    if (dateVal instanceof Date) {
      rowDate = dateVal;
    } else {
      const parts = String(dateVal).split('/');
      if (parts.length === 3) {
        rowDate = new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]));
      } else {
        continue; // skip unparseable dates
      }
    }

    if (rowDate.getMonth() + 1 !== billingMonth || rowDate.getFullYear() !== billingYear) {
      continue; // not in billing month
    }

    const key = playerName + '|||' + clinic;
    if (!playerData[key]) {
      playerData[key] = { name: playerName, clinic: clinic, status: status, sessions: 0 };
    }
    playerData[key].sessions++;
  }

  // Build billing rows
  const billingRows = [];
  const lastNames = {}; // track last names for sibling flagging

  for (const key in playerData) {
    const p = playerData[key];
    const total = getTotalCharge(p.clinic, p.status, p.sessions);
    const lastName = p.name.split(',')[0].trim();

    if (!lastNames[lastName]) lastNames[lastName] = [];
    lastNames[lastName].push(p.name);

    billingRows.push({
      name: p.name,
      clinic: p.clinic,
      status: p.status === 'G' ? 'Guest' : 'Member',
      sessions: p.sessions,
      total: total,
      lastName: lastName
    });
  }

  // Determine which last names have multiple players (potential siblings)
  const siblingLastNames = {};
  for (const ln in lastNames) {
    // Get unique player names for this last name
    const uniqueNames = [...new Set(lastNames[ln])];
    if (uniqueNames.length > 1) {
      siblingLastNames[ln] = true;
    }
  }

  // Sort by last name, then clinic
  billingRows.sort((a, b) => {
    const nameCompare = a.name.localeCompare(b.name);
    if (nameCompare !== 0) return nameCompare;
    return a.clinic.localeCompare(b.clinic);
  });

  // Write to billing sheet
  const billingSS = SpreadsheetApp.openById(BILLING_SHEET_ID);
  const tabName = monthName;

  // Delete existing tab for this month if it exists
  let billingSheet = billingSS.getSheetByName(tabName);
  if (billingSheet) {
    billingSS.deleteSheet(billingSheet);
  }
  billingSheet = billingSS.insertSheet(tabName);

  // Header row
  const headers = ['Player Name', 'Clinic', 'Status', 'Sessions', 'Total Charged', 'Sibling Discount Note'];
  billingSheet.appendRow(headers);

  // Format header
  const headerRange = billingSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#2e7d32');
  headerRange.setFontColor('white');

  // Data rows
  for (const row of billingRows) {
    const siblingNote = siblingLastNames[row.lastName]
      ? 'CHECK FOR SIBLING DISCOUNT'
      : '';

    billingSheet.appendRow([
      row.name,
      row.clinic,
      row.status,
      row.sessions,
      row.total,
      siblingNote
    ]);
  }

  // Format total column as currency
  if (billingRows.length > 0) {
    const totalRange = billingSheet.getRange(2, 5, billingRows.length, 1);
    totalRange.setNumberFormat('$#,##0.00');

    // Highlight sibling discount rows in yellow
    for (let i = 0; i < billingRows.length; i++) {
      if (siblingLastNames[billingRows[i].lastName]) {
        const rowRange = billingSheet.getRange(i + 2, 1, 1, headers.length);
        rowRange.setBackground('#fff9c4');
      }
    }
  }

  // Set column widths
  billingSheet.setColumnWidth(1, 200);  // Player Name
  billingSheet.setColumnWidth(2, 320);  // Clinic
  billingSheet.setColumnWidth(3, 80);   // Status
  billingSheet.setColumnWidth(4, 80);   // Sessions
  billingSheet.setColumnWidth(5, 120);  // Total Charged
  billingSheet.setColumnWidth(6, 250);  // Sibling Note

  // Freeze header
  billingSheet.setFrozenRows(1);

  // Add summary at bottom
  const summaryStartRow = billingRows.length + 3;
  billingSheet.getRange(summaryStartRow, 1).setValue('MONTHLY SUMMARY');
  billingSheet.getRange(summaryStartRow, 1).setFontWeight('bold');
  billingSheet.getRange(summaryStartRow + 1, 1).setValue('Total Players:');
  billingSheet.getRange(summaryStartRow + 1, 2).setValue(billingRows.length);
  billingSheet.getRange(summaryStartRow + 2, 1).setValue('Total Revenue:');

  const totalRevenue = billingRows.reduce((sum, r) => sum + r.total, 0);
  billingSheet.getRange(summaryStartRow + 2, 2).setValue(totalRevenue);
  billingSheet.getRange(summaryStartRow + 2, 2).setNumberFormat('$#,##0.00');

  Logger.log('Billing report generated for ' + monthName + ': ' + billingRows.length + ' line items, $' + totalRevenue + ' total');
}

// Convenience function: generate billing for the current month
function generateCurrentMonthBilling() {
  const now = new Date();
  generateMonthlyBilling(now.getMonth() + 1, now.getFullYear());
}

// Convenience function: generate billing for last month
function generateLastMonthBilling() {
  const now = new Date();
  let month = now.getMonth(); // 0-indexed, so this is "last month"
  let year = now.getFullYear();
  if (month === 0) {
    month = 12;
    year--;
  }
  generateMonthlyBilling(month, year);
}

// ============================================================
// AUTOMATIC MONTHLY TRIGGER
// ============================================================
// Run setupMonthlyTrigger() once from the Apps Script editor.
// It will schedule generateLastMonthBilling to run automatically
// on the 1st of every month between midnight and 1am.
// ============================================================

function setupMonthlyTrigger() {
  // Remove any existing billing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'generateLastMonthBilling') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create a new monthly trigger — runs on the 1st of each month
  ScriptApp.newTrigger('generateLastMonthBilling')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();

  Logger.log('Monthly billing trigger set up — will run on the 1st of each month');
}
