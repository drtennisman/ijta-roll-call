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
const ROSTER_SHEET_ID = '10nb7o9ZJ-fRyTnA2wosGa6OBCTZeEcGAKRAuCY7PZ8E';

/**
 * Convert a date string ("MM/DD/YYYY") into a month tab name (e.g. "March 2026").
 */
function getMonthTabName(dateStr) {
  const parts = dateStr.split('/');
  const month = parseInt(parts[0]);
  const year = parseInt(parts[2]);
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];
  return monthNames[month - 1] + ' ' + year;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { date, clinic, clinicTab, coaches, players, newCoaches } = data;

    const ss = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID);

    // Determine the month tab name from the submitted date (e.g. "March 2026")
    const tabName = getMonthTabName(date);
    let sheet = ss.getSheetByName(tabName);

    // Create the month tab with headers if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
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
    if (data.noAttendees) {
      // No one showed up — record a single row noting that
      sheet.appendRow([date, clinic, coachesStr, 'No Attendees', '']);
    } else {
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
    }

    // Auto-add new players to the master roster sheet
    const added = addNewPlayersToRoster(clinicTab, players);

    // Auto-add new coaches to the Coaches tab
    const coachesAdded = addNewCoachesToRoster(newCoaches || []);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, playersRecorded: players.length, rosterAdded: added, coachesAdded: coachesAdded }))
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
// AUTO-ADD NEW PLAYERS TO MASTER ROSTER
// ============================================================
// When attendance is submitted, any player not already on the
// clinic's roster tab gets appended automatically.
// ============================================================

function addNewPlayersToRoster(clinicTab, players) {
  if (!clinicTab || !players || players.length === 0) return 0;

  try {
    const rosterSS = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    const rosterSheet = rosterSS.getSheetByName(clinicTab);
    if (!rosterSheet) return 0;

    // Read existing roster names (columns A and B: Last Name, First Name)
    const lastRow = rosterSheet.getLastRow();
    const existingNames = new Set();

    if (lastRow > 1) {
      const nameData = rosterSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (const row of nameData) {
        const last = (row[0] || '').toString().trim();
        const first = (row[1] || '').toString().trim();
        if (last || first) {
          // Normalize to "Last, First" for comparison
          const fullName = first ? last + ', ' + first : last;
          existingNames.add(fullName.toLowerCase());
        }
      }
    }

    // Check each submitted player against the roster
    let addedCount = 0;
    for (const p of players) {
      const player = typeof p === 'string' ? { name: p, status: 'M' } : p;
      const name = (player.name || '').trim();
      if (!name) continue;

      if (!existingNames.has(name.toLowerCase())) {
        // Split "Last, First" into separate columns
        const parts = name.split(',');
        const lastName = (parts[0] || '').trim();
        const firstName = (parts[1] || '').trim();
        const status = player.status === 'G' ? 'G' : 'M';

        rosterSheet.appendRow([lastName, firstName, status]);
        existingNames.add(name.toLowerCase());
        addedCount++;
      }
    }

    return addedCount;
  } catch (error) {
    Logger.log('Error adding to roster: ' + error.toString());
    return 0;
  }
}

// ============================================================
// AUTO-ADD NEW COACHES TO ROSTER
// ============================================================
// When attendance is submitted, any coach marked as "Added"
// gets appended to the Coaches tab if not already there.
// ============================================================

function addNewCoachesToRoster(newCoaches) {
  if (!newCoaches || newCoaches.length === 0) return 0;

  try {
    const rosterSS = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    const sheet = rosterSS.getSheetByName('Coaches');
    if (!sheet) return 0;

    // Read existing coach names (column A)
    const lastRow = sheet.getLastRow();
    const existingNames = new Set();

    if (lastRow > 1) {
      const nameData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (const row of nameData) {
        const name = (row[0] || '').toString().trim();
        if (name) {
          existingNames.add(name.toLowerCase());
        }
      }
    }

    let addedCount = 0;
    for (const coach of newCoaches) {
      const name = (coach || '').trim();
      if (!name) continue;

      if (!existingNames.has(name.toLowerCase())) {
        sheet.appendRow([name]);
        existingNames.add(name.toLowerCase());
        addedCount++;
      }
    }

    return addedCount;
  } catch (error) {
    Logger.log('Error adding coaches to roster: ' + error.toString());
    return 0;
  }
}

// ============================================================
// SHARED HELPERS
// ============================================================

const BILLING_SHEET_ID = '1GXysHPQzxIRZnxPPnlnZksL-b7Vc2cIJamcgBR75-oI';

/**
 * Parse a date value from the spreadsheet (Date object or "MM/DD/YYYY" string).
 * Returns a Date object, or null if unparseable.
 */
function parseDate(dateVal) {
  if (dateVal instanceof Date) return dateVal;
  const parts = String(dateVal).split('/');
  if (parts.length === 3) {
    return new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]));
  }
  return null;
}

/**
 * Read attendance rows for a given month/year.
 * Checks the month-specific tab first (e.g. "March 2026"),
 * then falls back to the old "Attendance" tab for historical data.
 * Returns an array of { date, clinic, playerName, status } objects.
 */
function getAttendanceForMonth(billingMonth, billingYear) {
  const ss = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID);
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];
  const monthTabName = monthNames[billingMonth - 1] + ' ' + billingYear;

  // Collect sheets to read from: month-specific tab first, then legacy "Attendance"
  const sheetsToRead = [];
  const monthSheet = ss.getSheetByName(monthTabName);
  if (monthSheet) sheetsToRead.push(monthSheet);
  const legacySheet = ss.getSheetByName('Attendance');
  if (legacySheet) sheetsToRead.push(legacySheet);

  if (sheetsToRead.length === 0) return [];

  const rows = [];
  for (const sheet of sheetsToRead) {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) continue;

    for (let i = 1; i < data.length; i++) {
      const rowDate = parseDate(data[i][0]);
      if (!rowDate) continue;

      const clinic = data[i][1];
      const playerName = data[i][3];
      const status = data[i][4] || 'M';

      if (!playerName || !clinic) continue;
      if (String(playerName).trim() === 'No Attendees') continue;
      if (rowDate.getMonth() + 1 !== billingMonth || rowDate.getFullYear() !== billingYear) continue;

      rows.push({ date: rowDate, clinic: clinic, playerName: playerName, status: status });
    }
  }
  return rows;
}

/**
 * Read coach hourly rates from the "Coaches" tab in the roster spreadsheet.
 * Returns an object: { "Coach Name": hourlyRate, ... }
 */
function getCoachRates() {
  const rosterSS = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  const sheet = rosterSS.getSheetByName('Coaches');
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const rates = {};
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const name = (data[i][0] || '').toString().trim();
    // Strip $ signs, commas, spaces from rate (handles "$25", "$25.00", etc.)
    const rawRate = (data[i][1] || '').toString().replace(/[$,\s]/g, '');
    const rate = parseFloat(rawRate) || 0;
    if (name) {
      rates[name] = rate;
    }
  }
  return rates;
}

/**
 * Read clinic session durations from the "Clinic Config" tab in the roster spreadsheet.
 * Returns an object: { "Clinic Display Name": sessionHours, ... }
 */
function getClinicSessionDurations() {
  const rosterSS = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  const sheet = rosterSS.getSheetByName('Clinic Config');
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const durations = {};
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const clinicName = (data[i][0] || '').toString().trim();
    const rawHours = (data[i][1] || '').toString().replace(/[^0-9.]/g, '');
    const hours = parseFloat(rawHours) || 0;
    if (clinicName) {
      durations[clinicName] = hours;
    }
  }
  return durations;
}

/**
 * Read attendance rows for a given month/year, INCLUDING coaches data.
 * Returns:
 * {
 *   rows: [{ date, clinic, playerName, status }],
 *   sessionCoaches: { "dateStr|||clinic": ["Coach1", "Coach2", ...] }
 * }
 *
 * Coaches appear ONLY on the first row of each date+clinic session group (column C).
 * De-duplicates coaches in case of multiple submissions for same date+clinic.
 */
function getAttendanceWithCoachesForMonth(billingMonth, billingYear) {
  const ss = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID);
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];
  const monthTabName = monthNames[billingMonth - 1] + ' ' + billingYear;

  // Collect sheets to read from: month-specific tab first, then legacy "Attendance"
  const sheetsToRead = [];
  const monthSheet = ss.getSheetByName(monthTabName);
  if (monthSheet) sheetsToRead.push(monthSheet);
  const legacySheet = ss.getSheetByName('Attendance');
  if (legacySheet) sheetsToRead.push(legacySheet);

  if (sheetsToRead.length === 0) return { rows: [], sessionCoaches: {} };

  const rows = [];
  const sessionCoaches = {};

  for (const sheet of sheetsToRead) {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) continue;

    for (let i = 1; i < data.length; i++) {
      const rowDate = parseDate(data[i][0]);
      if (!rowDate) continue;

      const clinic = data[i][1];
      const coachesStr = (data[i][2] || '').toString().trim();
      const playerName = data[i][3];
      const status = data[i][4] || 'M';

      if (!playerName || !clinic) continue;
      if (String(playerName).trim() === 'No Attendees') continue;
      if (rowDate.getMonth() + 1 !== billingMonth || rowDate.getFullYear() !== billingYear) continue;

      rows.push({ date: rowDate, clinic: clinic, playerName: playerName, status: status });

      // Capture coaches for this session (date+clinic combo)
      if (coachesStr) {
        const dateStr = (rowDate.getMonth() + 1) + '/' + rowDate.getDate() + '/' + rowDate.getFullYear();
        const sessionKey = dateStr + '|||' + clinic;
        const newCoaches = coachesStr.split(', ').map(c => c.trim()).filter(c => c);
        if (!sessionCoaches[sessionKey]) {
          sessionCoaches[sessionKey] = [];
        }
        // De-duplicate coaches (handles multiple submissions for same session)
        for (const c of newCoaches) {
          if (!sessionCoaches[sessionKey].includes(c)) {
            sessionCoaches[sessionKey].push(c);
          }
        }
      }
    }
  }

  return { rows, sessionCoaches };
}

// ============================================================
// MONTHLY BILLING REPORT
// ============================================================
// Generates a billing summary in a separate Google Sheet.
// Run this function manually at the end of each month,
// or set up a monthly trigger (Edit > Triggers).
// ============================================================

// Pricing lookup tables — total charged for N sessions
// Taken directly from the pricing spreadsheet
const PRICING = {
  'Red Ball': {
    M: [0, 15, 30, 45, 60, 75, 90, 90, 105, 120, 135],
    G: [0, 20, 40, 60, 80, 100, 120, 120, 140, 160, 180]
  },
  'Orange Ball': {
    M: [0, 15, 30, 45, 60, 75, 90, 90, 105, 120, 135],
    G: [0, 20, 40, 60, 80, 100, 120, 120, 140, 160, 180]
  },
  'Green Ball': {
    M: [0, 20, 40, 60, 80, 100, 120, 140, 140, 160, 180],
    G: [0, 25, 50, 75, 100, 125, 150, 175, 175, 200, 225]
  },
  'MS Yellow Ball': {
    M: [0, 25, 50, 75, 100, 125, 150, 175, 175, 200, 225],
    G: [0, 30, 60, 90, 120, 150, 180, 210, 210, 240, 270]
  },
  'HS Yellow Ball': {
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
  'Red Ball':                                           { M: 15, G: 20 },
  'Orange Ball':                                        { M: 15, G: 20 },
  'Green Ball':                                         { M: 20, G: 25 },
  'MS Yellow Ball':                                     { M: 25, G: 30 },
  'HS Yellow Ball':                                     { M: 25, G: 30 },
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
  const now = new Date();
  const billingMonth = monthOverride || now.getMonth() + 1;
  const billingYear = yearOverride || now.getFullYear();

  const monthName = new Date(billingYear, billingMonth - 1, 1)
    .toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  const attendanceRows = getAttendanceForMonth(billingMonth, billingYear);
  if (attendanceRows.length === 0) {
    Logger.log('No attendance data found for ' + monthName);
    return;
  }

  // Tally sessions per player per clinic
  const playerData = {}; // key: "PlayerName|||Clinic" -> { name, clinic, status, sessions }

  for (const row of attendanceRows) {
    const key = row.playerName + '|||' + row.clinic;
    if (!playerData[key]) {
      playerData[key] = { name: row.playerName, clinic: row.clinic, status: row.status, sessions: 0 };
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
  const tabName = 'Billing Summary - ' + monthName;

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
  billingSheet.setColumnWidth(2, 160);  // Clinic
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

  // Per-clinic revenue breakdown
  const clinicSummary = {};
  for (const row of billingRows) {
    if (!clinicSummary[row.clinic]) {
      clinicSummary[row.clinic] = { players: 0, revenue: 0 };
    }
    clinicSummary[row.clinic].players++;
    clinicSummary[row.clinic].revenue += row.total;
  }

  let clinicRow = summaryStartRow + 4;
  billingSheet.getRange(clinicRow, 1).setValue('REVENUE BY CLINIC');
  billingSheet.getRange(clinicRow, 1).setFontWeight('bold');
  clinicRow++;

  const clinicHeaders = ['Clinic', 'Players', 'Revenue'];
  billingSheet.getRange(clinicRow, 1, 1, clinicHeaders.length).setValues([clinicHeaders]);
  billingSheet.getRange(clinicRow, 1, 1, clinicHeaders.length).setFontWeight('bold');
  clinicRow++;

  const clinicNames = Object.keys(clinicSummary).sort();
  for (const name of clinicNames) {
    const cs = clinicSummary[name];
    billingSheet.getRange(clinicRow, 1).setValue(name);
    billingSheet.getRange(clinicRow, 2).setValue(cs.players);
    billingSheet.getRange(clinicRow, 3).setValue(cs.revenue);
    billingSheet.getRange(clinicRow, 3).setNumberFormat('$#,##0.00');
    clinicRow++;
  }

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
// ATTENDANCE SUMMARY REPORT
// ============================================================
// Generates a separate attendance summary with one tab per clinic.
// Each tab shows date-by-date attendance with player names,
// plus totals and revenue for easy cross-checking before billing.
// ============================================================

function generateAttendanceSummary(monthOverride, yearOverride) {
  const now = new Date();
  const billingMonth = monthOverride || now.getMonth() + 1;
  const billingYear = yearOverride || now.getFullYear();

  const monthName = new Date(billingYear, billingMonth - 1, 1)
    .toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  const attendanceRows = getAttendanceForMonth(billingMonth, billingYear);
  if (attendanceRows.length === 0) {
    Logger.log('No attendance data found for ' + monthName);
    return;
  }

  // Organize data by clinic -> date -> list of player names
  const clinicData = {};

  for (const row of attendanceRows) {
    const dateStr = (row.date.getMonth() + 1) + '/' + row.date.getDate() + '/' + row.date.getFullYear();

    if (!clinicData[row.clinic]) {
      clinicData[row.clinic] = { dates: {}, playerStatus: {}, dateObjects: {} };
    }
    if (!clinicData[row.clinic].dates[dateStr]) {
      clinicData[row.clinic].dates[dateStr] = [];
      clinicData[row.clinic].dateObjects[dateStr] = row.date;
    }
    clinicData[row.clinic].dates[dateStr].push(row.playerName);
    clinicData[row.clinic].playerStatus[row.playerName] = row.status;
  }

  // Write to billing spreadsheet
  const billingSS = SpreadsheetApp.openById(BILLING_SHEET_ID);

  for (const clinic in clinicData) {
    const cd = clinicData[clinic];
    const tabName = clinic + ' - Attendance - ' + monthName;

    // Delete existing tab if it exists
    let clinicSheet = billingSS.getSheetByName(tabName);
    if (clinicSheet) {
      billingSS.deleteSheet(clinicSheet);
    }
    clinicSheet = billingSS.insertSheet(tabName);

    // Title
    clinicSheet.getRange(1, 1).setValue(clinic + ' — Attendance Summary — ' + monthName);
    clinicSheet.getRange(1, 1).setFontWeight('bold');
    clinicSheet.getRange(1, 1).setFontSize(12);

    // Headers
    const headers = ['Date', 'Players Present', 'Player Names'];
    clinicSheet.getRange(3, 1, 1, headers.length).setValues([headers]);
    clinicSheet.getRange(3, 1, 1, headers.length).setFontWeight('bold');
    clinicSheet.getRange(3, 1, 1, headers.length).setBackground('#2e7d32');
    clinicSheet.getRange(3, 1, 1, headers.length).setFontColor('white');

    // Sort dates chronologically
    const sortedDates = Object.keys(cd.dates).sort((a, b) => {
      return cd.dateObjects[a] - cd.dateObjects[b];
    });

    let currentRow = 4;
    let totalCheckIns = 0;

    for (const dateStr of sortedDates) {
      const players = cd.dates[dateStr].sort();
      totalCheckIns += players.length;
      clinicSheet.getRange(currentRow, 1).setValue(dateStr);
      clinicSheet.getRange(currentRow, 2).setValue(players.length);
      clinicSheet.getRange(currentRow, 3).setValue(players.join('; '));
      currentRow++;
    }

    // Summary section
    const uniquePlayers = [...new Set(Object.keys(cd.dates).flatMap(d => cd.dates[d]))];
    currentRow += 1;
    clinicSheet.getRange(currentRow, 1).setValue('SUMMARY');
    clinicSheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow++;
    clinicSheet.getRange(currentRow, 1).setValue('Total Sessions:');
    clinicSheet.getRange(currentRow, 2).setValue(sortedDates.length);
    currentRow++;
    clinicSheet.getRange(currentRow, 1).setValue('Total Check-ins:');
    clinicSheet.getRange(currentRow, 2).setValue(totalCheckIns);
    currentRow++;
    clinicSheet.getRange(currentRow, 1).setValue('Unique Players:');
    clinicSheet.getRange(currentRow, 2).setValue(uniquePlayers.length);
    currentRow++;

    // Calculate revenue for this clinic
    let clinicRevenue = 0;
    const playerSessions = {};
    for (const dateStr of sortedDates) {
      for (const player of cd.dates[dateStr]) {
        playerSessions[player] = (playerSessions[player] || 0) + 1;
      }
    }
    for (const player in playerSessions) {
      const status = cd.playerStatus[player] || 'M';
      clinicRevenue += getTotalCharge(clinic, status, playerSessions[player]);
    }

    clinicSheet.getRange(currentRow, 1).setValue('Total Revenue:');
    clinicSheet.getRange(currentRow, 2).setValue(clinicRevenue);
    clinicSheet.getRange(currentRow, 2).setNumberFormat('$#,##0.00');

    // Set column widths
    clinicSheet.setColumnWidth(1, 120);
    clinicSheet.setColumnWidth(2, 120);
    clinicSheet.setColumnWidth(3, 600);

    // Freeze header rows
    clinicSheet.setFrozenRows(3);

    Logger.log('Attendance summary generated for ' + clinic + ': ' + sortedDates.length + ' dates, ' + uniquePlayers.length + ' unique players, $' + clinicRevenue);
  }
}

function generateCurrentMonthAttendanceSummary() {
  const now = new Date();
  generateAttendanceSummary(now.getMonth() + 1, now.getFullYear());
}

function generateLastMonthAttendanceSummary() {
  const now = new Date();
  let month = now.getMonth();
  let year = now.getFullYear();
  if (month === 0) {
    month = 12;
    year--;
  }
  generateAttendanceSummary(month, year);
}

// ============================================================
// ATTENDANCE & STAFFING (A/S) SUMMARY REPORT
// ============================================================
// Generates per-clinic tabs showing attendance data alongside
// staffing costs and net profit calculations.
// Revenue - Staffing = Net Profit
// ============================================================

function generateAttendanceAndStaffingSummary(monthOverride, yearOverride) {
  const now = new Date();
  const billingMonth = monthOverride || now.getMonth() + 1;
  const billingYear = yearOverride || now.getFullYear();

  const monthName = new Date(billingYear, billingMonth - 1, 1)
    .toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  // Get attendance data WITH coaches
  const { rows: attendanceRows, sessionCoaches } =
    getAttendanceWithCoachesForMonth(billingMonth, billingYear);

  if (attendanceRows.length === 0) {
    Logger.log('No attendance data found for A/S summary: ' + monthName);
    return;
  }

  // Read config data from roster spreadsheet
  const coachRates = getCoachRates();
  const sessionDurations = getClinicSessionDurations();

  // Organize data by clinic
  const clinicData = {};

  for (const row of attendanceRows) {
    const dateStr = (row.date.getMonth() + 1) + '/' + row.date.getDate() + '/' + row.date.getFullYear();

    if (!clinicData[row.clinic]) {
      clinicData[row.clinic] = {
        dates: {},
        dateObjects: {},
        playerStatus: {},
        coachesByDate: {}
      };
    }
    const cd = clinicData[row.clinic];

    if (!cd.dates[dateStr]) {
      cd.dates[dateStr] = [];
      cd.dateObjects[dateStr] = row.date;
    }
    cd.dates[dateStr].push(row.playerName);
    cd.playerStatus[row.playerName] = row.status;

    // Map coaches to this clinic+date
    const sessionKey = dateStr + '|||' + row.clinic;
    if (sessionCoaches[sessionKey]) {
      cd.coachesByDate[dateStr] = sessionCoaches[sessionKey];
    }
  }

  // Write to billing spreadsheet
  const billingSS = SpreadsheetApp.openById(BILLING_SHEET_ID);

  for (const clinic in clinicData) {
    const cd = clinicData[clinic];
    const tabName = clinic + ' - A/S Summary - ' + monthName;

    // Delete existing tab if it exists
    let sheet = billingSS.getSheetByName(tabName);
    if (sheet) {
      billingSS.deleteSheet(sheet);
    }
    sheet = billingSS.insertSheet(tabName);

    const sessionHours = sessionDurations[clinic] || 1;

    // === SECTION 1: ATTENDANCE BY DATE ===
    sheet.getRange(1, 1).setValue(clinic + ' \u2014 Attendance & Staffing Summary \u2014 ' + monthName);
    sheet.getRange(1, 1).setFontWeight('bold');
    sheet.getRange(1, 1).setFontSize(12);

    const attendanceHeaders = ['Date', 'Players', 'Coaches Present', 'Player Names'];
    sheet.getRange(3, 1, 1, attendanceHeaders.length).setValues([attendanceHeaders]);
    sheet.getRange(3, 1, 1, attendanceHeaders.length).setFontWeight('bold');
    sheet.getRange(3, 1, 1, attendanceHeaders.length).setBackground('#2e7d32');
    sheet.getRange(3, 1, 1, attendanceHeaders.length).setFontColor('white');

    // Sort dates chronologically
    const sortedDates = Object.keys(cd.dates).sort((a, b) =>
      cd.dateObjects[a] - cd.dateObjects[b]
    );

    let currentRow = 4;
    let totalCheckIns = 0;

    // Track total sessions and dates each coach worked for this clinic
    const coachSessionCount = {};
    const coachSessionDates = {};

    for (const dateStr of sortedDates) {
      const players = cd.dates[dateStr].sort();
      totalCheckIns += players.length;

      const dateCoaches = cd.coachesByDate[dateStr] || [];
      const coachesDisplay = dateCoaches.join(', ') || '(none recorded)';

      sheet.getRange(currentRow, 1).setValue(dateStr);
      sheet.getRange(currentRow, 2).setValue(players.length);
      sheet.getRange(currentRow, 3).setValue(coachesDisplay);
      sheet.getRange(currentRow, 4).setValue(players.join('; '));

      // Tally coach sessions and track dates
      for (const coach of dateCoaches) {
        coachSessionCount[coach] = (coachSessionCount[coach] || 0) + 1;
        if (!coachSessionDates[coach]) coachSessionDates[coach] = [];
        coachSessionDates[coach].push(dateStr);
      }
      currentRow++;
    }

    // Attendance summary row
    const uniquePlayers = [...new Set(Object.keys(cd.dates).flatMap(d => cd.dates[d]))];
    currentRow++;
    sheet.getRange(currentRow, 1).setValue('Total Sessions: ' + sortedDates.length);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 2).setValue('Check-ins: ' + totalCheckIns);
    sheet.getRange(currentRow, 3).setValue('Unique Players: ' + uniquePlayers.length);
    currentRow += 2;

    // === SECTION 2: REVENUE ===
    sheet.getRange(currentRow, 1).setValue('REVENUE');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 1).setFontSize(11);
    currentRow++;

    const revenueHeaders = ['Player', 'Status', 'Sessions', 'Total Charged'];
    sheet.getRange(currentRow, 1, 1, revenueHeaders.length).setValues([revenueHeaders]);
    sheet.getRange(currentRow, 1, 1, revenueHeaders.length).setFontWeight('bold');
    sheet.getRange(currentRow, 1, 1, revenueHeaders.length).setBackground('#1565c0');
    sheet.getRange(currentRow, 1, 1, revenueHeaders.length).setFontColor('white');
    currentRow++;

    // Calculate per-player revenue
    const playerSessions = {};
    for (const dateStr of sortedDates) {
      for (const player of cd.dates[dateStr]) {
        playerSessions[player] = (playerSessions[player] || 0) + 1;
      }
    }

    let totalRevenue = 0;
    const playerNames = Object.keys(playerSessions).sort();
    const revenueStartRow = currentRow;

    for (const player of playerNames) {
      const status = cd.playerStatus[player] || 'M';
      const sessions = playerSessions[player];
      const charge = getTotalCharge(clinic, status, sessions);
      totalRevenue += charge;

      const statusLabel = status === 'G' ? 'Guest' : status === 'S' ? 'Social' : 'Member';

      sheet.getRange(currentRow, 1).setValue(player);
      sheet.getRange(currentRow, 2).setValue(statusLabel);
      sheet.getRange(currentRow, 3).setValue(sessions);
      sheet.getRange(currentRow, 4).setValue(charge);
      currentRow++;
    }

    // Format charges as currency
    if (playerNames.length > 0) {
      sheet.getRange(revenueStartRow, 4, playerNames.length, 1).setNumberFormat('$#,##0.00');
    }

    // Total revenue row
    sheet.getRange(currentRow, 1).setValue('TOTAL REVENUE');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 4).setValue(totalRevenue);
    sheet.getRange(currentRow, 4).setNumberFormat('$#,##0.00');
    sheet.getRange(currentRow, 4).setFontWeight('bold');
    currentRow += 2;

    // === SECTION 3: STAFFING COSTS ===
    sheet.getRange(currentRow, 1).setValue('STAFFING');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 1).setFontSize(11);
    currentRow++;

    const staffingHeaders = ['Coach', 'Sessions', 'Dates', 'Hours/Session', 'Total Hours', 'Rate ($/hr)', 'Total Cost'];
    sheet.getRange(currentRow, 1, 1, staffingHeaders.length).setValues([staffingHeaders]);
    sheet.getRange(currentRow, 1, 1, staffingHeaders.length).setFontWeight('bold');
    sheet.getRange(currentRow, 1, 1, staffingHeaders.length).setBackground('#e65100');
    sheet.getRange(currentRow, 1, 1, staffingHeaders.length).setFontColor('white');
    currentRow++;

    let totalStaffingCost = 0;
    const coachNames = Object.keys(coachSessionCount).sort();
    const staffingStartRow = currentRow;

    for (const coach of coachNames) {
      const sessions = coachSessionCount[coach];
      const dates = (coachSessionDates[coach] || []).join(', ');
      const rate = coachRates[coach] || 0;
      const totalHours = sessions * sessionHours;
      const cost = totalHours * rate;
      totalStaffingCost += cost;

      sheet.getRange(currentRow, 1).setValue(coach);
      sheet.getRange(currentRow, 2).setValue(sessions);
      sheet.getRange(currentRow, 3).setValue(dates);
      sheet.getRange(currentRow, 4).setValue(sessionHours);
      sheet.getRange(currentRow, 5).setValue(totalHours);
      sheet.getRange(currentRow, 6).setValue(rate);
      sheet.getRange(currentRow, 7).setValue(cost);
      currentRow++;
    }

    // Format currency columns
    if (coachNames.length > 0) {
      sheet.getRange(staffingStartRow, 6, coachNames.length, 1).setNumberFormat('$#,##0.00');
      sheet.getRange(staffingStartRow, 7, coachNames.length, 1).setNumberFormat('$#,##0.00');
    }

    // Total staffing cost row
    sheet.getRange(currentRow, 1).setValue('TOTAL STAFFING COST');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 7).setValue(totalStaffingCost);
    sheet.getRange(currentRow, 7).setNumberFormat('$#,##0.00');
    sheet.getRange(currentRow, 7).setFontWeight('bold');
    currentRow += 2;

    // === SECTION 4: NET PROFIT ===
    const netProfit = totalRevenue - totalStaffingCost;

    sheet.getRange(currentRow, 1).setValue('NET PROFIT');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 1).setFontSize(12);
    sheet.getRange(currentRow, 2).setValue(netProfit);
    sheet.getRange(currentRow, 2).setNumberFormat('$#,##0.00');
    sheet.getRange(currentRow, 2).setFontWeight('bold');
    sheet.getRange(currentRow, 2).setFontSize(12);

    // Color net profit green if positive, red if negative
    if (netProfit >= 0) {
      sheet.getRange(currentRow, 2).setFontColor('#2e7d32');
    } else {
      sheet.getRange(currentRow, 2).setFontColor('#c62828');
    }

    // Set column widths
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 400);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 120);

    // Freeze header rows
    sheet.setFrozenRows(3);

    Logger.log('A/S Summary generated for ' + clinic + ': Revenue=$' + totalRevenue +
      ', Staffing=$' + totalStaffingCost + ', Net=$' + netProfit);
  }
}

function generateCurrentMonthASSummary() {
  const now = new Date();
  generateAttendanceAndStaffingSummary(now.getMonth() + 1, now.getFullYear());
}

function generateLastMonthASSummary() {
  const now = new Date();
  let month = now.getMonth();
  let year = now.getFullYear();
  if (month === 0) {
    month = 12;
    year--;
  }
  generateAttendanceAndStaffingSummary(month, year);
}

// ============================================================
// GENERATE ALL REPORTS (billing + A/S summary)
// ============================================================

function generateAllReports(monthOverride, yearOverride) {
  generateMonthlyBilling(monthOverride, yearOverride);
  generateAttendanceAndStaffingSummary(monthOverride, yearOverride);
}

function generateCurrentMonthAllReports() {
  const now = new Date();
  generateAllReports(now.getMonth() + 1, now.getFullYear());
}

function generateLastMonthAllReports() {
  const now = new Date();
  let month = now.getMonth();
  let year = now.getFullYear();
  if (month === 0) {
    month = 12;
    year--;
  }
  generateAllReports(month, year);
}

// ============================================================
// CUSTOM SHEET MENU
// ============================================================
// Adds an "IJTA Reports" menu to the spreadsheet toolbar.
// This runs automatically when the spreadsheet is opened.
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('IJTA Reports')
    .addItem('Generate This Month — All Reports', 'menuCurrentMonthAll')
    .addItem('Generate Last Month — All Reports', 'menuLastMonthAll')
    .addSeparator()
    .addItem('Generate This Month — Billing Only', 'menuCurrentMonthBilling')
    .addItem('Generate This Month — A/S Summary Only', 'menuCurrentMonthAS')
    .addSeparator()
    .addItem('Generate Last Month — Billing Only', 'menuLastMonthBilling')
    .addItem('Generate Last Month — A/S Summary Only', 'menuLastMonthAS')
    .addToUi();
}

// Menu handler functions (with user-friendly alerts)

function menuCurrentMonthAll() {
  const now = new Date();
  const monthName = now.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  generateCurrentMonthAllReports();
  SpreadsheetApp.getUi().alert('Done! All reports generated for ' + monthName + '.');
}

function menuLastMonthAll() {
  const now = new Date();
  let month = now.getMonth();
  let year = now.getFullYear();
  if (month === 0) { month = 12; year--; } else { month; }
  const monthName = new Date(year, month - 1, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  generateLastMonthAllReports();
  SpreadsheetApp.getUi().alert('Done! All reports generated for ' + monthName + '.');
}

function menuCurrentMonthBilling() {
  const now = new Date();
  const monthName = now.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  generateMonthlyBilling(now.getMonth() + 1, now.getFullYear());
  SpreadsheetApp.getUi().alert('Done! Billing report generated for ' + monthName + '.');
}

function menuCurrentMonthAS() {
  const now = new Date();
  const monthName = now.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  generateAttendanceAndStaffingSummary(now.getMonth() + 1, now.getFullYear());
  SpreadsheetApp.getUi().alert('Done! A/S Summary generated for ' + monthName + '.');
}

function menuLastMonthBilling() {
  const now = new Date();
  let month = now.getMonth();
  let year = now.getFullYear();
  if (month === 0) { month = 12; year--; }
  const monthName = new Date(year, month - 1, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  generateMonthlyBilling(month, year);
  SpreadsheetApp.getUi().alert('Done! Billing report generated for ' + monthName + '.');
}

function menuLastMonthAS() {
  const now = new Date();
  let month = now.getMonth();
  let year = now.getFullYear();
  if (month === 0) { month = 12; year--; }
  const monthName = new Date(year, month - 1, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  generateAttendanceAndStaffingSummary(month, year);
  SpreadsheetApp.getUi().alert('Done! A/S Summary generated for ' + monthName + '.');
}

// ============================================================
// AUTOMATIC MONTHLY TRIGGER
// ============================================================
// Run setupMonthlyTrigger() once from the Apps Script editor.
// It will schedule generateLastMonthAllReports to run automatically
// on the 1st of every month between midnight and 1am.
// ============================================================

function setupMonthlyTrigger() {
  // Remove any existing billing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    const fn = trigger.getHandlerFunction();
    if (fn === 'generateLastMonthBilling' || fn === 'generateLastMonthAllReports') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create a new monthly trigger — runs on the 1st of each month
  ScriptApp.newTrigger('generateLastMonthAllReports')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();

  Logger.log('Monthly trigger set up — billing + A/S summary will run on the 1st of each month');
}
