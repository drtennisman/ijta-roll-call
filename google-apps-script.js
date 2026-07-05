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

// Missing-roll reminders: don't flag/alert on anything before this date
// (the schedule changed for summer, so older "gaps" aren't real misses).
const REMINDER_GO_LIVE = '2026-06-29';
// Live roll app URL — included in alert emails.
const ROLL_APP_URL = 'https://drtennisman.github.io/ijta-roll-call/';

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

    // Clinic cancelled (rain-out / holiday) — record a single marker row.
    // This clears the missing-roll flag for the day; reports skip these rows.
    if (data.cancelled) {
      const reason = (data.cancelReason || 'Other').toString();
      sheet.appendRow([date, clinic, '', 'Clinic Cancelled (' + reason + ')', '']);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, cancelled: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // coaches may be strings (legacy/No Staffing) or { name, hours } objects
    const coachesStr = coaches.map(c =>
      typeof c === 'string' ? c : `${c.name} (${c.hours}h)`
    ).join(', ');

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

// This handles GET requests (health check + the app's missing-roll lookup)
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'missingRolls') {
    const missing = remindersEnabled() ? getCurrentMissingRolls() : [];
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, missing: missing }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'IJTA Roll Call API is running', version: 'families-autofill-v1' }))
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
        const status = player.status === 'G' ? 'G' : player.status === 'S' ? 'S' : 'M';

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
      if (String(playerName).trim().indexOf('Clinic Cancelled') === 0) continue;
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
 * Parse a coaches string into { name, hours } objects.
 * Handles new format "J.C. (1h), Joey (0.5h)" and legacy "J.C., Joey".
 * defaultHours is used for legacy entries without an explicit hours value.
 */
function parseCoachEntries(coachesStr, defaultHours) {
  return coachesStr.split(',').map(entry => {
    entry = entry.trim();
    const match = entry.match(/^(.+?)\s*\(([0-9.]+)h\)$/);
    if (match) {
      return { name: match[1].trim(), hours: parseFloat(match[2]) };
    }
    return { name: entry, hours: defaultHours };
  }).filter(c => c.name);
}

/**
 * Read attendance rows for a given month/year, INCLUDING coaches data.
 * Returns:
 * {
 *   rows: [{ date, clinic, playerName, status }],
 *   sessionCoaches: { "dateStr|||clinic": [{ name, hours }, ...] }
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

  // Clinic session durations — used as the default hours for bare coach
  // names (entries without an explicit "(Xh)" tag, e.g. legacy data).
  const sessionDurations = getClinicSessionDurations();

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
      if (String(playerName).trim().indexOf('Clinic Cancelled') === 0) continue;
      if (rowDate.getMonth() + 1 !== billingMonth || rowDate.getFullYear() !== billingYear) continue;

      rows.push({ date: rowDate, clinic: clinic, playerName: playerName, status: status });

      // Capture coaches for this session (date+clinic combo)
      if (coachesStr) {
        const dateStr = (rowDate.getMonth() + 1) + '/' + rowDate.getDate() + '/' + rowDate.getFullYear();
        const sessionKey = dateStr + '|||' + clinic;
        // Bare names (no "(Xh)" tag) default to this clinic's full session length
        const defaultHours = sessionDurations[clinic] || 1;
        const newCoaches = parseCoachEntries(coachesStr, defaultHours);
        if (!sessionCoaches[sessionKey]) {
          sessionCoaches[sessionKey] = [];
        }
        // De-duplicate coaches by name (handles multiple submissions for same session)
        for (const c of newCoaches) {
          if (!sessionCoaches[sessionKey].some(existing => existing.name === c.name)) {
            sessionCoaches[sessionKey].push(c);
          }
        }
      }
    }
  }

  return { rows, sessionCoaches };
}

// ============================================================
// SIBLING DISCOUNT (10% off every sibling except the highest-priced)
// ============================================================
// Siblings are matched by last name WITHIN each clinic. The self-filling
// "Families" tab in the roster spreadsheet manages the exceptions:
//   Players (Last, First; Last, First) | Siblings? (Yes/No) | Clinic (auto)
// It auto-populates with every detected same-last-name group (Siblings? =
// Yes). Flip a row to "No" to un-link an accidental match (two unrelated
// families sharing a last name), or add a row by hand for real siblings
// with DIFFERENT last names. The highest-priced sibling pays full; the
// rest get 10% off.
// ============================================================

// Clinic roster tab names in the ROSTER spreadsheet (match the app's CLINICS)
const CLINIC_ROSTER_TABS = ['Red Ball', 'Orange Ball', 'Green Ball', 'Middle School', 'High School', 'Bruno'];

function getSiblingOverrides() {
  const result = { notPairs: {}, families: [] };
  const sheet = SpreadsheetApp.openById(ROSTER_SHEET_ID).getSheetByName('Families');
  if (!sheet) return result;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const namesStr = (data[i][0] || '').toString().trim();
    if (!namesStr) continue;
    // Names are "Last, First" so entries are separated by semicolons
    const names = namesStr.split(';').map(s => s.trim().toLowerCase()).filter(Boolean);
    if (names.length < 2) continue;

    const answer = (data[i][1] || '').toString().trim().toLowerCase();
    const isNo = (answer === 'no' || answer === 'n' || answer === 'false');
    if (isNo) {
      // Not siblings — break any auto-match between these names
      for (let a = 0; a < names.length; a++) {
        for (let b = a + 1; b < names.length; b++) {
          result.notPairs[[names[a], names[b]].sort().join('|||')] = true;
        }
      }
    } else {
      // Confirmed siblings (covers cross-last-name families; harmless if same-name)
      result.families.push(names);
    }
  }
  return result;
}

// Groups one clinic's billing rows into families and applies the discount.
// Mutates each row: adds discount, finalTotal, siblingNote, isSibling.
// Returns the total discount given.
function applySiblingDiscounts(rows, overrides) {
  const n = rows.length;
  const parent = [];
  for (let i = 0; i < n; i++) parent.push(i);
  const find = (x) => { while (parent[x] !== x) { parent[x] = parent[parent[x]]; x = parent[x]; } return x; };
  const union = (a, b) => { const ra = find(a), rb = find(b); if (ra !== rb) parent[ra] = rb; };
  const lowerNames = rows.map(r => r.name.toLowerCase());
  const pairKey = (a, b) => [a, b].sort().join('|||');

  // Same last name = same family, unless marked "Not Siblings"
  for (let i = 0; i < n; i++) {
    for (let j = i + 1; j < n; j++) {
      if (rows[i].lastName.toLowerCase() !== rows[j].lastName.toLowerCase()) continue;
      if (overrides.notPairs[pairKey(lowerNames[i], lowerNames[j])]) continue;
      union(i, j);
    }
  }
  // Explicit "Siblings" rows (different last names)
  for (const family of overrides.families) {
    const present = [];
    for (let i = 0; i < n; i++) {
      if (family.indexOf(lowerNames[i]) !== -1) present.push(i);
    }
    for (let k = 1; k < present.length; k++) union(present[0], present[k]);
  }

  const groups = {};
  for (let i = 0; i < n; i++) {
    const root = find(i);
    if (!groups[root]) groups[root] = [];
    groups[root].push(i);
  }

  let totalDiscount = 0;
  rows.forEach(r => { r.discount = 0; r.finalTotal = r.total; r.siblingNote = ''; r.isSibling = false; });

  for (const g in groups) {
    const members = groups[g];
    if (members.length < 2) continue;
    // Highest total pays full (ties broken alphabetically for consistency)
    members.sort((a, b) => rows[b].total - rows[a].total || rows[a].name.localeCompare(rows[b].name));
    rows[members[0]].siblingNote = 'Sibling - full price';
    rows[members[0]].isSibling = true;
    for (let k = 1; k < members.length; k++) {
      const r = rows[members[k]];
      r.discount = Math.round(r.total * 10) / 100;
      r.finalTotal = Math.round((r.total - r.discount) * 100) / 100;
      r.siblingNote = 'SIBLING DISCOUNT (-10%)';
      r.isSibling = true;
      totalDiscount += r.discount;
    }
  }
  return totalDiscount;
}

// Builds discounted billing rows for one clinic from raw player data.
// players: [{ name, status ('M'|'G'|'S'), sessions }]
// Returns { rows, gross, totalDiscount, net } — rows sorted by name.
function buildClinicBillingRows(clinic, players, overrides) {
  const rows = [];
  for (const p of players) {
    const total = getTotalCharge(clinic, p.status, p.sessions);
    rows.push({
      name: p.name,
      status: p.status === 'G' ? 'Guest' : p.status === 'S' ? 'Social' : 'Member',
      sessions: p.sessions,
      total: total,
      lastName: p.name.split(',')[0].trim()
    });
  }
  rows.sort((a, b) => a.name.localeCompare(b.name));
  const totalDiscount = applySiblingDiscounts(rows, overrides);
  let gross = 0, net = 0;
  rows.forEach(r => { gross += r.total; net += r.finalTotal; });
  return { rows: rows, gross: gross, totalDiscount: totalDiscount, net: net };
}

// Self-fills the "Families" tab: scans every clinic roster, finds groups
// of 2+ players sharing a last name, and APPENDS any not already listed
// (Siblings? = "Yes"). Existing rows — and your Yes/No answers — are never
// touched. Runs automatically at billing time and on demand from the menu.
// Returns the number of new families added.
function updateFamiliesList() {
  const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  let sheet = ss.getSheetByName('Families');
  if (!sheet) {
    sheet = ss.insertSheet('Families');
    sheet.appendRow(['Players (siblings share these)', 'Siblings? (Yes/No)', 'Clinic (auto)']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#021f3d').setFontColor('white');
    sheet.setColumnWidth(1, 320);
    sheet.setColumnWidth(2, 130);
    sheet.setColumnWidth(3, 170);
    sheet.setFrozenRows(1);
  }

  // Existing entries, keyed by the sorted lowercase set of names
  const data = sheet.getDataRange().getValues();
  const existing = {};
  for (let i = 1; i < data.length; i++) {
    const names = (data[i][0] || '').toString().split(';')
      .map(s => s.trim().toLowerCase()).filter(Boolean).sort();
    if (names.length) existing[names.join('|||')] = true;
  }

  // Detect candidate families per clinic roster (same last name, 2+ kids)
  const candidates = {}; // key -> { players: [display], clinics: {} }
  for (const tab of CLINIC_ROSTER_TABS) {
    const rs = ss.getSheetByName(tab);
    if (!rs) continue;
    const rd = rs.getDataRange().getValues();
    const byLast = {};
    for (let i = 1; i < rd.length; i++) {
      const last = (rd[i][0] || '').toString().trim();
      const first = (rd[i][1] || '').toString().trim();
      if (!last && !first) continue;
      const lk = last.toLowerCase();
      if (!lk) continue;
      const display = first ? last + ', ' + first : last;
      (byLast[lk] = byLast[lk] || []).push(display);
    }
    for (const lk in byLast) {
      if (byLast[lk].length < 2) continue;
      const players = byLast[lk].slice().sort();
      const key = players.map(p => p.toLowerCase()).join('|||');
      if (!candidates[key]) candidates[key] = { players: players, clinics: {} };
      candidates[key].clinics[tab] = true;
    }
  }

  // Append only genuinely new families
  let added = 0;
  for (const key in candidates) {
    if (existing[key]) continue;
    const c = candidates[key];
    sheet.appendRow([c.players.join('; '), 'Yes', Object.keys(c.clinics).join(', ')]);
    added++;
  }

  // Keep the Yes/No column a simple dropdown
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build();
    sheet.getRange(2, 2, lastRow - 1, 1).setDataValidation(rule);
  }

  Logger.log('Families list updated: ' + added + ' new famil' + (added === 1 ? 'y' : 'ies') + ' added.');
  return added;
}

// Menu handler: refresh the Families list and report what was added.
function menuUpdateFamilies() {
  const added = updateFamiliesList();
  SpreadsheetApp.getUi().alert(added === 0
    ? 'Families list is up to date — no new families found.'
    : 'Added ' + added + ' new famil' + (added === 1 ? 'y' : 'ies') +
      ' to the Families tab (set to "Yes"). Review and flip any to "No" if they are not actually siblings.');
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

  const s = (status === 'G' || status === 'S') ? 'G' : 'M';
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

  // Group sessions by clinic -> player
  const clinicData = {}; // { clinic: { playerKey: { name, status, sessions } } }

  for (const row of attendanceRows) {
    if (!clinicData[row.clinic]) clinicData[row.clinic] = {};
    const cd = clinicData[row.clinic];
    if (!cd[row.playerName]) {
      cd[row.playerName] = { name: row.playerName, status: row.status, sessions: 0 };
    }
    cd[row.playerName].sessions++;
  }

  updateFamiliesList();  // self-fill the Families tab before applying discounts
  const billingSS = SpreadsheetApp.openById(BILLING_SHEET_ID);
  const siblingOverrides = getSiblingOverrides();

  for (const clinic in clinicData) {
    const playerMap = clinicData[clinic];

    // Build billing rows with sibling discounts applied automatically
    const players = [];
    for (const key in playerMap) players.push(playerMap[key]);
    const billing = buildClinicBillingRows(clinic, players, siblingOverrides);
    const billingRows = billing.rows;

    const tabName = clinic + ' - Billing - ' + monthName;
    let sheet = billingSS.getSheetByName(tabName);
    if (sheet) billingSS.deleteSheet(sheet);
    sheet = billingSS.insertSheet(tabName);

    // Header
    const headers = ['Player Name', 'Status', 'Sessions', 'Total', 'Sibling Discount', 'Final Charge', 'Note'];
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#2e7d32');
    headerRange.setFontColor('white');

    // Data rows
    for (const row of billingRows) {
      sheet.appendRow([
        row.name,
        row.status,
        row.sessions,
        row.total,
        row.discount > 0 ? row.discount : '',
        row.finalTotal,
        row.siblingNote
      ]);
    }

    // Format currency and highlight sibling families
    if (billingRows.length > 0) {
      sheet.getRange(2, 4, billingRows.length, 3).setNumberFormat('$#,##0.00');
      for (let i = 0; i < billingRows.length; i++) {
        if (billingRows[i].isSibling) {
          sheet.getRange(i + 2, 1, 1, headers.length).setBackground('#fff9c4');
        }
      }
    }

    // Column widths
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 130);
    sheet.setColumnWidth(6, 110);
    sheet.setColumnWidth(7, 220);
    sheet.setFrozenRows(1);

    // Summary at bottom
    const summaryStartRow = billingRows.length + 3;
    sheet.getRange(summaryStartRow, 1).setValue('SUMMARY');
    sheet.getRange(summaryStartRow, 1).setFontWeight('bold');
    sheet.getRange(summaryStartRow + 1, 1).setValue('Total Players:');
    sheet.getRange(summaryStartRow + 1, 2).setValue(billingRows.length);
    sheet.getRange(summaryStartRow + 2, 1).setValue('Gross Revenue:');
    sheet.getRange(summaryStartRow + 2, 2).setValue(billing.gross);
    sheet.getRange(summaryStartRow + 2, 2).setNumberFormat('$#,##0.00');
    sheet.getRange(summaryStartRow + 3, 1).setValue('Sibling Discounts:');
    sheet.getRange(summaryStartRow + 3, 2).setValue(-billing.totalDiscount);
    sheet.getRange(summaryStartRow + 3, 2).setNumberFormat('$#,##0.00');
    sheet.getRange(summaryStartRow + 4, 1).setValue('Net Revenue:');
    sheet.getRange(summaryStartRow + 4, 2).setValue(billing.net);
    sheet.getRange(summaryStartRow + 4, 2).setNumberFormat('$#,##0.00');
    sheet.getRange(summaryStartRow + 4, 1, 1, 2).setFontWeight('bold');

    Logger.log('Billing report generated for ' + clinic + ': ' + billingRows.length +
      ' players, gross $' + billing.gross + ', discounts $' + billing.totalDiscount +
      ', net $' + billing.net);
  }
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

    // Calculate revenue for this clinic (net of sibling discounts)
    const playerSessions = {};
    for (const dateStr of sortedDates) {
      for (const player of cd.dates[dateStr]) {
        playerSessions[player] = (playerSessions[player] || 0) + 1;
      }
    }
    const revBilling = buildClinicBillingRows(clinic,
      Object.keys(playerSessions).map(p => ({
        name: p,
        status: cd.playerStatus[p] || 'M',
        sessions: playerSessions[p]
      })),
      getSiblingOverrides());
    const clinicRevenue = revBilling.net;

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
  const siblingOverrides = getSiblingOverrides();

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

    // Track total hours and dates each coach worked for this clinic
    const coachTotalHours = {};
    const coachSessionDates = {};

    for (const dateStr of sortedDates) {
      const players = cd.dates[dateStr].sort();
      totalCheckIns += players.length;

      const dateCoaches = cd.coachesByDate[dateStr] || [];
      // Display coach names with hours only if not a full session
      const coachesDisplay = dateCoaches.map(c =>
        c.hours !== 1 ? `${c.name} (${c.hours}h)` : c.name
      ).join(', ') || '(none recorded)';

      sheet.getRange(currentRow, 1).setValue(dateStr);
      sheet.getRange(currentRow, 2).setValue(players.length);
      sheet.getRange(currentRow, 3).setValue(coachesDisplay);
      sheet.getRange(currentRow, 4).setValue(players.join('; '));

      // Tally actual hours and track dates per coach
      for (const coach of dateCoaches) {
        coachTotalHours[coach.name] = (coachTotalHours[coach.name] || 0) + coach.hours;
        if (!coachSessionDates[coach.name]) coachSessionDates[coach.name] = [];
        coachSessionDates[coach.name].push(dateStr);
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

    const revenueHeaders = ['Player', 'Status', 'Sessions', 'Total', 'Sibling Discount', 'Final'];
    sheet.getRange(currentRow, 1, 1, revenueHeaders.length).setValues([revenueHeaders]);
    sheet.getRange(currentRow, 1, 1, revenueHeaders.length).setFontWeight('bold');
    sheet.getRange(currentRow, 1, 1, revenueHeaders.length).setBackground('#1565c0');
    sheet.getRange(currentRow, 1, 1, revenueHeaders.length).setFontColor('white');
    currentRow++;

    // Calculate per-player revenue with sibling discounts applied
    const playerSessions = {};
    for (const dateStr of sortedDates) {
      for (const player of cd.dates[dateStr]) {
        playerSessions[player] = (playerSessions[player] || 0) + 1;
      }
    }

    const revBilling = buildClinicBillingRows(clinic,
      Object.keys(playerSessions).map(p => ({
        name: p,
        status: cd.playerStatus[p] || 'M',
        sessions: playerSessions[p]
      })),
      siblingOverrides);
    const totalRevenue = revBilling.net;
    const revenueStartRow = currentRow;

    for (const r of revBilling.rows) {
      sheet.getRange(currentRow, 1).setValue(r.name);
      sheet.getRange(currentRow, 2).setValue(r.status);
      sheet.getRange(currentRow, 3).setValue(r.sessions);
      sheet.getRange(currentRow, 4).setValue(r.total);
      if (r.discount > 0) sheet.getRange(currentRow, 5).setValue(-r.discount);
      sheet.getRange(currentRow, 6).setValue(r.finalTotal);
      currentRow++;
    }

    // Format charges as currency
    if (revBilling.rows.length > 0) {
      sheet.getRange(revenueStartRow, 4, revBilling.rows.length, 3).setNumberFormat('$#,##0.00');
    }

    // Total revenue row (net of sibling discounts)
    sheet.getRange(currentRow, 1).setValue('TOTAL REVENUE');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 6).setValue(totalRevenue);
    sheet.getRange(currentRow, 6).setNumberFormat('$#,##0.00');
    sheet.getRange(currentRow, 6).setFontWeight('bold');
    currentRow += 2;

    // === SECTION 3: STAFFING COSTS ===
    sheet.getRange(currentRow, 1).setValue('STAFFING');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 1).setFontSize(11);
    currentRow++;

    const staffingHeaders = ['Coach', 'Sessions', 'Dates', 'Total Hours', 'Rate ($/hr)', 'Total Cost'];
    sheet.getRange(currentRow, 1, 1, staffingHeaders.length).setValues([staffingHeaders]);
    sheet.getRange(currentRow, 1, 1, staffingHeaders.length).setFontWeight('bold');
    sheet.getRange(currentRow, 1, 1, staffingHeaders.length).setBackground('#e65100');
    sheet.getRange(currentRow, 1, 1, staffingHeaders.length).setFontColor('white');
    currentRow++;

    let totalStaffingCost = 0;
    const coachNames = Object.keys(coachTotalHours).sort();
    const staffingStartRow = currentRow;

    for (const coach of coachNames) {
      const totalHours = coachTotalHours[coach];
      const sessions = (coachSessionDates[coach] || []).length;
      const dates = (coachSessionDates[coach] || []).join(', ');
      const rate = coachRates[coach] || 0;
      const cost = totalHours * rate;
      totalStaffingCost += cost;

      sheet.getRange(currentRow, 1).setValue(coach);
      sheet.getRange(currentRow, 2).setValue(sessions);
      sheet.getRange(currentRow, 3).setValue(dates);
      sheet.getRange(currentRow, 4).setValue(totalHours);
      sheet.getRange(currentRow, 5).setValue(rate);
      sheet.getRange(currentRow, 6).setValue(cost);
      currentRow++;
    }

    // Format currency columns
    if (coachNames.length > 0) {
      sheet.getRange(staffingStartRow, 5, coachNames.length, 1).setNumberFormat('$#,##0.00');
      sheet.getRange(staffingStartRow, 6, coachNames.length, 1).setNumberFormat('$#,##0.00');
    }

    // Total staffing cost row
    sheet.getRange(currentRow, 1).setValue('TOTAL STAFFING COST');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 6).setValue(totalStaffingCost);
    sheet.getRange(currentRow, 6).setNumberFormat('$#,##0.00');
    sheet.getRange(currentRow, 6).setFontWeight('bold');
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
    sheet.setColumnWidth(6, 120);

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
    .addSeparator()
    .addItem('Update Families List', 'menuUpdateFamilies')
    .addToUi();

  ui.createMenu('Roll Reminders')
    .addItem('Send Me a Test Alert', 'menuTestReminder')
    .addItem('Check for Missing Rolls Now', 'menuCheckMissingNow')
    .addItem('Show What’s Missing', 'menuShowMissing')
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

// ============================================================
// MISSING-ROLL REMINDERS
// ============================================================
// Flags clinics that were scheduled but have no roll logged,
// emails alerts (8pm same day + 7am next morning), and exposes
// the list to the app for an in-app warning badge.
//
// EVERYTHING is managed from tabs in the ROSTER spreadsheet
// (same place as Coaches & Clinic Config) — no code changes needed:
//   "Clinic Schedule"   — Clinic Name | Days | Owner (coach name)
//                         (alerts for a missing roll go to that clinic's
//                          owner, with the email looked up on the Coaches tab)
//   "Coaches"           — add an "Email" column so names resolve to emails
//   "Alert Recipients"  — Name | Email (admins: get EVERY alert; also the
//                         fallback when a clinic has no owner set)
//   "Reminder Settings" — "Reminders On?" | Yes/No  (master switch)
//
// ONE-TIME SETUP (run each once from the editor's Run button):
//   1. setupReminderTabs()         — creates the three tabs, pre-filled
//   2. setupMissingRollTriggers()  — schedules the 8pm + 7am checks
// Then redeploy (New version) so the app can read the missing list.
//
// DAY-TO-DAY: use the "Roll Reminders" menu in the spreadsheet.
// ============================================================

const DAY_ABBR_TO_NUM = { sun: 0, mon: 1, tue: 2, wed: 3, thu: 4, fri: 5, sat: 6 };
const DAY_NUM_TO_NAME = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// ---- Config readers (all from the ROSTER spreadsheet) ----

function remindersEnabled() {
  const sheet = SpreadsheetApp.openById(ROSTER_SHEET_ID).getSheetByName('Reminder Settings');
  if (!sheet) return true; // default ON if the settings tab isn't there yet
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if ((data[i][0] || '').toString().toLowerCase().indexOf('remind') !== -1) {
      const val = (data[i][1] || '').toString().trim().toLowerCase();
      return !(val === 'no' || val === 'off' || val === 'false' || val === '0');
    }
  }
  return true;
}

// Reads the Clinic Schedule tab. Supports an optional "Owner" column
// (located by header name) for targeted alerts, and multiple rows per clinic.
// Returns: { clinicName: { days: [dayNums], owners: [names] } }
function getClinicScheduleDetailed() {
  const sheet = SpreadsheetApp.openById(ROSTER_SHEET_ID).getSheetByName('Clinic Schedule');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  // Locate the optional Owner column by header text
  let ownerCol = -1;
  for (let c = 0; c < data[0].length; c++) {
    if ((data[0][c] || '').toString().toLowerCase().indexOf('owner') !== -1) {
      ownerCol = c;
      break;
    }
  }

  const schedule = {};
  for (let i = 1; i < data.length; i++) {
    const clinic = (data[i][0] || '').toString().trim();
    const daysStr = (data[i][1] || '').toString().toLowerCase();
    if (!clinic) continue;
    const days = [];
    for (const abbr in DAY_ABBR_TO_NUM) {
      if (daysStr.indexOf(abbr) !== -1) days.push(DAY_ABBR_TO_NUM[abbr]);
    }
    if (!days.length) continue;

    const owner = ownerCol >= 0 ? (data[i][ownerCol] || '').toString().trim() : '';

    if (!schedule[clinic]) schedule[clinic] = { days: [], owners: [] };
    days.forEach(d => {
      if (schedule[clinic].days.indexOf(d) === -1) schedule[clinic].days.push(d);
    });
    if (owner && schedule[clinic].owners.indexOf(owner) === -1) {
      schedule[clinic].owners.push(owner);
    }
  }
  return schedule;
}

function getAlertRecipients() {
  const sheet = SpreadsheetApp.openById(ROSTER_SHEET_ID).getSheetByName('Alert Recipients');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const emails = [];
  for (let i = 1; i < data.length; i++) {
    const email = (data[i][1] || '').toString().trim();
    if (email && email.indexOf('@') !== -1) emails.push(email);
  }
  return emails;
}

// Maps coach names to emails from the Coaches tab's "Email" column
// (located by header name). Returns {} if the column doesn't exist yet.
function getCoachEmails() {
  const sheet = SpreadsheetApp.openById(ROSTER_SHEET_ID).getSheetByName('Coaches');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  let emailCol = -1;
  for (let c = 0; c < data[0].length; c++) {
    if ((data[0][c] || '').toString().toLowerCase().indexOf('email') !== -1) {
      emailCol = c;
      break;
    }
  }
  if (emailCol === -1) return {};

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const name = (data[i][0] || '').toString().trim();
    const email = (data[i][emailCol] || '').toString().trim();
    if (name && email.indexOf('@') !== -1) map[name.toLowerCase()] = email;
  }
  return map;
}

// ---- Missing-roll detection ----

// Build a set of "M/D/YYYY|||Clinic" keys for every roll already logged
// in the month tabs spanning [startDate, endDate].
function getLoggedSet(startDate, endDate) {
  const ss = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID);
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];
  const logged = {};
  const cur = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const last = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
  while (cur <= last) {
    const sheet = ss.getSheetByName(monthNames[cur.getMonth()] + ' ' + cur.getFullYear());
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const rd = parseDate(data[i][0]);
        if (!rd) continue;
        const clinic = (data[i][1] || '').toString().trim();
        if (!clinic) continue;
        // ANY row counts as logged — including a "No Attendees" row.
        logged[(rd.getMonth() + 1) + '/' + rd.getDate() + '/' + rd.getFullYear() + '|||' + clinic] = true;
      }
    }
    cur.setMonth(cur.getMonth() + 1);
  }
  return logged;
}

// Returns [{ date: "M/D/YYYY", clinic, day }] for every scheduled
// clinic-day in range with no roll logged. Never reaches before go-live.
function getMissingRolls(startDate, endDate) {
  const schedule = getClinicScheduleDetailed();
  if (Object.keys(schedule).length === 0) return [];

  const goLive = new Date(REMINDER_GO_LIVE + 'T00:00:00');
  const start = new Date(Math.max(startDate.getTime(), goLive.getTime()));
  start.setHours(0, 0, 0, 0);
  const end = new Date(endDate);
  end.setHours(0, 0, 0, 0);
  if (start > end) return [];

  const logged = getLoggedSet(start, end);
  const missing = [];
  const d = new Date(start);
  while (d <= end) {
    const wd = d.getDay();
    const dateStr = (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
    for (const clinic in schedule) {
      if (schedule[clinic].days.indexOf(wd) !== -1 && !logged[dateStr + '|||' + clinic]) {
        missing.push({ date: dateStr, clinic: clinic, day: DAY_NUM_TO_NAME[wd] });
      }
    }
    d.setDate(d.getDate() + 1);
  }
  missing.sort((a, b) => new Date(a.date) - new Date(b.date));
  return missing;
}

// Outstanding missing rolls from go-live through YESTERDAY
// (today's clinics may not have happened yet, so today is excluded).
function getCurrentMissingRolls() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  return getMissingRolls(new Date(REMINDER_GO_LIVE + 'T00:00:00'), yesterday);
}

// ---- Email alerts ----

function sendMissingRollEmail(recipients, missing, isTest) {
  const n = missing.length;
  // Subject must stay plain ASCII — emoji and special dashes get garbled by mail clients
  const subject = (isTest ? '[TEST] ' : '') + 'Missing Roll Alert - ' +
    n + (n !== 1 ? ' clinics need' : ' clinic needs') + ' attendance';

  let html = '<div style="font-family:Arial,sans-serif;color:#333;">';
  html += '<h2 style="color:#c62828;margin-bottom:4px;">&#9888;&#65039; Missing Roll' + (n !== 1 ? 's' : '') + '</h2>';
  html += '<p>These scheduled clinics have <strong>no attendance logged</strong>:</p><ul>';
  for (const m of missing) {
    html += '<li><strong>' + m.day + ' ' + m.date + '</strong> &mdash; ' + m.clinic + '</li>';
  }
  html += '</ul>';
  html += '<p><a href="' + ROLL_APP_URL + '" style="display:inline-block;background:#021f3d;color:#fff;' +
    'padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:bold;">Open the Roll App</a></p>';
  if (isTest) html += '<p style="color:#888;font-size:12px;">This is a test — your reminder system is working.</p>';
  html += '</div>';

  MailApp.sendEmail({ to: recipients.join(','), subject: subject, htmlBody: html });
}

// Core check used by the triggers. includeToday=true for the evening run
// (clinics are done by 8pm); false for the morning run (through yesterday).
function emailMissingRolls(includeToday) {
  if (!remindersEnabled()) return;
  const end = new Date();
  end.setHours(0, 0, 0, 0);
  if (!includeToday) end.setDate(end.getDate() - 1);

  const missing = getMissingRolls(new Date(REMINDER_GO_LIVE + 'T00:00:00'), end);
  if (missing.length === 0) return;
  sendTargetedAlerts(missing, false);
}

// Routes each missing roll to that clinic's OWNER (email resolved via the
// Coaches tab "Email" column), plus everyone on Alert Recipients (admins
// get every alert, and are the safety net when a clinic has no owner set
// or the owner's email can't be resolved). Each person receives ONE email
// listing only their clinics. Returns emails sent.
function sendTargetedAlerts(missing, isTest) {
  const admins = getAlertRecipients();
  const coachEmails = getCoachEmails();
  const schedule = getClinicScheduleDetailed();

  const buckets = {}; // email -> [missing entries]
  const addTo = (email, m) => {
    const key = email.toLowerCase();
    if (!buckets[key]) buckets[key] = [];
    buckets[key].push(m);
  };

  for (const m of missing) {
    const targets = [];
    const info = schedule[m.clinic];
    if (info) {
      info.owners.forEach(name => {
        const em = coachEmails[name.toLowerCase()];
        if (em && targets.indexOf(em) === -1) targets.push(em);
      });
    }
    // Admins always get every alert (also covers clinics with no owner set)
    admins.forEach(a => { if (targets.indexOf(a) === -1) targets.push(a); });

    targets.forEach(em => addTo(em, m));
  }

  let sent = 0;
  for (const email in buckets) {
    sendMissingRollEmail([email], buckets[email], isTest);
    sent++;
  }
  if (sent === 0) {
    Logger.log('Missing rolls found but no recipients resolved — check Alert Recipients and the Coaches Email column.');
  }
  return sent;
}

function checkMissingRollsEvening() { emailMissingRolls(true); }   // 8pm — includes today
function checkMissingRollsMorning() { emailMissingRolls(false); }  // 7am — through yesterday

// ---- One-time setup ----

function setupReminderTabs() {
  const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);

  if (!ss.getSheetByName('Clinic Schedule')) {
    const s = ss.insertSheet('Clinic Schedule');
    s.appendRow(['Clinic Name', 'Days (e.g. Tue, Wed, Thu)']);
    s.appendRow(['Red Ball', 'Wed']);
    s.appendRow(['Orange Ball', 'Wed']);
    s.appendRow(['Green Ball', 'Wed']);
    s.appendRow(['MS Yellow Ball', 'Tue, Wed, Thu']);
    s.appendRow(['HS Yellow Ball', 'Tue, Wed, Thu']);
    s.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#021f3d').setFontColor('white');
    s.setColumnWidth(1, 180); s.setColumnWidth(2, 230); s.setFrozenRows(1);
  }

  if (!ss.getSheetByName('Alert Recipients')) {
    const r = ss.insertSheet('Alert Recipients');
    r.appendRow(['Name', 'Email']);
    r.appendRow(['J.C.', 'jcdfreeman@gmail.com']);
    r.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#021f3d').setFontColor('white');
    r.setColumnWidth(1, 150); r.setColumnWidth(2, 260); r.setFrozenRows(1);
  }

  if (!ss.getSheetByName('Reminder Settings')) {
    const t = ss.insertSheet('Reminder Settings');
    t.appendRow(['Setting', 'Value']);
    t.appendRow(['Reminders On?', 'Yes']);
    t.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#021f3d').setFontColor('white');
    t.setColumnWidth(1, 180); t.setColumnWidth(2, 120); t.setFrozenRows(1);
  }

  Logger.log('Reminder tabs ready in the roster spreadsheet.');
}

// One-time upgrade for targeted alerts: adds an "Owner" column to
// Clinic Schedule and an "Email" column to Coaches (skips any that
// already exist). Fill them in afterward — the owner name must match
// the Coaches tab spelling.
function setupOwnerColumns() {
  const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);

  const sched = ss.getSheetByName('Clinic Schedule');
  if (sched) {
    const headers = sched.getRange(1, 1, 1, sched.getLastColumn()).getValues()[0];
    const hasOwner = headers.some(h => (h || '').toString().toLowerCase().indexOf('owner') !== -1);
    if (!hasOwner) {
      const col = sched.getLastColumn() + 1;
      sched.getRange(1, col).setValue('Owner (coach name)')
        .setFontWeight('bold').setBackground('#021f3d').setFontColor('white');
      sched.setColumnWidth(col, 180);
    }
  }

  const coaches = ss.getSheetByName('Coaches');
  if (coaches) {
    const headers = coaches.getRange(1, 1, 1, coaches.getLastColumn()).getValues()[0];
    const hasEmail = headers.some(h => (h || '').toString().toLowerCase().indexOf('email') !== -1);
    if (!hasEmail) {
      const col = coaches.getLastColumn() + 1;
      coaches.getRange(1, col).setValue('Email').setFontWeight('bold');
      coaches.setColumnWidth(col, 240);
    }
  }

  Logger.log('Owner/Email columns ready — fill them in on the roster spreadsheet.');
}

function setupMissingRollTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    const fn = t.getHandlerFunction();
    if (fn === 'checkMissingRollsEvening' || fn === 'checkMissingRollsMorning') {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger('checkMissingRollsEvening').timeBased().everyDays(1).atHour(20).create();
  ScriptApp.newTrigger('checkMissingRollsMorning').timeBased().everyDays(1).atHour(7).create();
  Logger.log('Missing-roll triggers set: 8pm (evening) + 7am (morning).');
}

// ---- "Roll Reminders" menu handlers ----

function menuTestReminder() {
  const ui = SpreadsheetApp.getUi();
  const recipients = getAlertRecipients();
  if (recipients.length === 0) {
    ui.alert('No recipients yet. Add at least one email to the "Alert Recipients" tab, then try again.');
    return;
  }
  const missing = getCurrentMissingRolls();
  if (missing.length === 0) {
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: '[TEST] Roll reminder system is working',
      htmlBody: '<div style="font-family:Arial,sans-serif;"><h2 style="color:#2e7d32;">&#9989; Test successful</h2>' +
        '<p>Your missing-roll reminder system is set up and can email you. Nothing is currently missing.</p></div>'
    });
  } else {
    sendMissingRollEmail(recipients, missing, true);
  }
  ui.alert('Test alert sent to: ' + recipients.join(', '));
}

function menuCheckMissingNow() {
  const ui = SpreadsheetApp.getUi();
  if (!remindersEnabled()) {
    ui.alert('Reminders are currently OFF. Set "Reminders On?" to Yes in the "Reminder Settings" tab.');
    return;
  }
  const missing = getCurrentMissingRolls();
  if (missing.length === 0) {
    ui.alert('✅ All caught up — no missing rolls.');
    return;
  }
  const sent = sendTargetedAlerts(missing, false);
  if (sent === 0) {
    ui.alert('Found ' + missing.length + ' missing roll(s), but no recipients could be resolved. Check the "Alert Recipients" tab and the Email column on the Coaches tab.');
  } else {
    ui.alert('Sent ' + sent + ' alert email(s) covering ' + missing.length + ' missing roll(s) — each person only gets their own clinics.');
  }
}

function menuShowMissing() {
  const ui = SpreadsheetApp.getUi();
  const missing = getCurrentMissingRolls();
  if (missing.length === 0) {
    ui.alert('✅ All caught up — no missing rolls.');
    return;
  }
  let txt = 'Missing rolls (scheduled but not logged):\n\n';
  for (const m of missing) txt += '• ' + m.day + ' ' + m.date + ' — ' + m.clinic + '\n';
  ui.alert(txt);
}
