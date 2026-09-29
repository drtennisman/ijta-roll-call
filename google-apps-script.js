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
// 11. Copy the Web App URL - you'll need it for the app
// ============================================================

const ATTENDANCE_SHEET_ID = '1ipQEh5KCRywBOin8GM4xjzvGh9iK1YWp8VD9BXGH_YA';
const ROSTER_SHEET_ID = '10nb7o9ZJ-fRyTnA2wosGa6OBCTZeEcGAKRAuCY7PZ8E';

// Where old report tabs get moved by "Archive Old Reports" (see the
// ARCHIVE section near the bottom). Create an empty Google Sheet named
// something like "IJTA Billing Archive", then paste its ID here - it's the
// long string in that sheet's URL between /d/ and /edit. Leave blank until
// you've made one; archiving simply refuses to run without it.
const ARCHIVE_SHEET_ID = '1xoU6HgaYjgvcRQZtz6rpsL2E4A-I6X-moK6_gz60hW4';

// The clinic sign-up Form response sheet - source for parent contact info.
// One row per family: parent name/email/phone plus up to four children.
const SIGNUP_SHEET_ID = '1DszkseXqMekH_erFHVEVELcKRI6UgPLexSZ9YtLYqBc';

// The collections worklist - one row per FAMILY who still owes, with the
// contact details and follow-up history. This is the only sheet the shop
// manager needs. Kept separate from the billing report on purpose: billing
// is regenerated every month, while outreach notes are typed by hand and
// must never be rebuilt out from under her.
const COLLECTIONS_SHEET_ID = '14s6nX2OXuH4cvzj575Ld4GT3JN2Na2aprIyEFkP8Dog';

// Missing-roll reminders: don't flag/alert on anything before this date
// (the schedule changed for summer, so older "gaps" aren't real misses).
const REMINDER_GO_LIVE = '2026-06-29';
// Live roll app URL - included in alert emails.
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
  const lock = LockService.getScriptLock();
  try {
    // Queue simultaneous submissions (e.g. two coaches hitting Submit at
    // the same moment) so each roll's rows land as one contiguous block
    // instead of interleaving. Waits up to 30s for the other to finish.
    lock.waitLock(30000);

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

    // What's already logged for this clinic+date? Guards against double
    // submissions ("Retry" after a timeout whose first attempt actually
    // landed) and re-takes writing duplicate player rows.
    const existingPlayers = {};
    const existingCoachNames = {};   // lowercase name -> true, across the session
    let sessionHasRows = false;
    let sessionFirstRow = -1;        // 1-based sheet row where coaches live
    {
      const existing = sheet.getDataRange().getValues();
      const target = parseDate(date);
      for (let i = 1; i < existing.length; i++) {
        const rd = parseDate(existing[i][0]);
        if (!rd || !target) continue;
        if (rd.getFullYear() !== target.getFullYear() ||
            rd.getMonth() !== target.getMonth() ||
            rd.getDate() !== target.getDate()) continue;
        if ((existing[i][1] || '').toString().trim() !== clinic) continue;
        sessionHasRows = true;
        if (sessionFirstRow === -1) sessionFirstRow = i + 1;
        const pn = (existing[i][3] || '').toString().trim().toLowerCase();
        if (pn) existingPlayers[pn] = true;
        const cs = (existing[i][2] || '').toString().trim();
        if (cs) {
          parseCoachEntries(cs, 1).forEach(c => {
            if (c.name) existingCoachNames[c.name.toLowerCase()] = true;
          });
        }
      }
    }

    // Clinic cancelled (rain-out / holiday) - record a single marker row.
    // This clears the missing-roll flag for the day; reports skip these rows.
    if (data.cancelled) {
      const reason = (data.cancelReason || 'Other').toString();
      // Don't stack a second marker (or contradict rows already logged)
      if (!sessionHasRows) {
        sheet.appendRow([date, clinic, '', 'Clinic Cancelled (' + reason + ')', '']);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, cancelled: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // coaches may be strings (legacy/No Staffing) or { name, hours } objects
    const coachesStr = coaches.map(c =>
      typeof c === 'string' ? sanitizeCoachName(c) : `${sanitizeCoachName(c.name)} (${c.hours}h)`
    ).join(', ');

    // "Add Staffing Only" - adding a coach to a roll already submitted for
    // this date, with no players to record. Refuse loudly if there's no
    // roll to attach them to, rather than accepting and losing the coach.
    if (data.staffingOnly && !sessionHasRows) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'No roll has been submitted for ' + clinic + ' on ' + date +
                 '. Submit the roll with its players first, then add staffing.'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Adding staff to a session already logged (e.g. a coach who was left
    // off a past roll): merge them into that session's coach cell. Without
    // this, a submission whose players are all duplicates would write no
    // rows at all and the coach would be silently lost.
    // "No Staffing" is never merged into a session that already has rows -
    // it would just add a meaningless $0 line beside the real coaches.
    let coachesAddedToSession = [];
    if (sessionHasRows && sessionFirstRow !== -1) {
      const toAdd = coaches
        .filter(c => typeof c !== 'string' && c.name && c.name !== 'No Staffing')
        .map(c => ({ name: sanitizeCoachName(c.name), hours: c.hours }))
        .filter(c => c.name && !existingCoachNames[c.name.toLowerCase()]);
      if (toAdd.length > 0) {
        const cell = sheet.getRange(sessionFirstRow, 3);
        const current = (cell.getValue() || '').toString().trim();
        const addition = toAdd.map(c => `${c.name} (${c.hours}h)`).join(', ');
        // Replace a lone "No Staffing" rather than appending beside real coaches
        const base = (current && current !== 'No Staffing') ? current + ', ' : '';
        cell.setValue(base + addition);
        toAdd.forEach(c => { existingCoachNames[c.name.toLowerCase()] = true; });
        coachesAddedToSession = toAdd.map(c => c.name);
      }
    }

    // Players now come as objects: { name: "Last, First", status: "M"|"G" }
    // Add one row per player - but only players NOT already logged for
    // this clinic+date (duplicates from retries/re-takes are skipped).
    // Coaches only appear on the first row of each written batch.
    let recorded = 0;
    let duplicatesSkipped = 0;
    if (data.noAttendees) {
      // No one showed up - record a single row noting that
      // (skip if this session already has rows - that would contradict them)
      if (!sessionHasRows) {
        sheet.appendRow([date, clinic, coachesStr, 'No Attendees', '']);
      }
    } else {
      const newPlayers = [];
      for (const p of players) {
        const player = typeof p === 'string' ? { name: p, status: 'M' } : p;
        const key = (player.name || '').trim().toLowerCase();
        if (!key) continue;
        if (existingPlayers[key]) { duplicatesSkipped++; continue; }
        existingPlayers[key] = true;  // also catches dupes within one submission
        newPlayers.push(player);
      }
      for (let i = 0; i < newPlayers.length; i++) {
        const player = newPlayers[i];
        // Coaches go on the first row of a NEW session. For an existing
        // session they were merged into its coach cell above, so leave
        // these blank rather than repeating them.
        sheet.appendRow([
          date,
          clinic,
          (!sessionHasRows && i === 0) ? coachesStr : '',
          player.name,
          player.status || 'M'
        ]);
        recorded++;
      }
    }

    // Auto-add new players to the master roster sheet
    const added = addNewPlayersToRoster(clinicTab, players);

    // Auto-add new coaches to the Coaches tab
    const coachesAdded = addNewCoachesToRoster(newCoaches || []);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        playersRecorded: recorded,
        duplicatesSkipped: duplicatesSkipped,
        coachesAddedToSession: coachesAddedToSession,
        addedToExistingSession: sessionHasRows,
        staffingOnly: data.staffingOnly === true,
        rosterAdded: added,
        coachesAdded: coachesAdded
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
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
    .createTextOutput(JSON.stringify({ status: 'IJTA Roll Call API is running', version: 'collections-sheet-v1' }))
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
      const name = sanitizeCoachName(coach);
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
// Coach names live in one comma-separated cell, so a comma inside a name
// would split it into bogus coaches. Convert "Last, First" to "First Last"
// and drop stray commas. Mirrors normalizeCoachName() in the app, so an
// older cached app version can't corrupt the sheet.
function sanitizeCoachName(raw) {
  let name = (raw || '').toString().trim().replace(/\s+/g, ' ');
  if (name.indexOf(',') !== -1) {
    const parts = name.split(',').map(s => s.trim()).filter(String);
    name = (parts.length === 2) ? parts[1] + ' ' + parts[0] : parts.join(' ');
  }
  return name.replace(/,/g, '').replace(/\s+/g, ' ').trim();
}

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

  if (sheetsToRead.length === 0) return { rows: [], sessionCoaches: {}, sessionMarkers: {} };

  // Clinic session durations - used as the default hours for bare coach
  // names (entries without an explicit "(Xh)" tag, e.g. legacy data).
  const sessionDurations = getClinicSessionDurations();

  const rows = [];
  const sessionCoaches = {};
  const sessionMarkers = {}; // "dateStr|||clinic" -> 'No Attendees' | 'Cancelled (Reason)'

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
      if (rowDate.getMonth() + 1 !== billingMonth || rowDate.getFullYear() !== billingYear) continue;

      // Marker rows (session with nobody / cancelled): keep for the A/S
      // session diary, but exclude from revenue and staffing entirely
      const pn = String(playerName).trim();
      if (pn === 'No Attendees' || pn.indexOf('Clinic Cancelled') === 0) {
        const mDateStr = (rowDate.getMonth() + 1) + '/' + rowDate.getDate() + '/' + rowDate.getFullYear();
        sessionMarkers[mDateStr + '|||' + String(clinic).trim()] =
          pn === 'No Attendees' ? 'No Attendees' : pn.replace('Clinic Cancelled', 'Cancelled');
        continue;
      }

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

  return { rows, sessionCoaches, sessionMarkers };
}

// ============================================================
// SIBLING DISCOUNT (10% off every sibling except the highest-priced)
// ============================================================
// Siblings are matched by last name WITHIN each clinic. The self-filling
// "Families" tab in the roster spreadsheet is the source of truth:
//   Players (Last, First; Last, First) | Siblings? (Yes/No) | Clinic (auto)
// It auto-populates with every detected same-last-name group (Siblings? =
// Yes). A "Yes" row IS the family - exactly its members; kids on a Yes row
// are exempt from last-name auto-matching, so you can trim an unrelated
// same-name kid out of a family row and the edit sticks (the auto-fill
// never re-adds a grouping that touches a kid already on the tab).
// Flip a row to "No" to un-link unlisted kids who aren't related, or add
// a row by hand for real siblings with DIFFERENT last names.
// The highest-priced sibling pays full; the rest get 10% off.
// ============================================================

// Clinic roster tab names in the ROSTER spreadsheet (match the app's CLINICS)
const CLINIC_ROSTER_TABS = ['Red Ball', 'Orange Ball', 'Green Ball', 'Middle School', 'High School', 'Bruno'];

function getSiblingOverrides() {
  const result = { notPairs: {}, families: [], familyMembers: {} };
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
      // Not siblings - break any auto-match between these names
      for (let a = 0; a < names.length; a++) {
        for (let b = a + 1; b < names.length; b++) {
          result.notPairs[[names[a], names[b]].sort().join('|||')] = true;
        }
      }
    } else {
      // A "Yes" row IS the family - exactly these members. Kids on a Yes
      // row are claimed: last-name auto-matching leaves them alone, so an
      // unrelated same-name kid (e.g. Lucy Davidson) can't get pulled in.
      result.families.push(names);
      names.forEach(n => { result.familyMembers[n] = true; });
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

  // Same last name = same family - but only between kids NOT claimed by
  // an explicit "Yes" family row, and not vetoed by a "No" row
  for (let i = 0; i < n; i++) {
    for (let j = i + 1; j < n; j++) {
      if (rows[i].lastName.toLowerCase() !== rows[j].lastName.toLowerCase()) continue;
      if (overrides.familyMembers[lowerNames[i]] || overrides.familyMembers[lowerNames[j]]) continue;
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
// Returns { rows, gross, totalDiscount, net } - rows sorted by name.
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
// (Siblings? = "Yes"). Existing rows - and your Yes/No answers - are never
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

  // Existing entries, keyed by the sorted lowercase set of names.
  // Also track every kid mentioned anywhere on the tab - groups touching
  // them are never re-added, so human edits stay as the human left them.
  const data = sheet.getDataRange().getValues();
  const existing = {};
  const mentioned = {};
  for (let i = 1; i < data.length; i++) {
    const names = (data[i][0] || '').toString().split(';')
      .map(s => s.trim().toLowerCase()).filter(Boolean).sort();
    if (names.length) {
      existing[names.join('|||')] = true;
      names.forEach(n => { mentioned[n] = true; });
    }
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

  // Append only genuinely new families - and never a grouping that
  // includes a kid already listed on the tab (e.g. after J.C. trims an
  // unrelated same-name kid out of a family, that edit is permanent)
  let added = 0;
  for (const key in candidates) {
    if (existing[key]) continue;
    const c = candidates[key];
    if (c.players.some(p => mentioned[p.toLowerCase()])) continue;
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
    ? 'Families list is up to date - no new families found.'
    : 'Added ' + added + ' new famil' + (added === 1 ? 'y' : 'ies') +
      ' to the Families tab (set to "Yes"). Review and flip any to "No" if they are not actually siblings.');
}

// ============================================================
// PLAYER CONTACTS
// ============================================================
// A "Contacts" tab in the ROSTER spreadsheet holds one row per player:
//   Player (Last, First) | Parent | Phone | Email
// Player names self-fill from the clinic rosters (like the Families tab),
// so every kid gets a row and missing contacts are visible at a glance.
// Billing tabs then show Parent and Phone beside each player, so the shop
// manager can charge and call from one place.
// ============================================================

function getPlayerContacts() {
  const sheet = SpreadsheetApp.openById(ROSTER_SHEET_ID).getSheetByName('Contacts');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  const h = data[0].map(v => (v || '').toString().toLowerCase());
  const find = (word) => h.findIndex(x => x.indexOf(word) !== -1);
  const pCol = 0;
  const parentCol = find('parent');
  const phoneCol = find('phone');
  const emailCol = find('email');

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const name = (data[i][pCol] || '').toString().trim();
    if (!name) continue;
    map[name.toLowerCase()] = {
      parent: parentCol >= 0 ? (data[i][parentCol] || '').toString().trim() : '',
      phone: phoneCol >= 0 ? (data[i][phoneCol] || '').toString().trim() : '',
      email: emailCol >= 0 ? (data[i][emailCol] || '').toString().trim() : ''
    };
  }
  return map;
}

// Appends a row for every player on a clinic roster who isn't listed yet.
// Never touches existing rows. Returns how many were added.
function updateContactsList() {
  const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  let sheet = ss.getSheetByName('Contacts');
  if (!sheet) {
    sheet = ss.insertSheet('Contacts');
    sheet.appendRow(['Player (Last, First)', 'Parent', 'Phone', 'Email', 'Source']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#021f3d').setFontColor('white');
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 240);
    sheet.setColumnWidth(5, 190);
    sheet.setFrozenRows(1);
  }

  const existing = {};
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const n = (data[i][0] || '').toString().trim().toLowerCase();
    if (n) existing[n] = true;
  }

  const seen = {};
  const toAdd = [];
  for (const tab of CLINIC_ROSTER_TABS) {
    const rs = ss.getSheetByName(tab);
    if (!rs) continue;
    const rd = rs.getDataRange().getValues();
    for (let i = 1; i < rd.length; i++) {
      const last = (rd[i][0] || '').toString().trim();
      const first = (rd[i][1] || '').toString().trim();
      if (!last && !first) continue;
      const display = first ? last + ', ' + first : last;
      const key = display.toLowerCase();
      if (existing[key] || seen[key]) continue;
      seen[key] = true;
      toAdd.push([display, '', '', '', '']);
    }
  }

  if (toAdd.length > 0) {
    toAdd.sort((a, b) => a[0].localeCompare(b[0]));
    sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 5).setValues(toAdd);
  }
  Logger.log('Contacts list updated: ' + toAdd.length + ' player(s) added.');
  return toAdd.length;
}

function menuUpdateContacts() {
  const added = updateContactsList();
  const missing = countContactsMissing();
  SpreadsheetApp.getUi().alert(
    (added === 0 ? 'Contacts list is up to date - no new players.'
                 : 'Added ' + added + ' player' + (added !== 1 ? 's' : '') + ' to the Contacts tab.') +
    (missing > 0 ? '\n\n' + missing + ' player' + (missing !== 1 ? 's have' : ' has') +
                   ' no phone number yet.' : ''));
}

// Reads the sign-up Form responses and fills BLANK contact fields on the
// Contacts tab. Anything typed by hand is never overwritten.
//
// Matching runs in two tiers:
//   1. Exact name - the child's name resolves to the roster's "Last, First"
//   2. Unique surname - no exact match, but exactly one family in the
//      sign-up sheet shares that surname (catches siblings whose own name
//      was never entered). Marked in the Source column so it's reviewable.
// A surname shared by two DIFFERENT families is left alone and reported,
// rather than guessed at.
function syncContactsFromSignup() {
  if (!SIGNUP_SHEET_ID) throw new Error('SIGNUP_SHEET_ID is not set.');

  updateContactsList();   // make sure every roster player has a row

  const norm = (s) => (s || '').toString().toLowerCase().replace(/[^a-z ]/g, '').replace(/\s+/g, ' ').trim();
  const su = SpreadsheetApp.openById(SIGNUP_SHEET_ID).getSheets()[0].getDataRange().getValues();
  if (su.length < 2) return { filled: 0, bySurname: 0, ambiguous: [], missing: [] };

  // Locate columns by header text (the Form's questions are long)
  const head = su[0].map(v => (v || '').toString().toLowerCase());
  let parentCol = -1, emailCol = -1, phoneCol = -1;
  const childCols = [];
  head.forEach((t, i) => {
    if (parentCol === -1 && t.indexOf("parent's name") !== -1) parentCol = i;
    if (emailCol === -1 && t.indexOf('best email') !== -1) emailCol = i;
    if (phoneCol === -1 && t.indexOf('cell phone') !== -1) phoneCol = i;
    if (/^child #\d+ name/.test(t)) childCols.push(i);
  });
  if (parentCol === -1 || childCols.length === 0) {
    throw new Error('Could not find the parent/child columns on the sign-up sheet.');
  }

  const byFull = {}, bySurname = {};
  for (let i = 1; i < su.length; i++) {
    const parent = (su[i][parentCol] || '').toString().trim();
    if (!parent || /^test/i.test(parent)) continue;
    const email = emailCol >= 0 ? (su[i][emailCol] || '').toString().trim() : '';
    const phone = phoneCol >= 0 ? (su[i][phoneCol] || '').toString().trim() : '';
    if (!phone && email.indexOf('@') === -1) continue;   // nothing useful
    const rec = { parent: parent, phone: phone, email: email };
    const parentLast = norm(parent).split(' ').pop();

    for (const col of childCols) {
      const raw = (su[i][col] || '').toString().trim();
      if (!raw || /^(test|clinic)$/i.test(raw)) continue;
      const p = norm(raw).split(' ').filter(String);
      let surname = null;
      if (p.length >= 2) {
        surname = p[p.length - 1];
        const k1 = surname + ' ' + p.slice(0, -1).join(' ');       // last word is surname
        const k2 = p.slice(1).join(' ') + ' ' + p[0];              // first word is given name
        if (!byFull[k1]) byFull[k1] = rec;
        if (!byFull[k2]) byFull[k2] = rec;
      } else if (parentLast) {
        surname = parentLast;                                       // bare first name
        const k = parentLast + ' ' + p[0];
        if (!byFull[k]) byFull[k] = rec;
      }
      if (surname) (bySurname[surname] = bySurname[surname] || []).push(rec);
    }
  }

  // Fill blanks on the Contacts tab
  const sheet = SpreadsheetApp.openById(ROSTER_SHEET_ID).getSheetByName('Contacts');
  const data = sheet.getDataRange().getValues();
  const ambiguous = [], missing = [];
  let filled = 0, viaSurname = 0;

  for (let i = 1; i < data.length; i++) {
    const name = (data[i][0] || '').toString().trim();
    if (!name) continue;
    const hasParent = (data[i][1] || '').toString().trim();
    const hasPhone = (data[i][2] || '').toString().trim();
    const hasEmail = (data[i][3] || '').toString().trim();
    if (hasParent && hasPhone && hasEmail) continue;   // already complete

    const parts = name.split(',');
    const key = norm((parts[0] || '') + ' ' + (parts[1] || ''));
    let rec = byFull[key], source = 'sign-up';

    if (!rec) {
      const fam = bySurname[norm(parts[0])];
      if (fam && fam.length) {
        const uniq = {};
        fam.forEach(f => { uniq[f.parent + '|' + f.phone] = f; });
        const keys = Object.keys(uniq);
        if (keys.length === 1) { rec = uniq[keys[0]]; source = 'sign-up (surname match)'; viaSurname++; }
        else { ambiguous.push(name); continue; }
      }
    }
    if (!rec) { missing.push(name); continue; }

    // Only ever fill blanks - never overwrite what someone typed
    let touched = false;
    if (!hasParent && rec.parent) { sheet.getRange(i + 1, 2).setValue(rec.parent); touched = true; }
    if (!hasPhone && rec.phone) { sheet.getRange(i + 1, 3).setValue(rec.phone); touched = true; }
    if (!hasEmail && rec.email && rec.email.indexOf('@') !== -1) { sheet.getRange(i + 1, 4).setValue(rec.email); touched = true; }
    if (touched) { sheet.getRange(i + 1, 5).setValue(source); filled++; }
  }

  Logger.log('Contacts sync: filled ' + filled + ' (' + viaSurname + ' by surname), ' +
    ambiguous.length + ' ambiguous, ' + missing.length + ' with no record.');
  return { filled: filled, bySurname: viaSurname, ambiguous: ambiguous, missing: missing };
}

function menuSyncContacts() {
  const ui = SpreadsheetApp.getUi();
  const r = syncContactsFromSignup();
  let msg = 'Filled contact details for ' + r.filled + ' player' + (r.filled !== 1 ? 's' : '') +
    (r.bySurname > 0 ? ' (' + r.bySurname + ' matched by surname - check the Source column)' : '') + '.\n\n' +
    'Existing entries were left untouched.';
  if (r.ambiguous.length) {
    msg += '\n\nSame surname as more than one family - set these by hand:\n' +
      r.ambiguous.slice(0, 10).join(', ') + (r.ambiguous.length > 10 ? ', ...' : '');
  }
  if (r.missing.length) {
    msg += '\n\nNo record in the sign-up sheet (' + r.missing.length + '):\n' +
      r.missing.slice(0, 15).join(', ') + (r.missing.length > 15 ? ', ...' : '');
  }
  ui.alert(msg);
}

function countContactsMissing() {
  const contacts = getPlayerContacts();
  let n = 0;
  for (const k in contacts) if (!contacts[k].phone) n++;
  return n;
}

// ============================================================
// MONTHLY BILLING REPORT
// ============================================================
// Generates a billing summary in a separate Google Sheet.
// Run this function manually at the end of each month,
// or set up a monthly trigger (Edit > Triggers).
// ============================================================

// Pricing lookup tables - total charged for N sessions
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

  // Beyond the table - use last table value + extra sessions at per-session rate
  const lastIndex = lookup.length - 1;
  const extraSessions = sessions - lastIndex;
  return lookup[lastIndex] + (extraSessions * rate[s]);
}

// onlyClinic: rebuild just that clinic's tab (used when a status is corrected
// from the Collections sheet); every clinic when left out.
function generateMonthlyBilling(monthOverride, yearOverride, onlyClinic) {
  const now = new Date();
  const billingMonth = monthOverride || now.getMonth() + 1;
  const billingYear = yearOverride || now.getFullYear();

  const monthName = new Date(billingYear, billingMonth - 1, 1)
    .toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  const statusOverrides = getStatusOverrides(monthName);   // Lisa's Member/Guest corrections

  const attendanceRows = getAttendanceForMonth(billingMonth, billingYear);
  if (attendanceRows.length === 0) {
    Logger.log('No attendance data found for ' + monthName);
    return;
  }

  // Group sessions by clinic -> player
  const clinicData = {}; // { clinic: { playerKey: { name, status, sessions } } }

  // A clinic meets at most once per day, so each kid counts at most one
  // session per clinic per day - neutralizes any duplicate rows.
  const seenSession = {};
  for (const row of attendanceRows) {
    const dayKey = row.clinic + '|||' + row.playerName.toString().trim().toLowerCase() + '|||' +
      (row.date.getMonth() + 1) + '/' + row.date.getDate() + '/' + row.date.getFullYear();
    if (seenSession[dayKey]) continue;
    seenSession[dayKey] = true;

    if (!clinicData[row.clinic]) clinicData[row.clinic] = {};
    const cd = clinicData[row.clinic];
    if (!cd[row.playerName]) {
      cd[row.playerName] = { name: row.playerName, status: row.status, sessions: 0 };
    }
    cd[row.playerName].sessions++;
  }

  for (const clinic in clinicData) {
    for (const key in clinicData[clinic]) {
      const p = clinicData[clinic][key];
      p.status = resolvePlayerStatus(statusOverrides, clinic, p.name, p.status);
    }
  }

  // Self-fill the Families tab before applying discounts. A single-clinic
  // rebuild uses the tab as it stands, to keep a status change quick.
  if (!onlyClinic) updateFamiliesList();
  const billingSS = SpreadsheetApp.openById(BILLING_SHEET_ID);
  const siblingOverrides = getSiblingOverrides();

  for (const clinic in clinicData) {
    if (onlyClinic && clinic !== onlyClinic) continue;
    const playerMap = clinicData[clinic];

    // Build billing rows with sibling discounts applied automatically
    const players = [];
    for (const key in playerMap) players.push(playerMap[key]);
    const billing = buildClinicBillingRows(clinic, players, siblingOverrides);
    const billingRows = billing.rows;

    const tabName = clinic + ' - Billing - ' + monthName;
    let sheet = billingSS.getSheetByName(tabName);

    // Preserve Charged?/Charged On from the existing tab so regenerating the
    // report never loses which families have already been charged
    const prevState = {};
    if (sheet) {
      const prevData = sheet.getDataRange().getValues();
      if (prevData.length > 1) {
        const prevHeaders = prevData[0];
        const cCol = prevHeaders.indexOf('Charged?');
        const dCol = prevHeaders.indexOf('Charged On');
        for (let i = 1; i < prevData.length; i++) {
          const pname = (prevData[i][0] || '').toString().trim().toLowerCase();
          if (!pname) continue;
          if (cCol === -1 || prevData[i][cCol] !== true) continue;
          prevState[pname] = {
            charged: true,
            chargedOn: dCol !== -1 ? prevData[i][dCol] : ''
          };
        }
      }
      billingSS.deleteSheet(sheet);
    }
    sheet = billingSS.insertSheet(tabName);

    // Header
    // Contact details and follow-up live on the Collections sheet, not here.
    // This tab stays a clean financial record: what was owed, what was charged.
    const headers = ['Player Name', 'Status', 'Sessions', 'Total', 'Sibling Discount',
      'Final Charge', 'Charged?', 'Charged On', 'Note'];
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#2e7d32');
    headerRange.setFontColor('white');

    // Data rows (restoring any charged ticks captured above)
    for (const row of billingRows) {
      const prev = prevState[row.name.toLowerCase()] || {};
      sheet.appendRow([
        row.name,
        row.status,
        row.sessions,
        row.total,
        row.discount > 0 ? row.discount : '',
        row.finalTotal,
        prev.charged === true,
        prev.charged === true ? (prev.chargedOn || new Date()) : '',
        row.siblingNote
      ]);
    }

    // Checkboxes, formats, and sibling highlights
    if (billingRows.length > 0) {
      const col = (name) => headers.indexOf(name) + 1;
      const chargedRange = sheet.getRange(2, col('Charged?'), billingRows.length, 1);
      chargedRange.insertCheckboxes();
      // insertCheckboxes() sets every cell it touches to FALSE, which would
      // untick every family already charged each time a month is regenerated.
      // Put the ticks carried over from the old tab back on.
      chargedRange.setValues(billingRows.map(r =>
        [(prevState[r.name.toLowerCase()] || {}).charged === true]));
      sheet.getRange(2, col('Total'), billingRows.length, 3).setNumberFormat('$#,##0.00');
      sheet.getRange(2, col('Charged On'), billingRows.length, 1).setNumberFormat('M/d/yyyy');
      for (let i = 0; i < billingRows.length; i++) {
        if (billingRows[i].isSibling) {
          sheet.getRange(i + 2, 1, 1, headers.length).setBackground('#fff9c4');
        }
      }
    }

    // Column widths
    const widths = [200, 80, 80, 100, 130, 110, 90, 110, 220];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
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
// UNCHARGED BILLING TRACKER
// ============================================================
// Each billing tab has a per-player "Charged?" checkbox; "Charged On" stamps
// itself automatically. The separate Collections spreadsheet is the shop
// manager's worklist: one row per family (biggest balance first) with the
// phone number and follow-up history, and a Paid box per month underneath.
// Ticking Paid there ticks Charged? on the billing tab. A short weekly email
// (Mon ~7am) to everyone on the self-installing "Billing Reminders" tab (in
// the ROSTER spreadsheet) says how many families owe and links to the sheet.
//
// ONE-TIME SETUP: run setupBillingReminders() once from the editor, then
// add your shop manager's email to the "Billing Reminders" tab.
// ============================================================

const BILLING_AGING_DAYS = 14;

// Columns on the Collections sheet. It is grouped by FAMILY: one bold row per
// family (who to text, what they owe in total, and the follow-up so far), then
// one row per month per clinic underneath, each with its own Paid box.
//
//   Family row: Family | Phone |       |        |         |        | Owed (total) |      | Outreach | Last Contact | Notes
//   Month row:         |       | Month | Clinic | Players | Status | Owed         | Paid |
//
// Families are ordered by what they owe, most first. Outreach / Last Contact /
// Notes belong to the shop manager and are carried across every refresh.
// Ticking Paid ticks Charged? on the billing tab, and changing Status re-prices
// that month (see onCollectionsEdit).
// The hidden Key column is how the script tells the two kinds of row apart and
// which billing tab and players a Paid box belongs to - nobody needs to see it.
const COLLECTIONS_HEADERS = ['Family', 'Phone', 'Month', 'Clinic', 'Players',
  'Status', 'Owed', 'Paid', 'Outreach', 'Last Contact', 'Notes', 'Key'];
const COLL_COL = { family: 1, phone: 2, month: 3, clinic: 4, players: 5, status: 6,
                   owed: 7, paid: 8, outreach: 9, last: 10, notes: 11, key: 12 };

// Attendance stores players as "Last, First". On this sheet they read the way
// you would say them out loud on a call.
function playerFullName(raw) {
  const s = (raw || '').toString().trim().replace(/\s+/g, ' ');
  if (s.indexOf(',') === -1) return s;
  const parts = s.split(',').map(x => x.trim()).filter(String);
  return (parts.length === 2 ? parts[1] + ' ' + parts[0] : parts.join(' ')).trim();
}

// Clinics in the order they appear on every report: youngest to oldest.
const CLINIC_DISPLAY_ORDER = ['Red Ball', 'Orange Ball', 'Green Ball',
  'MS Yellow Ball', 'HS Yellow Ball', 'Bruno'];

// How a family is labelled on the sheet. Families with no contact on file are
// keyed "Walter (no contact)"; the Phone cell already says "no number", so
// the name just reads "WALTER FAMILY".
function familyDisplayName(family) {
  const m = (family || '').toString().match(/^(.*) \(no contact\)$/);
  return (m ? m[1] + ' family' : (family || '').toString()).toUpperCase();
}

// "2052403601" -> "205-240-3601": easier to read, still plain text to copy.
// Anything that is not a 10-digit US number is left as it was.
function prettyPhone(p) {
  const raw = (p || '').toString().trim();
  let d = raw.replace(/\D/g, '');
  if (d.length === 11 && d.charAt(0) === '1') d = d.slice(1);
  return d.length === 10 ? d.slice(0, 3) + '-' + d.slice(3, 6) + '-' + d.slice(6) : raw;
}

// "September 2026" -> "Sep 2026". Written as plain text like the long form,
// so Sheets cannot turn it into a date.
function shortMonthLabel(label) {
  const parts = (label || '').toString().trim().split(' ');
  return MONTH_NAMES_FULL.indexOf(parts[0]) === -1 || !parts[1]
    ? (label || '').toString() : parts[0].slice(0, 3) + ' ' + parts[1];
}

// Clinics we do not collect money for. They still appear on the billing and
// A/S reports - they are simply left off the Collections worklist, since
// chasing those balances is not ours to do. Matched case-insensitively.
const COLLECTIONS_SKIP_CLINICS = ['Bruno'];

function collectionsSkips(clinic) {
  const c = (clinic || '').toString().trim().toLowerCase();
  return COLLECTIONS_SKIP_CLINICS.some(x => x.toLowerCase() === c);
}

// Hidden Key column. A family row is "F|||<family>", plus "|||<phone>" when
// someone typed a number for that family on this sheet (see
// keepCollectionsPhone); a month row is
// "M|||<family>|||<month>|||<clinic>|||<Last, First>;;<Last, First>" - the
// players exactly as the billing tab spells them, so a Paid tick can find them.
const COLL_KEY_SEP = '|||';

function familyRowKey(family, typedPhone) {
  return 'F' + COLL_KEY_SEP + family + (typedPhone ? COLL_KEY_SEP + typedPhone : '');
}

function monthRowKey(line) {
  return ['M', line.family, line.month, line.clinic, line.raws.join(';;')].join(COLL_KEY_SEP);
}

function parseCollectionsRowKey(v) {
  const parts = (v || '').toString().split(COLL_KEY_SEP);
  if (parts[0] === 'F' && parts[1]) return { type: 'F', family: parts[1], phone: parts[2] || '' };
  if (parts[0] === 'M' && parts.length >= 4) {
    return { type: 'M', family: parts[1], month: parts[2], clinic: parts[3],
             raws: (parts[4] || '').split(';;').filter(String) };
  }
  return null;
}

// Rebuilds the Collections sheet from the billing tabs. The billing sheet is
// the truth for who has been charged; the sheet is rewritten each time so it
// stays sorted, but Outreach / Last Contact / Notes are carried across by
// family, so nothing the shop manager typed is ever lost.
function refreshCollections(monthsBack) {
  if (!COLLECTIONS_SHEET_ID) throw new Error('COLLECTIONS_SHEET_ID is not set.');
  const months = monthsBack || 3;

  const contacts = getPlayerContacts();
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - (months - 1));
  cutoff.setDate(1); cutoff.setHours(0,0,0,0);

  // --- Every billed player in the window, charged or not ------------------
  // Grouped by month + clinic + family. Charged players are gathered too, so
  // a month that has just been paid can still be shown (greyed, box ticked).
  const ss = SpreadsheetApp.openById(BILLING_SHEET_ID);
  const lines = {};
  for (const sheet of ss.getSheets()) {
    const name = sheet.getName();
    const idx = name.indexOf(' - Billing - ');
    if (idx === -1) continue;
    const clinic = name.substring(0, idx).trim();
    if (collectionsSkips(clinic)) continue;        // not ours to collect
    const monthLabel = name.substring(idx + ' - Billing - '.length).trim();
    const parts = monthLabel.split(' ');
    const mIdx = MONTH_NAMES_FULL.indexOf(parts[0]);
    if (mIdx === -1) continue;
    const monthDate = new Date(parseInt(parts[1], 10), mIdx, 1);
    if (monthDate < cutoff) continue;                       // too old to chase

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;
    const h = data[0];
    const iC = h.indexOf('Charged?'), iF = h.indexOf('Final Charge'), iS = h.indexOf('Sessions');
    const iSt = h.indexOf('Status');
    if (iC === -1 || iF === -1 || iS === -1) continue;

    for (let i = 1; i < data.length; i++) {
      if (typeof data[i][iS] !== 'number') continue;        // skips summary rows
      const player = (data[i][0] || '').toString().trim();
      if (!player) continue;

      const c = contacts[player.toLowerCase()] || {};
      // Families are keyed by the parent we'd actually text; without a
      // contact on file, fall back to surname so siblings still group.
      const family = c.parent || (player.split(',')[0].trim() + ' (no contact)');
      const key = collectionsKey(monthLabel, clinic, family);
      if (!lines[key]) {
        lines[key] = { month: monthLabel, monthDate: monthDate, clinic: clinic, family: family,
                       phone: '', open: { players: [], raws: [], amount: 0, statuses: [] },
                       done: { players: [], raws: [], amount: 0, statuses: [] } };
      }
      const part = data[i][iC] === true ? lines[key].done : lines[key].open;
      part.players.push(playerFullName(player));
      part.raws.push(player);
      part.amount += Number(data[i][iF]) || 0;
      part.statuses.push(iSt === -1 ? 'M' : (statusCode(data[i][iSt]) || 'M'));
      if (!lines[key].phone && c.phone) lines[key].phone = c.phone;
    }
  }

  // --- Preserve whatever the shop manager typed last time -----------------
  const cs = SpreadsheetApp.openById(COLLECTIONS_SHEET_ID);
  const sheet = cs.getSheets()[0];
  const kept = readCollectionsNotes(sheet);
  const history = (family) => kept.families[family.toLowerCase()] || {};

  // --- Group months under families ----------------------------------------
  const fams = {};
  let added = 0, carried = 0;
  for (const key in lines) {
    const L = lines[key];
    const paid = L.open.raws.length === 0;
    const prev = kept.lines[key];
    if (paid) {
      // Charged before it ever reached this sheet: nothing to follow up.
      if (!prev) continue;
      // A paid month stays for one refresh so the tick is visible, and for as
      // long as the family has outreach or notes on file. Otherwise it retires.
      const hist = history(L.family);
      if (prev.paid && !hist.outreach && !hist.notes) continue;
    }
    if (prev) carried++; else added++;
    const part = paid ? L.done : L.open;
    const fk = L.family.toLowerCase();
    if (!fams[fk]) fams[fk] = { family: L.family, phone: '', owed: 0, lines: [] };
    const f = fams[fk];
    if (!f.phone && L.phone) f.phone = L.phone;
    if (!paid) f.owed += part.amount;
    f.lines.push({ family: L.family, month: L.month, monthDate: L.monthDate, clinic: L.clinic,
                   players: part.players, raws: part.raws, amount: part.amount, paid: paid,
                   statuses: part.statuses });
  }

  // --- Order: biggest balance first; each family's months oldest first ----
  const clinicRank = (c) => {
    const i = CLINIC_DISPLAY_ORDER.indexOf(c);
    return i === -1 ? CLINIC_DISPLAY_ORDER.length : i;
  };
  const cents = (x) => Math.round(x * 100);
  const famList = Object.keys(fams).map(k => fams[k]);
  famList.forEach(f => f.lines.sort((a, b) =>
    (a.monthDate - b.monthDate) || (clinicRank(a.clinic) - clinicRank(b.clinic))));
  famList.sort((a, b) => (cents(b.owed) - cents(a.owed)) ||
    (a.family.toLowerCase() < b.family.toLowerCase() ? -1 : 1));

  // --- Rewrite the sheet --------------------------------------------------
  sheet.clear();
  sheet.clearConditionalFormatRules();
  const whole = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  // clear() wipes contents and formatting but NOT data validation, so a
  // dropdown or checkbox from an earlier layout would be left stranded on
  // whatever column now sits in that position. The old layout also merged a
  // month band across the row, which would swallow a family row.
  whole.clearDataValidations();
  whole.breakApart();
  // clear() leaves number formats behind as well. One cell carrying a stray
  // "Plain text" format turned every FALSE written into that Paid checkbox
  // into the text "false", so the box showed as invalid text on every
  // refresh (Shepherd Wilson's July row, 9/28). Back to Automatic everywhere;
  // the columns that need a format get it again below.
  whole.setNumberFormat('General');
  const W = COLLECTIONS_HEADERS.length;
  sheet.getRange(1, 1, 1, W).setValues([COLLECTIONS_HEADERS])
    .setFontWeight('bold').setBackground('#021f3d').setFontColor('white')
    .setFontSize(10).setVerticalAlignment('middle');
  sheet.setRowHeight(1, 28);
  [230, 115, 80, 120, 210, 75, 85, 45, 115, 95, 280, 80]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);          // the family name stays put when scrolling right
  sheet.setHiddenGridlines(true);     // a line between families does the separating
  if (sheet.getName() !== 'Collections') sheet.setName('Collections');

  const out = [];          // values to write
  const famRows = [];      // { row, owing, escalate, noPhone }
  const monthRows = [];    // { row, clinic, paid }
  let owedTotal = 0, owing = 0, escalate = 0, notContacted = 0, noPhone = 0;

  for (const f of famList) {
    const hist = history(f.family);
    const isOwing = cents(f.owed) > 0;
    const isEscalated = OUTREACH_ESCALATE.indexOf(hist.outreach || '') !== -1;
    // A number typed on this sheet beats the one on file.
    const phone = hist.phone || f.phone;
    if (isOwing) {
      owing++;
      owedTotal += f.owed;
      if (!hist.outreach) notContacted++;
      if (isEscalated) escalate++;
      if (!phone) noPhone++;
    }
    // What the cells SHOW is for reading; the hidden Key keeps the real family
    // name, which is what outreach history and the Paid / Status edits go by.
    out.push([familyDisplayName(f.family),
              phone ? prettyPhone(phone) : (isOwing ? 'no number' : ''),
              '', '', '', '', f.owed, '',
              hist.outreach || '', hist.last || '', hist.notes || '',
              familyRowKey(f.family, hist.phone)]);
    famRows.push({ row: out.length + 1, owing: isOwing, escalate: isOwing && isEscalated,
                   noPhone: !phone });
    for (const l of f.lines) {
      out.push(['', '', shortMonthLabel(l.month), l.clinic, l.players.join(', '),
                statusCellLabel(l.statuses), l.amount, l.paid, '', '', '', monthRowKey(l)]);
      monthRows.push({ row: out.length + 1, clinic: l.clinic, paid: l.paid });
    }
  }

  if (out.length > 0) {
    const n = out.length;
    // Formats go on BEFORE the values. Sheets silently converts anything that
    // looks like a date on write, so "August 2026" would land as a real Date
    // and read back as "Sat Aug 01 2026 00:00:00 GMT-0500...". Same for phone
    // numbers, which would otherwise lose a leading zero.
    sheet.getRange(2, COLL_COL.phone, n, 1).setNumberFormat('@');
    sheet.getRange(2, COLL_COL.month, n, 1).setNumberFormat('@');
    sheet.getRange(2, COLL_COL.owed, n, 1).setNumberFormat('$#,##0.00');
    sheet.getRange(2, COLL_COL.last, n, 1).setNumberFormat('M/d/yyyy');
    sheet.getRange(2, COLL_COL.key, n, 1).setNumberFormat('@');
    // insertCheckboxes() resets every cell it touches to unticked, so the
    // boxes go in before the values that say which ones are ticked.
    monthRows.forEach(m => sheet.getRange(m.row, COLL_COL.paid).insertCheckboxes());
    sheet.getRange(2, 1, n, W).setValues(out);
  }
  sheet.hideColumns(COLL_COL.key);

  // Layout: each family reads as one block. The family line carries the
  // colour, weight and a navy rule above it; the months under it are plain,
  // smaller and grey. Colour is kept for what needs attention - escalation,
  // no phone number - and nothing else.
  if (out.length > 0) {
    sheet.getRange(2, 1, out.length, W - 1)
      .setFontSize(10).setFontColor('#5f6368').setVerticalAlignment('middle');
  }
  famRows.forEach(d => {
    sheet.getRange(d.row, 1, 1, W - 1)
      .setBackground(d.escalate ? '#ffebee' : '#eef1f7')
      .setFontColor(d.owing ? '#021f3d' : '#9e9e9e').setFontSize(11)
      .setBorder(true, null, null, null, null, null, '#021f3d',
                 SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(d.row, COLL_COL.family).setFontWeight('bold');
    sheet.getRange(d.row, COLL_COL.owed).setFontWeight('bold');
    sheet.setRowHeight(d.row, 30);
    if (d.owing && d.noPhone) {
      sheet.getRange(d.row, COLL_COL.phone).setFontColor('#c62828').setFontStyle('italic');
    }
    if (d.owing) {
      sheet.getRange(d.row, COLL_COL.outreach).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(OUTREACH_LEVELS, true).build());
    }
  });
  monthRows.forEach(d => {
    if (d.paid) {
      sheet.getRange(d.row, 1, 1, W - 1).setFontColor('#b0b0b0');
    } else {
      // A paid month has gone through Jonas, so only unpaid ones can be re-priced.
      sheet.getRange(d.row, COLL_COL.status).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(STATUS_LABELS, true).build());
    }
  });

  Logger.log('Collections: ' + owing + ' families owing $' + owedTotal.toFixed(2) +
    ' (' + famList.length + ' families, ' + monthRows.length + ' month rows).');
  return { added: added, updated: carried, outstanding: owing, owed: owedTotal,
           families: famList.length, escalate: escalate, notContacted: notContacted,
           noPhone: noPhone };
}

// month|||clinic|||family, lowercased - the identity of one month row.
function collectionsKey(month, clinic, family) {
  return [month, clinic, family].map(s => (s || '').toString().trim().toLowerCase()).join('|||');
}

// A Month cell should read "August 2026". Older refreshes let Sheets coerce
// that into a real date, which then got written back as text, so an old sheet
// can hold three shapes: the label, a Date, or "Sat Aug 01 2026 00:00:00
// GMT-0500 (Central Daylight Time)". All three come back as the label here.
function monthLabelOf(v) {
  if (v instanceof Date) return monthLabelOfDate(v);
  const s = (v || '').toString().trim();
  if (!s) return '';
  const parts = s.split(' ');
  if (MONTH_NAMES_FULL.indexOf(parts[0]) !== -1 && /^\d{4}$/.test(parts[1] || '')) {
    return parts[0] + ' ' + parts[1];                    // already a clean label
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : monthLabelOfDate(d);
}

// A month label always meant the 1st. A timezone shift can render that instant
// as the last evening of the month before, so take whichever of local/UTC
// still lands on the 1st - otherwise "July 2026" comes back as "June 2026".
function monthLabelOfDate(d) {
  const useUTC = d.getUTCDate() === 1 && d.getDate() !== 1;
  const m = useUTC ? d.getUTCMonth() : d.getMonth();
  const y = useUTC ? d.getUTCFullYear() : d.getFullYear();
  return MONTH_NAMES_FULL[m] + ' ' + y;
}

// Reads back everything a refresh must not clobber:
//   families - outreach / last contact / notes / a typed phone, by family
//   lines    - whether each month row was ticked Paid, by month|||clinic|||family
// Understands the current family-grouped layout and the older layouts that
// had one row per month per clinic with follow-up on every row. When an older
// sheet has several rows for one family, the furthest-along outreach wins
// (with its date) and every distinct note is kept.
function readCollectionsNotes(sheet) {
  const kept = { families: {}, lines: {} };
  if (!sheet || sheet.getLastRow() < 2) return kept;
  const data = sheet.getDataRange().getValues();
  const h = data[0].map(v => (v || '').toString().trim());
  const col = (n) => h.indexOf(n);
  const iOut = col('Outreach'), iLast = col('Last Contact'), iNo = col('Notes');
  const text = (row, i) => i >= 0 ? (row[i] || '').toString().trim() : '';

  const noteLists = {};
  const addHistory = (family, row) => {
    const k = family.toLowerCase();
    const f = kept.families[k] ||
      (kept.families[k] = { family: family, outreach: '', last: '', notes: '', phone: '' });
    const outreach = text(row, iOut);
    if (outreach && (!f.outreach ||
        OUTREACH_LEVELS.indexOf(outreach) > OUTREACH_LEVELS.indexOf(f.outreach))) {
      f.outreach = outreach;
      f.last = iLast >= 0 ? row[iLast] : '';
    }
    const note = text(row, iNo);
    const list = noteLists[k] || (noteLists[k] = []);
    if (note && list.indexOf(note) === -1) list.push(note);
  };
  const finish = () => {
    for (const k in kept.families) kept.families[k].notes = (noteLists[k] || []).join('; ');
    return kept;
  };

  const iKey = col('Key');
  if (iKey !== -1) {
    const iPaid = col('Paid');
    for (let i = 1; i < data.length; i++) {
      const k = parseCollectionsRowKey(data[i][iKey]);
      if (!k) continue;
      if (k.type === 'F') {
        addHistory(k.family, data[i]);
        if (k.phone) kept.families[k.family.toLowerCase()].phone = k.phone;
      }
      else kept.lines[collectionsKey(k.month, k.clinic, k.family)] =
        { paid: iPaid >= 0 && data[i][iPaid] === true };
    }
    return finish();
  }

  // Older layout. Month band rows have no Family, so they are skipped.
  const iM = col('Month'), iC = col('Clinic'), iFam = col('Family'), iSt = col('Status');
  if (iFam === -1) return kept;
  for (let i = 1; i < data.length; i++) {
    const family = text(data[i], iFam);
    if (!family) continue;
    const month = iM >= 0 ? monthLabelOf(data[i][iM]) : '';
    kept.lines[collectionsKey(month, text(data[i], iC), family)] =
      { paid: text(data[i], iSt) === 'Paid' };
    addHistory(family, data[i]);
  }
  return finish();
}

// Installable onEdit trigger on the COLLECTIONS spreadsheet.
//  - Outreach set on a family row: stamps Last Contact.
//  - Paid ticked on a month row: ticks Charged? (and stamps Charged On) for
//    those players on the matching billing tab. Unticking reverses it.
//  - Status changed on a month row: re-prices that month (see
//    changeCollectionsStatus).
//  - Phone typed on a family row: kept for that family (keepCollectionsPhone).
function onCollectionsEdit(e) {
  try {
    const sheet = e.range.getSheet();
    const h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const oCol = h.indexOf('Outreach') + 1, lCol = h.indexOf('Last Contact') + 1;
    const pCol = h.indexOf('Paid') + 1, kCol = h.indexOf('Key') + 1;
    const sCol = h.indexOf('Status') + 1, wCol = h.indexOf('Owed') + 1;
    const first = e.range.getColumn(), last = first + e.range.getNumColumns() - 1;
    const start = e.range.getRow(), n = e.range.getNumRows();
    const touches = (c) => c > 0 && c >= first && c <= last;

    if (touches(oCol) && lCol) {
      for (let r = 0; r < n; r++) {
        const row = start + r;
        if (row === 1) continue;
        const v = sheet.getRange(row, oCol).getValue();
        const cell = sheet.getRange(row, lCol);
        cell.setValue(v === '' || v === null ? '' : new Date());
        cell.setNumberFormat('M/d/yyyy');
      }
    }
    if (touches(pCol) && kCol) syncCollectionsPaid(sheet, start, n, pCol, kCol);
    if (touches(sCol) && kCol && pCol && wCol) {
      changeCollectionsStatus(sheet, start, n,
        { status: sCol, paid: pCol, key: kCol, owed: wCol }, e);
    }
    const phCol = h.indexOf('Phone') + 1;
    if (touches(phCol) && kCol) keepCollectionsPhone(sheet, start, n, { phone: phCol, key: kCol }, e);
  } catch (err) { /* never break an edit */ }
}

// ---- A phone number typed on a family line ---------------------------------
// The Phone cell is redrawn on every refresh, so a number typed into it is
// kept in that family line's hidden Key instead ("F|||Walter (no contact)|||
// 205-555-7777"), and every refresh shows it in place of the number on file.
// It only affects this sheet: the Contacts list is not touched. Clearing the
// cell goes back to the number on file, if there is one.
function keepCollectionsPhone(sheet, start, n, cols, e) {
  const keys = sheet.getRange(start, cols.key, n, 1).getValues();
  const vals = sheet.getRange(start, cols.phone, n, 1).getValues();
  const single = n === 1 && e && e.range && e.range.getNumColumns() === 1;
  const before = single && e.oldValue !== undefined ? e.oldValue : '';
  const today = (new Date().getMonth() + 1) + '/' + new Date().getDate();

  for (let i = 0; i < n; i++) {
    const row = start + i;
    if (row === 1) continue;
    const k = parseCollectionsRowKey(keys[i][0]);
    if (!k) continue;
    const cell = sheet.getRange(row, cols.phone);
    if (k.type !== 'F') {                                     // month lines have no phone
      cell.setValue('');
      cell.setNote('Type the number on the family line above.');
      continue;
    }
    const keyCell = sheet.getRange(row, cols.key);
    const typed = (vals[i][0] || '').toString().trim();

    if (!typed || typed.toLowerCase() === 'no number') {
      if (k.phone) {
        keyCell.setValue(familyRowKey(k.family));             // forget the typed number
        cell.setNote('Removed ' + k.phone + '. The next refresh shows the number on file, if there is one.');
      } else {
        cell.setValue(before);                                 // nothing typed to remove
        cell.setNote('This number is the one on file, so it comes back on every refresh. ' +
          'To use a different one, type it here.');
      }
      continue;
    }

    let digits = typed.replace(/\D/g, '');
    if (digits.length === 11 && digits.charAt(0) === '1') digits = digits.slice(1);
    if (digits.length !== 10) {
      cell.setValue(before);
      cell.setNote('"' + typed + '" is not a phone number - nothing was saved. ' +
        'Type 10 digits, e.g. 205-555-1234.');
      continue;
    }
    const pretty = prettyPhone(digits);
    const was = k.phone || (before && before !== 'no number' ? before : '');
    keyCell.setValue(familyRowKey(k.family, pretty));
    cell.setValue(pretty);
    cell.setFontColor('#021f3d').setFontStyle('normal');
    cell.setNote('Typed on this sheet ' + today + (was && was !== pretty ? ' (was ' + was + ')' : '') +
      '. Stays for this family on every refresh. Clear the cell to go back to the number on file.');
  }
}

// ---- Member / Guest corrections ------------------------------------------
// Lisa can correct a kid's status from the Collections sheet. The correction
// is NOT written into the Master Attendance sheet - the coaches' check-in rows
// stay exactly as recorded. It goes on the "Status Changes" tab of the billing
// spreadsheet (private; the roster sheet is readable by link), and every
// report that turns status into dollars reads it: billing and the three
// summaries. The roster is fixed as well, so check-ins from then on come in
// right and later months need no correction at all.
//
// Where a month has no correction, a kid's FIRST check-in of the month decides.
// Billing always worked that way; the summaries used to take the last one, so
// the two could disagree about a kid whose badge changed mid-month.
const STATUS_CHANGES_TAB = 'Status Changes';
const STATUS_CHANGES_HEADERS = ['Player', 'Clinic', 'Month', 'New Status', 'Old Status',
  'Old Charge', 'New Charge', 'Changed On', 'Changed By'];
const STATUS_LABELS = ['Member', 'Guest', 'Social'];

// 'M' | 'G' | 'S' from a letter or a word; '' when it is neither.
function statusCode(v) {
  const s = (v || '').toString().trim().toUpperCase();
  if (s === 'M' || s === 'MEMBER') return 'M';
  if (s === 'G' || s === 'GUEST') return 'G';
  if (s === 'S' || s === 'SOCIAL') return 'S';
  return '';
}

function statusLabel(code) {
  return code === 'G' ? 'Guest' : code === 'S' ? 'Social' : 'Member';
}

// What the Status cell shows for a row of one or more siblings.
function statusCellLabel(codes) {
  const uniq = (codes || []).filter((c, i, a) => c && a.indexOf(c) === i);
  if (uniq.length === 0) return '';
  return uniq.length === 1 ? statusLabel(uniq[0]) : 'Mixed';
}

function statusOverrideKey(clinic, player) {
  return (clinic || '').toString().trim().toLowerCase() + '|||' +
         (player || '').toString().trim().replace(/\s+/g, ' ').toLowerCase();
}

// { 'clinic|||last, first': 'M'|'G'|'S' } for one month ("September 2026").
// The tab is a log - a later row wins, so changing a kid back and forth
// keeps every step and uses the latest.
function getStatusOverrides(monthLabel) {
  const out = {};
  const sheet = SpreadsheetApp.openById(BILLING_SHEET_ID).getSheetByName(STATUS_CHANGES_TAB);
  if (!sheet || sheet.getLastRow() < 2) return out;
  const want = (monthLabel || '').toString().trim().toLowerCase();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (monthLabelOf(data[i][2]).toLowerCase() !== want) continue;
    const code = statusCode(data[i][3]);
    if (code) out[statusOverrideKey(data[i][1], data[i][0])] = code;
  }
  return out;
}

// A kid's status for one clinic-month: Lisa's correction if there is one,
// otherwise what was recorded at check-in.
function resolvePlayerStatus(overrides, clinic, player, recorded) {
  return (overrides && overrides[statusOverrideKey(clinic, player)]) || recorded || 'M';
}

function statusChangesSheet(billing) {
  let sheet = billing.getSheetByName(STATUS_CHANGES_TAB);
  if (!sheet) {
    sheet = billing.insertSheet(STATUS_CHANGES_TAB);
    sheet.getRange(1, 1, 1, STATUS_CHANGES_HEADERS.length).setValues([STATUS_CHANGES_HEADERS])
      .setFontWeight('bold').setBackground('#021f3d').setFontColor('white');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Reads one billing tab as { 'last, first': { name, code, amount, charged } }.
function readBillingPlayers(billing, tabName) {
  const tab = billing.getSheetByName(tabName);
  if (!tab) throw new Error('There is no billing tab called "' + tabName + '".');
  const data = tab.getDataRange().getValues();
  const h = data[0];
  const iSt = h.indexOf('Status'), iF = h.indexOf('Final Charge'),
        iC = h.indexOf('Charged?'), iS = h.indexOf('Sessions');
  const out = {};
  for (let i = 1; i < data.length; i++) {
    if (iS !== -1 && typeof data[i][iS] !== 'number') continue;   // summary rows
    const name = (data[i][0] || '').toString().trim();
    if (!name) continue;
    out[name.toLowerCase()] = { name: name, code: statusCode(data[i][iSt]) || 'M',
      amount: Number(data[i][iF]) || 0, charged: iC !== -1 && data[i][iC] === true };
  }
  return out;
}

// Handles a change to the Status column. Each changed month row is re-priced;
// anything that cannot be done is put back, with a note on the cell saying
// why, so the sheet never shows a status that billing does not agree with.
function changeCollectionsStatus(sheet, start, n, cols, e) {
  const lock = LockService.getScriptLock();
  const keys = sheet.getRange(start, cols.key, n, 1).getValues();
  const vals = sheet.getRange(start, cols.status, n, 1).getValues();
  const paid = sheet.getRange(start, cols.paid, n, 1).getValues();
  let billing = null;
  const current = (k) => {                  // what billing says right now
    try {
      const players = readBillingPlayers(billing, k.clinic + ' - Billing - ' + k.month);
      return statusCellLabel(k.raws.map(r => (players[r.toLowerCase()] || {}).code));
    } catch (err) { return ''; }
  };

  let locked = false;
  try { locked = lock.tryLock(30000); } catch (err) { locked = false; }
  try {
    try { billing = SpreadsheetApp.openById(BILLING_SHEET_ID); } catch (err) { billing = null; }
    let who = '';
    try { who = (e && e.user && e.user.getEmail && e.user.getEmail()) || ''; } catch (err) { who = ''; }

    const changedRows = [];
    for (let i = 0; i < n; i++) {
      const row = start + i;
      if (row === 1) continue;
      const k = parseCollectionsRowKey(keys[i][0]);
      if (!k) continue;
      const cell = sheet.getRange(row, cols.status);
      if (k.type !== 'M') { cell.setValue(''); continue; }        // family rows have none
      const putBack = (why) => {
        cell.setValue(billing ? current(k) : '');
        cell.setNote(why + ' Nothing was changed.');
      };
      if (!locked) { putBack('Another change was still being saved. Try again in a moment.'); continue; }
      if (!billing) { putBack('Could not open the billing sheet.'); continue; }
      if (paid[i][0] === true) {
        putBack('This month is ticked Paid, so it has already been charged in Jonas. ' +
          'Refund or charge the difference in Jonas instead.');
        continue;
      }
      const code = statusCode(vals[i][0]);
      if (!code) { putBack('Pick Member, Guest or Social.'); continue; }

      let r;
      try {
        r = applyStatusChange(billing, k, code, who);
      } catch (err) {
        putBack(err.message);
        continue;
      }
      cell.setValue(statusLabel(code));
      if (r.unchanged) { cell.clearNote(); continue; }
      const d = new Date();
      cell.setNote((d.getMonth() + 1) + '/' + d.getDate() + ': ' + r.oldLabel + ' -> ' +
        statusLabel(code) + ', $' + r.oldAmt.toFixed(2) + ' -> $' + r.newAmt.toFixed(2) + '.' + r.warn);
      sheet.getRange(row, cols.owed).setValue(r.newAmt);
      changedRows.push(row);
    }
    if (changedRows.length) updateCollectionsFamilyTotals(sheet, changedRows);
  } finally {
    if (locked) lock.releaseLock();
  }
}

// Re-prices one Collections month row at a new status: logs the correction,
// rebuilds that clinic's billing tab for that month (Charged? ticks are kept
// by the rebuild), and fixes the roster for future check-ins. Throws with a
// sentence the sheet can show if it cannot.
function applyStatusChange(billing, k, code, who) {
  const tabName = k.clinic + ' - Billing - ' + k.month;
  const before = readBillingPlayers(billing, tabName);
  const players = k.raws.map(r => before[r.toLowerCase()]);
  const missing = k.raws.filter((r, i) => !players[i]);
  if (missing.length) {
    throw new Error('Could not find ' + missing.map(playerFullName).join(', ') +
      ' on the "' + tabName + '" tab.');
  }
  if (players.some(p => p.charged)) {
    throw new Error('Already charged on the billing sheet - refresh the Collections list.');
  }
  const oldAmt = players.reduce((s, p) => s + p.amount, 0);
  const oldLabel = statusCellLabel(players.map(p => p.code));
  if (players.every(p => p.code === code)) {
    return { unchanged: true, oldLabel: oldLabel, oldAmt: oldAmt, newAmt: oldAmt, warn: '' };
  }

  const parts = k.month.split(' ');
  const mIdx = MONTH_NAMES_FULL.indexOf(parts[0]);
  const year = parseInt(parts[1], 10);
  if (mIdx === -1 || !year) throw new Error('Could not read the month "' + k.month + '".');

  // 1. Log it. The rebuild below reads this.
  const log = statusChangesSheet(billing);
  const first = log.getLastRow() + 1;
  const now = new Date();
  const rows = players.map(p => [p.name, k.clinic, k.month, statusLabel(code),
    statusLabel(p.code), p.amount, '', now, who]);
  const W = STATUS_CHANGES_HEADERS.length;
  log.getRange(first, 3, rows.length, 1).setNumberFormat('@');    // before the value - see refreshCollections
  log.getRange(first, 1, rows.length, W).setValues(rows);
  log.getRange(first, 6, rows.length, 2).setNumberFormat('$#,##0.00');
  log.getRange(first, 8, rows.length, 1).setNumberFormat('M/d/yyyy h:mm am/pm');

  // 2. Rebuild just this clinic's billing tab for this month.
  const chargedBefore = {};
  for (const nm in before) if (before[nm].charged) chargedBefore[nm] = before[nm].amount;
  try {
    generateMonthlyBilling(mIdx + 1, year, k.clinic);
  } catch (err) {
    log.deleteRows(first, rows.length);     // a logged change that was never applied would re-price later
    throw new Error('Could not rebuild the "' + tabName + '" tab (' + err.message + ').');
  }
  const after = readBillingPlayers(billing, tabName);
  const amounts = k.raws.map(r => (after[r.toLowerCase()] || {}).amount || 0);
  log.getRange(first, 7, rows.length, 1).setValues(amounts.map(a => [a]));
  const newAmt = amounts.reduce((s, a) => s + a, 0);

  // 3. Fix the roster so the app shows the right badge from now on.
  let warn = '';
  try {
    updateRosterStatus(k.raws, code);
  } catch (err) {
    warn += ' The roster could not be updated (' + err.message + ') - change it there by hand.';
  }

  // A sibling discount can move between kids when one of them changes price.
  // If it moved onto a kid who was already charged, Jonas needs a look.
  const moved = [];
  for (const nm in chargedBefore) {
    const a = after[nm];
    if (a && Math.round(a.amount * 100) !== Math.round(chargedBefore[nm] * 100)) {
      moved.push(playerFullName(a.name) + ' ($' + chargedBefore[nm].toFixed(2) + ' -> $' + a.amount.toFixed(2) + ')');
    }
  }
  if (moved.length) {
    warn += ' Heads up: the sibling discount changed for ' + moved.join(', ') +
      ', who was already charged - check Jonas.';
  }
  return { unchanged: false, oldLabel: oldLabel, oldAmt: oldAmt, newAmt: newAmt, warn: warn };
}

// Sets a kid's status on every clinic roster tab they appear on - membership
// belongs to the family, not to one clinic. Returns how many rows changed.
function updateRosterStatus(raws, code) {
  const want = raws.map(r => r.toString().trim().replace(/\s+/g, ' ').toLowerCase());
  const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  let changed = 0;
  for (const tabName of CLINIC_ROSTER_TABS) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet || sheet.getLastRow() < 2) continue;
    const data = sheet.getDataRange().getValues();
    const iSt = data[0].map(v => (v || '').toString().trim().toLowerCase()).indexOf('status');
    const col = iSt === -1 ? 3 : iSt + 1;
    for (let i = 1; i < data.length; i++) {
      const last = (data[i][0] || '').toString().trim();
      const firstName = (data[i][1] || '').toString().trim();
      const full = (firstName ? last + ', ' + firstName : last).replace(/\s+/g, ' ').toLowerCase();
      if (!full || want.indexOf(full) === -1) continue;
      if (statusCode(data[i][col - 1]) === code) continue;
      sheet.getRange(i + 1, col).setValue(code);
      changed++;
    }
  }
  return changed;
}

// Pushes Paid ticks from the Collections sheet to the billing tabs. If a tick
// cannot be applied, the box is put back and a note on the cell says why -
// the Collections sheet must never claim a charge the billing sheet lacks.
function syncCollectionsPaid(sheet, start, n, pCol, kCol) {
  const keys = sheet.getRange(start, kCol, n, 1).getValues();
  const ticks = sheet.getRange(start, pCol, n, 1).getValues();
  let billing = null, openError = '';
  try {
    billing = SpreadsheetApp.openById(BILLING_SHEET_ID);
  } catch (err) {
    openError = 'Could not open the billing sheet (' + err.message + ').';
  }

  const changedRows = [];
  for (let i = 0; i < n; i++) {
    const row = start + i;
    if (row === 1) continue;
    const k = parseCollectionsRowKey(keys[i][0]);
    if (!k || k.type !== 'M') continue;              // family rows have no box
    if (typeof ticks[i][0] !== 'boolean') continue;
    const paid = ticks[i][0];
    const cell = sheet.getRange(row, pCol);
    const problem = openError || setBillingCharged(billing, k, paid);
    if (problem) {
      cell.setValue(!paid);
      cell.setNote(problem + ' Nothing was changed on the billing sheet.');
      continue;
    }
    cell.clearNote();
    sheet.getRange(row, 1, 1, COLLECTIONS_HEADERS.length - 1)
      .setFontColor(paid ? '#9e9e9e' : '#000000');
    changedRows.push(row);
  }
  if (changedRows.length) updateCollectionsFamilyTotals(sheet, changedRows);
}

// Ticks or unticks Charged? for a month row's players on its billing tab.
// Returns '' on success, or a sentence saying what went wrong.
function setBillingCharged(billing, k, paid) {
  const tabName = k.clinic + ' - Billing - ' + k.month;
  const tab = billing.getSheetByName(tabName);
  if (!tab) return 'There is no billing tab called "' + tabName + '".';
  const data = tab.getDataRange().getValues();
  const h = data[0];
  const iC = h.indexOf('Charged?'), iD = h.indexOf('Charged On'), iS = h.indexOf('Sessions');
  if (iC === -1) return 'The "' + tabName + '" tab has no Charged? column.';
  const want = k.raws.map(x => x.toLowerCase());
  let matched = 0;
  for (let i = 1; i < data.length; i++) {
    if (iS !== -1 && typeof data[i][iS] !== 'number') continue;   // summary rows
    const name = (data[i][0] || '').toString().trim().toLowerCase();
    if (want.indexOf(name) === -1) continue;
    matched++;
    if ((data[i][iC] === true) === paid) continue;               // already right
    tab.getRange(i + 1, iC + 1).setValue(paid);
    if (iD !== -1) {
      tab.getRange(i + 1, iD + 1).setValue(paid ? new Date() : '').setNumberFormat('M/d/yyyy');
    }
  }
  if (matched === 0) {
    return 'Could not find ' + k.raws.map(playerFullName).join(', ') + ' on the "' + tabName + '" tab.';
  }
  return '';
}

// Recomputes the Owed total on the family row above each changed month row.
function updateCollectionsFamilyTotals(sheet, rows) {
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const iKey = h.indexOf('Key'), iOwed = h.indexOf('Owed'), iPaid = h.indexOf('Paid');
  if (iKey === -1 || iOwed === -1 || iPaid === -1) return;
  const typeAt = (i) => { const k = parseCollectionsRowKey(data[i][iKey]); return k ? k.type : ''; };
  const done = {};
  rows.forEach(row => {
    let f = row - 1;                                  // 0-based index into data
    while (f > 0 && typeAt(f) !== 'F') f--;
    if (f <= 0 || done[f]) return;
    done[f] = true;
    let owed = 0;
    for (let i = f + 1; i < data.length && typeAt(i) !== 'F'; i++) {
      if (typeAt(i) === 'M' && data[i][iPaid] !== true) owed += Number(data[i][iOwed]) || 0;
    }
    sheet.getRange(f + 1, iOwed + 1).setValue(owed);
    sheet.getRange(f + 1, 1, 1, COLLECTIONS_HEADERS.length - 1)
      .setFontColor(Math.round(owed * 100) > 0 ? '#021f3d' : '#9e9e9e');
  });
}

function setupCollections() {
  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction() === 'onCollectionsEdit') ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger('onCollectionsEdit').forSpreadsheet(COLLECTIONS_SHEET_ID).onEdit().create();
  const r = refreshCollections();
  Logger.log('Collections ready - Outreach stamps Last Contact, and Paid ticks the billing sheet.');
  return r;
}

// Outreach dropdown options, in escalation order. The last two mark a
// family as needing escalation in the weekly digest.
const OUTREACH_LEVELS = ['1st attempt', '2nd attempt', '3rd attempt', 'No response'];
const OUTREACH_ESCALATE = ['3rd attempt', 'No response'];

function getBillingReminderRecipients() {
  const sheet = SpreadsheetApp.openById(ROSTER_SHEET_ID).getSheetByName('Billing Reminders');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const emails = [];
  for (let i = 1; i < data.length; i++) {
    const email = (data[i][1] || '').toString().trim();
    if (email && email.indexOf('@') !== -1) emails.push(email);
  }
  return emails;
}

// Installable onEdit trigger on the BILLING spreadsheet: when a Charged?
// box is ticked, stamp today's date in Charged On (clear it when unticked).
// Also stamps Last Contact whenever the Outreach dropdown is set.
function onBillingEdit(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName().indexOf(' - Billing - ') === -1) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const chargedCol = headers.indexOf('Charged?') + 1;
    const chargedOnCol = headers.indexOf('Charged On') + 1;
    if (!chargedCol || !chargedOnCol) return;
    if (e.range.getColumn() !== chargedCol) return;

    // Outreach lives on the Collections sheet now - see onCollectionsEdit.
    const startRow = e.range.getRow();
    for (let r = 0; r < e.range.getNumRows(); r++) {
      const row = startRow + r;
      if (row === 1) continue;
      const checked = sheet.getRange(row, chargedCol).getValue() === true;
      const cell = sheet.getRange(row, chargedOnCol);
      cell.setValue(checked ? new Date() : '');
      cell.setNumberFormat('M/d/yyyy');
    }
  } catch (err) {
    // Never let the stamper break someone's edit
  }
}

// Monday email: a short heads-up with a link. The Collections sheet is the
// worklist, so the email does not repeat it - and it carries no phone
// numbers, which mail apps turn into links that are awkward to copy.
// Sends nothing when everyone has paid. Returns the number of families owing.
function sendUnchargedBillingDigest() {
  const recipients = getBillingReminderRecipients();
  if (recipients.length === 0) {
    Logger.log('No recipients on the "Billing Reminders" tab - collections email not sent.');
    return 0;
  }

  const r = refreshCollections();
  if (r.outstanding === 0) {
    Logger.log('Everyone has paid - no collections email sent.');
    return 0;
  }

  const url = 'https://docs.google.com/spreadsheets/d/' + COLLECTIONS_SHEET_ID + '/edit';
  const one = r.outstanding === 1;
  const families = r.outstanding + (one ? ' family owes' : ' families owe');
  const owed = '$' + r.owed.toFixed(2);

  const extras = [];
  if (r.escalate > 0) extras.push(r.escalate + (r.escalate === 1 ? ' needs' : ' need') + ' escalation');
  if (r.notContacted > 0) extras.push(r.notContacted + ' not contacted yet');
  if (r.noPhone > 0) extras.push(r.noPhone + ' with no phone number');

  const subject = 'IJTA Collections: ' + families + ' ' + owed;
  const html = '<div style="font-family:Arial,sans-serif;color:#333;max-width:520px;">' +
    '<h2 style="color:#021f3d;margin:0 0 8px;">IJTA Collections</h2>' +
    '<p style="font-size:16px;margin:0 0 4px;"><strong>' + families + '</strong> ' +
    '<strong>' + owed + '</strong></p>' +
    (extras.length ? '<p style="color:#666;margin:0 0 18px;">' + extras.join(' &middot; ') + '</p>' : '') +
    '<p style="margin:18px 0;"><a href="' + url + '" style="display:inline-block;background:#021f3d;' +
    'color:#fff;padding:11px 20px;border-radius:8px;text-decoration:none;font-weight:bold;">' +
    'Open the Collections sheet</a></p>' +
    '<p style="color:#888;font-size:12px;">Biggest balances are at the top. Tick Paid on a month ' +
    'once it has been charged &mdash; the billing sheet updates itself.</p></div>';

  MailApp.sendEmail({ to: recipients.join(','), subject: subject, htmlBody: html });
  return r.outstanding;
}

// One-time setup: creates the "Billing Reminders" recipients tab, the
// weekly digest trigger (Mon ~7am), and the Charged On date-stamper.
function setupBillingReminders() {
  const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  if (!ss.getSheetByName('Billing Reminders')) {
    const s = ss.insertSheet('Billing Reminders');
    s.appendRow(['Name', 'Email']);
    s.appendRow(['J.C.', 'jcdfreeman@gmail.com']);
    s.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#021f3d').setFontColor('white');
    s.setColumnWidth(1, 150);
    s.setColumnWidth(2, 260);
    s.setFrozenRows(1);
  }

  // Replace any existing triggers to avoid duplicates
  for (const t of ScriptApp.getProjectTriggers()) {
    const fn = t.getHandlerFunction();
    if (fn === 'sendUnchargedBillingDigest' || fn === 'onBillingEdit') {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger('sendUnchargedBillingDigest')
    .timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  ScriptApp.newTrigger('onBillingEdit')
    .forSpreadsheet(BILLING_SHEET_ID).onEdit().create();

  Logger.log('Billing reminders ready: weekly Monday digest + Charged On date-stamper. Add your shop manager to the "Billing Reminders" tab.');
}

// Menu handler: send the digest on demand.
// One-time setup, safe to re-run: installs the date-stamping trigger on the
// Collections sheet and builds the list for the first time. Kept separate from
// "Refresh Collections List" because only this one touches triggers.
function menuSetupCollections() {
  const ui = SpreadsheetApp.getUi();
  if (!COLLECTIONS_SHEET_ID) {
    ui.alert('No collections spreadsheet set. Create one, then paste its ID ' +
      'into COLLECTIONS_SHEET_ID at the top of this script and Save.');
    return;
  }
  let r;
  try {
    r = setupCollections();          // installs the trigger and builds the sheet
  } catch (err) {
    ui.alert('Could not set up the collections sheet.\n\n' + err.message +
      '\n\nIf that says permission, make sure the collections spreadsheet is ' +
      'shared with this account as an Editor.');
    return;
  }
  ui.alert('Collections sheet is ready.\n\n' + collectionsSummary(r) + '\n\n' +
    'Last Contact stamps itself when Outreach is set, and ticking Paid on a ' +
    'month ticks Charged? on the billing sheet.');
}

function collectionsSummary(r) {
  return r.outstanding + (r.outstanding === 1 ? ' family owes $' : ' families owe $') +
    (r.owed || 0).toFixed(2) + ', biggest balance first.';
}

function menuRefreshCollections() {
  const r = refreshCollections();
  SpreadsheetApp.getUi().alert(collectionsSummary(r) + '\n\n' +
    r.added + ' new month row' + (r.added !== 1 ? 's' : '') + ', ' +
    r.updated + ' carried over. Outreach and Notes were kept.');
}

function menuSendUnchargedDigest() {
  const ui = SpreadsheetApp.getUi();
  const recipients = getBillingReminderRecipients();
  if (recipients.length === 0) {
    ui.alert('No recipients yet. Run setupBillingReminders() once in the editor, then add emails to the "Billing Reminders" tab.');
    return;
  }
  const count = sendUnchargedBillingDigest();
  ui.alert(count === 0
    ? 'All charged - nothing outstanding, so no email was sent.'
    : 'Collections email sent (' + count + ' famil' + (count !== 1 ? 'ies' : 'y') +
      ' owing) to: ' + recipients.join(', '));
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
  const statusOverrides = getStatusOverrides(monthName);   // Lisa's Member/Guest corrections

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
    // Skip duplicate rows for the same kid on the same day
    if (clinicData[row.clinic].dates[dateStr].indexOf(row.playerName) === -1) {
      clinicData[row.clinic].dates[dateStr].push(row.playerName);
    }
    if (!(row.playerName in clinicData[row.clinic].playerStatus)) {
      clinicData[row.clinic].playerStatus[row.playerName] = row.status;   // first check-in wins
    }
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
    clinicSheet.getRange(1, 1).setValue(clinic + ' - Attendance Summary - ' + monthName);
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
        status: resolvePlayerStatus(statusOverrides, clinic, p, cd.playerStatus[p]),
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
  const statusOverrides = getStatusOverrides(monthName);   // Lisa's Member/Guest corrections

  // Get attendance data WITH coaches and no-attendee/cancelled markers
  const { rows: attendanceRows, sessionCoaches, sessionMarkers } =
    getAttendanceWithCoachesForMonth(billingMonth, billingYear);

  if (attendanceRows.length === 0 && Object.keys(sessionMarkers).length === 0) {
    Logger.log('No attendance data found for A/S summary: ' + monthName);
    return;
  }

  // Read config data from roster spreadsheet
  const coachRates = getCoachRates();
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
        coachesByDate: {},
        markersByDate: {}
      };
    }
    const cd = clinicData[row.clinic];

    if (!cd.dates[dateStr]) {
      cd.dates[dateStr] = [];
      cd.dateObjects[dateStr] = row.date;
    }
    // Skip duplicate rows for the same kid on the same day
    if (cd.dates[dateStr].indexOf(row.playerName) === -1) {
      cd.dates[dateStr].push(row.playerName);
    }
    if (!(row.playerName in cd.playerStatus)) cd.playerStatus[row.playerName] = row.status;   // first check-in wins

    // Map coaches to this clinic+date
    const sessionKey = dateStr + '|||' + row.clinic;
    if (sessionCoaches[sessionKey]) {
      cd.coachesByDate[dateStr] = sessionCoaches[sessionKey];
    }
  }

  // Fold in no-attendee / cancelled dates so they appear in the session
  // diary (zero players, zero dollars - informational only)
  for (const key in sessionMarkers) {
    const sep = key.indexOf('|||');
    const dateStr = key.substring(0, sep);
    const clinic = key.substring(sep + 3);
    if (!clinicData[clinic]) {
      clinicData[clinic] = { dates: {}, dateObjects: {}, playerStatus: {}, coachesByDate: {}, markersByDate: {} };
    }
    const cd = clinicData[clinic];
    if (!cd.markersByDate) cd.markersByDate = {};
    cd.markersByDate[dateStr] = sessionMarkers[key];
    if (!cd.dates[dateStr]) {
      cd.dates[dateStr] = [];
      const p = dateStr.split('/');
      cd.dateObjects[dateStr] = new Date(parseInt(p[2]), parseInt(p[0]) - 1, parseInt(p[1]));
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

    let heldCount = 0, noAttCount = 0, cancelledCount = 0;

    for (const dateStr of sortedDates) {
      const players = cd.dates[dateStr].sort();
      totalCheckIns += players.length;
      const marker = (cd.markersByDate && cd.markersByDate[dateStr]) || null;

      const dateCoaches = cd.coachesByDate[dateStr] || [];
      // Always show hours beside every coach, so a shift is never ambiguous
      const coachesDisplay = dateCoaches.map(c =>
        `${c.name} (${c.hours}h)`
      ).join(', ') || '(none recorded)';

      sheet.getRange(currentRow, 1).setValue(dateStr);
      sheet.getRange(currentRow, 2).setValue(players.length);
      if (players.length === 0 && marker) {
        // Diary row: session with nobody, or a cancellation
        sheet.getRange(currentRow, 3).setValue(marker);
        sheet.getRange(currentRow, 1, 1, 4).setFontColor('#999999').setFontStyle('italic');
        if (marker.indexOf('Cancelled') === 0) cancelledCount++;
        else noAttCount++;
      } else {
        sheet.getRange(currentRow, 3).setValue(coachesDisplay);
        sheet.getRange(currentRow, 4).setValue(players.join('; '));
        heldCount++;
      }

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
    let sessionsLabel = 'Total Sessions: ' + heldCount + ' held';
    if (noAttCount > 0) sessionsLabel += ', ' + noAttCount + ' no attendees';
    if (cancelledCount > 0) sessionsLabel += ', ' + cancelledCount + ' cancelled';
    sheet.getRange(currentRow, 1).setValue(sessionsLabel);
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
        status: resolvePlayerStatus(statusOverrides, clinic, p, cd.playerStatus[p]),
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

// ============================================================
// MASTER A/S SUMMARY (all clinics on one tab)
// ============================================================
// One consolidated page for the controller: a row per clinic showing
// revenue, staffing cost, and net profit, plus a grand total. Uses the
// exact same math as the per-clinic A/S tabs (sibling discounts applied,
// duplicate rows collapsed, cancelled/no-attendee dates excluded from
// dollars). Also lists total hours and cost per coach across all clinics.
// ============================================================

function generateMasterASSummary(monthOverride, yearOverride) {
  const now = new Date();
  const billingMonth = monthOverride || now.getMonth() + 1;
  const billingYear = yearOverride || now.getFullYear();

  const monthName = new Date(billingYear, billingMonth - 1, 1)
    .toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  const statusOverrides = getStatusOverrides(monthName);   // Lisa's Member/Guest corrections

  const { rows: attendanceRows, sessionCoaches, sessionMarkers } =
    getAttendanceWithCoachesForMonth(billingMonth, billingYear);

  if (attendanceRows.length === 0 && Object.keys(sessionMarkers).length === 0) {
    Logger.log('No data for master A/S summary: ' + monthName);
    return null;
  }

  const coachRates = getCoachRates();
  const siblingOverrides = getSiblingOverrides();

  // Gather per-clinic attendance, coaches, and markers
  const clinicData = {};
  const ensure = (clinic) => {
    if (!clinicData[clinic]) {
      clinicData[clinic] = { dates: {}, playerStatus: {}, coachesByDate: {}, markers: {} };
    }
    return clinicData[clinic];
  };

  for (const row of attendanceRows) {
    const dateStr = (row.date.getMonth() + 1) + '/' + row.date.getDate() + '/' + row.date.getFullYear();
    const cd = ensure(row.clinic);
    if (!cd.dates[dateStr]) cd.dates[dateStr] = [];
    if (cd.dates[dateStr].indexOf(row.playerName) === -1) cd.dates[dateStr].push(row.playerName);
    if (!(row.playerName in cd.playerStatus)) cd.playerStatus[row.playerName] = row.status;   // first check-in wins
    const sessionKey = dateStr + '|||' + row.clinic;
    if (sessionCoaches[sessionKey]) cd.coachesByDate[dateStr] = sessionCoaches[sessionKey];
  }
  for (const key in sessionMarkers) {
    const sep = key.indexOf('|||');
    ensure(key.substring(sep + 3)).markers[key.substring(0, sep)] = sessionMarkers[key];
  }

  // Roll up each clinic
  const clinicRows = [];
  const coachTotals = {}; // name -> { hours, cost }
  let gGross = 0, gDiscount = 0, gNet = 0, gStaffing = 0,
      gSessions = 0, gCheckIns = 0, gNoAtt = 0, gCancelled = 0;

  for (const clinic in clinicData) {
    const cd = clinicData[clinic];

    let held = 0, checkIns = 0;
    for (const dateStr in cd.dates) {
      if (cd.dates[dateStr].length > 0) { held++; checkIns += cd.dates[dateStr].length; }
    }
    let noAtt = 0, cancelled = 0;
    for (const dateStr in cd.markers) {
      if (String(cd.markers[dateStr]).indexOf('Cancelled') === 0) cancelled++;
      else noAtt++;
    }

    // Revenue (sibling discounts applied)
    const playerSessions = {};
    for (const dateStr in cd.dates) {
      for (const p of cd.dates[dateStr]) playerSessions[p] = (playerSessions[p] || 0) + 1;
    }
    const billing = buildClinicBillingRows(clinic,
      Object.keys(playerSessions).map(p => ({
        name: p, status: resolvePlayerStatus(statusOverrides, clinic, p, cd.playerStatus[p]), sessions: playerSessions[p]
      })), siblingOverrides);

    // Staffing cost from actual recorded hours
    let staffing = 0;
    for (const dateStr in cd.coachesByDate) {
      for (const c of cd.coachesByDate[dateStr]) {
        if (!c.name || c.name === 'No Staffing') continue;
        const rate = coachRates[c.name] || 0;
        const cost = c.hours * rate;
        staffing += cost;
        if (!coachTotals[c.name]) coachTotals[c.name] = { hours: 0, cost: 0 };
        coachTotals[c.name].hours += c.hours;
        coachTotals[c.name].cost += cost;
      }
    }

    clinicRows.push({
      clinic: clinic,
      sessions: held,
      noAtt: noAtt,
      cancelled: cancelled,
      players: Object.keys(playerSessions).length,
      checkIns: checkIns,
      gross: billing.gross,
      discount: billing.totalDiscount,
      net: billing.net,
      staffing: staffing,
      profit: billing.net - staffing
    });

    gGross += billing.gross; gDiscount += billing.totalDiscount; gNet += billing.net;
    gStaffing += staffing; gSessions += held; gCheckIns += checkIns;
    gNoAtt += noAtt; gCancelled += cancelled;
  }

  // Program progression order (youngest to oldest), not alphabetical.
  // Anything not listed falls to the end, alphabetically.
  const CLINIC_ORDER = ['Red Ball', 'Orange Ball', 'Green Ball', 'MS Yellow Ball', 'HS Yellow Ball', 'Bruno'];
  const orderOf = (name) => {
    const i = CLINIC_ORDER.indexOf(name);
    return i === -1 ? CLINIC_ORDER.length : i;
  };
  clinicRows.sort((a, b) =>
    orderOf(a.clinic) - orderOf(b.clinic) || a.clinic.localeCompare(b.clinic));

  // ---- Write the tab ----
  const ss = SpreadsheetApp.openById(BILLING_SHEET_ID);
  const tabName = 'MASTER Summary - ' + monthName;
  let sheet = ss.getSheetByName(tabName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(tabName, 0);  // keep it as the first tab

  sheet.getRange(1, 1).setValue('IJTA - All Clinics Summary - ' + monthName)
    .setFontWeight('bold').setFontSize(14);
  sheet.getRange(2, 1).setValue('Revenue net of sibling discounts, minus staffing cost.')
    .setFontColor('#666666').setFontStyle('italic');

  const headers = ['Clinic', 'Sessions', 'Players', 'Check-ins',
    'Gross Revenue', 'Sibling Discounts', 'Net Revenue', 'Staffing Cost', 'Net Profit'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#021f3d').setFontColor('white');

  let r = 5;
  for (const c of clinicRows) {
    sheet.getRange(r, 1, 1, headers.length).setValues([[
      c.clinic, c.sessions, c.players, c.checkIns,
      c.gross, c.discount > 0 ? -c.discount : 0, c.net, -c.staffing, c.profit
    ]]);
    r++;
  }

  const firstDataRow = 5;
  const dataCount = clinicRows.length;
  if (dataCount > 0) {
    sheet.getRange(firstDataRow, 5, dataCount, 5).setNumberFormat('$#,##0.00');
  }

  // Grand total
  const totalRow = firstDataRow + dataCount;
  sheet.getRange(totalRow, 1, 1, headers.length).setValues([[
    'TOTAL - ALL CLINICS', gSessions, '', gCheckIns,
    gGross, gDiscount > 0 ? -gDiscount : 0, gNet, -gStaffing, gNet - gStaffing
  ]]);
  sheet.getRange(totalRow, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#e8eef5').setBorder(true, true, true, true, false, false);
  sheet.getRange(totalRow, 5, 1, 5).setNumberFormat('$#,##0.00');
  sheet.getRange(totalRow, 9).setFontColor((gNet - gStaffing) >= 0 ? '#2e7d32' : '#c62828');

  // Bottom-line callout
  let row = totalRow + 2;
  sheet.getRange(row, 1).setValue('NET PROFIT').setFontWeight('bold').setFontSize(13);
  sheet.getRange(row, 2).setValue(gNet - gStaffing).setNumberFormat('$#,##0.00')
    .setFontWeight('bold').setFontSize(13)
    .setFontColor((gNet - gStaffing) >= 0 ? '#2e7d32' : '#c62828');
  row += 2;

  if (gNoAtt > 0 || gCancelled > 0) {
    let note = 'Also this month: ';
    const bits = [];
    if (gNoAtt > 0) bits.push(gNoAtt + ' session' + (gNoAtt !== 1 ? 's' : '') + ' with no attendees');
    if (gCancelled > 0) bits.push(gCancelled + ' cancelled session' + (gCancelled !== 1 ? 's' : ''));
    sheet.getRange(row, 1).setValue(note + bits.join(', ') + ' (no revenue or staffing cost).')
      .setFontColor('#666666').setFontStyle('italic');
    row += 2;
  }

  // Staffing breakdown across all clinics
  sheet.getRange(row, 1).setValue('STAFFING BY COACH (all clinics)').setFontWeight('bold').setFontSize(11);
  row++;
  const staffHeaders = ['Coach', 'Total Hours', 'Rate ($/hr)', 'Total Cost'];
  sheet.getRange(row, 1, 1, staffHeaders.length).setValues([staffHeaders])
    .setFontWeight('bold').setBackground('#e65100').setFontColor('white');
  row++;
  const coachNames = Object.keys(coachTotals).sort();
  const staffStart = row;
  for (const name of coachNames) {
    sheet.getRange(row, 1, 1, 4).setValues([[
      name, coachTotals[name].hours, coachRates[name] || 0, coachTotals[name].cost
    ]]);
    row++;
  }
  if (coachNames.length > 0) {
    sheet.getRange(staffStart, 3, coachNames.length, 2).setNumberFormat('$#,##0.00');
    sheet.getRange(row, 1).setValue('TOTAL STAFFING').setFontWeight('bold');
    sheet.getRange(row, 4).setValue(gStaffing).setNumberFormat('$#,##0.00').setFontWeight('bold');
  }

  sheet.setColumnWidth(1, 220);
  for (let c = 2; c <= 4; c++) sheet.setColumnWidth(c, 90);
  for (let c = 5; c <= 9; c++) sheet.setColumnWidth(c, 130);
  sheet.setFrozenRows(4);

  Logger.log('Master A/S summary for ' + monthName + ': net revenue $' + gNet +
    ', staffing $' + gStaffing + ', net profit $' + (gNet - gStaffing));
  return { net: gNet, staffing: gStaffing, profit: gNet - gStaffing, monthName: monthName };
}

function generateCurrentMonthMasterSummary() {
  const now = new Date();
  return generateMasterASSummary(now.getMonth() + 1, now.getFullYear());
}

function generateLastMonthMasterSummary() {
  const now = new Date();
  let month = now.getMonth();
  let year = now.getFullYear();
  if (month === 0) { month = 12; year--; }
  return generateMasterASSummary(month, year);
}

function menuCurrentMonthMaster() {
  const result = generateCurrentMonthMasterSummary();
  const ui = SpreadsheetApp.getUi();
  ui.alert(result
    ? 'Master summary generated for ' + result.monthName + '.\n\n' +
      'Net revenue: $' + result.net.toFixed(2) + '\n' +
      'Staffing: $' + result.staffing.toFixed(2) + '\n' +
      'NET PROFIT: $' + result.profit.toFixed(2)
    : 'No attendance data found for this month.');
}

function menuLastMonthMaster() {
  const result = generateLastMonthMasterSummary();
  const ui = SpreadsheetApp.getUi();
  ui.alert(result
    ? 'Master summary generated for ' + result.monthName + '.\n\n' +
      'Net revenue: $' + result.net.toFixed(2) + '\n' +
      'Staffing: $' + result.staffing.toFixed(2) + '\n' +
      'NET PROFIT: $' + result.profit.toFixed(2)
    : 'No attendance data found for last month.');
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
  generateMasterASSummary(monthOverride, yearOverride);
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
// ARCHIVE OLD REPORT TABS
// ============================================================
// Every month adds ~13 tabs to the billing spreadsheet (a billing tab and
// an A/S tab per clinic, plus the master summary). After a year that's
// 150+ tabs and the sheet becomes slow or impossible to open - the data
// stays fine, but the editor can't render it.
//
// This moves report tabs older than the months you're keeping into the
// archive spreadsheet (ARCHIVE_SHEET_ID), then removes them from the live
// one. Tabs are COPIED before being deleted, so the Charged? ticks and
// outreach history are preserved rather than lost.
//
// SETUP: create an empty spreadsheet, paste its ID into ARCHIVE_SHEET_ID
// at the top of this file. Then use IJTA Reports > Archive Old Reports.
// ============================================================

const MONTH_NAMES_FULL = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];

// Pulls the "August 2026" off the end of a report tab name.
// Returns { month (1-12), year } or null if it isn't a dated report tab.
function reportTabMonth(name) {
  const m = (name || '').match(/([A-Za-z]+)\s+(\d{4})\s*$/);
  if (!m) return null;
  const idx = MONTH_NAMES_FULL.indexOf(m[1]);
  if (idx === -1) return null;
  // Only tabs that look like generated reports
  if (name.indexOf(' - Billing - ') === -1 &&
      name.indexOf(' - A/S Summary - ') === -1 &&
      name.indexOf(' - Attendance - ') === -1 &&
      name.indexOf('MASTER Summary - ') !== 0) return null;
  return { month: idx + 1, year: parseInt(m[2], 10) };
}

// Which tabs are older than the retention window? monthsToKeep counts the
// current month, so 3 in August keeps June, July, August.
function findArchivableTabs(monthsToKeep) {
  const keep = monthsToKeep || 3;
  const now = new Date();
  const cutoff = new Date(now.getFullYear(), now.getMonth() - (keep - 1), 1);

  const ss = SpreadsheetApp.openById(BILLING_SHEET_ID);
  const out = [];
  for (const sheet of ss.getSheets()) {
    const info = reportTabMonth(sheet.getName());
    if (!info) continue;
    if (new Date(info.year, info.month - 1, 1) < cutoff) out.push(sheet);
  }
  return out;
}

// Counts players still unticked on the tabs about to be archived, so we
// never quietly move away money that hasn't been collected.
function countUnchargedOn(sheets) {
  let n = 0;
  for (const sheet of sheets) {
    if (sheet.getName().indexOf(' - Billing - ') === -1) continue;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;
    const chargedCol = data[0].indexOf('Charged?');
    const sessionsCol = data[0].indexOf('Sessions');
    if (chargedCol === -1 || sessionsCol === -1) continue;
    for (let i = 1; i < data.length; i++) {
      if (typeof data[i][sessionsCol] !== 'number') continue;   // skip summary rows
      if (!(data[i][0] || '').toString().trim()) continue;
      if (data[i][chargedCol] !== true) n++;
    }
  }
  return n;
}

// Copies each tab into the archive spreadsheet, then deletes it here.
// Returns { moved, skipped, names }.
function archiveOldReports(monthsToKeep) {
  if (!ARCHIVE_SHEET_ID) {
    throw new Error('No archive spreadsheet set. Create an empty Google Sheet, ' +
      'then paste its ID into ARCHIVE_SHEET_ID at the top of this script.');
  }
  const archiveSS = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  const liveSS = SpreadsheetApp.openById(BILLING_SHEET_ID);
  const existing = {};
  archiveSS.getSheets().forEach(s => { existing[s.getName()] = true; });

  const targets = findArchivableTabs(monthsToKeep);
  const moved = [], skipped = [];

  for (const sheet of targets) {
    const name = sheet.getName();
    try {
      if (existing[name]) {
        // Already archived (a previous run) - just remove the live copy
        liveSS.deleteSheet(sheet);
        moved.push(name);
        continue;
      }
      const copy = sheet.copyTo(archiveSS);
      copy.setName(name);
      existing[name] = true;
      liveSS.deleteSheet(sheet);
      moved.push(name);
    } catch (e) {
      skipped.push(name + ' (' + e.message + ')');
    }
  }

  // A brand-new spreadsheet starts with an empty "Sheet1" - clear it out
  try {
    if (moved.length > 0 && archiveSS.getSheets().length > 1) {
      const blank = archiveSS.getSheetByName('Sheet1');
      if (blank && blank.getLastRow() === 0) archiveSS.deleteSheet(blank);
    }
  } catch (e) { /* harmless */ }

  Logger.log('Archived ' + moved.length + ' tab(s); skipped ' + skipped.length);
  return { moved: moved.length, skipped: skipped.length, names: moved, errors: skipped };
}

function menuArchiveOldReports() {
  const ui = SpreadsheetApp.getUi();
  if (!ARCHIVE_SHEET_ID) {
    ui.alert('Set up the archive first:\n\n' +
      '1. Create an empty Google Sheet (name it "IJTA Billing Archive")\n' +
      '2. Copy its ID from the URL - the long string between /d/ and /edit\n' +
      '3. Paste it into ARCHIVE_SHEET_ID at the top of this script, and Save');
    return;
  }

  const KEEP = 3;
  const targets = findArchivableTabs(KEEP);
  if (targets.length === 0) {
    ui.alert('Nothing to archive - no report tabs older than the last ' + KEEP + ' months.');
    return;
  }

  const uncharged = countUnchargedOn(targets);
  let msg = 'Move ' + targets.length + ' report tab' + (targets.length !== 1 ? 's' : '') +
    ' older than the last ' + KEEP + ' months into the archive spreadsheet?\n\n' +
    'They are copied to the archive first, so nothing is lost - including ' +
    'Charged? ticks and outreach history.\n\n';
  if (uncharged > 0) {
    msg += 'HEADS UP: ' + uncharged + ' player' + (uncharged !== 1 ? 's are' : ' is') +
      ' still not marked as charged on those tabs. Archiving removes them from ' +
      'the weekly uncharged digest.\n\n';
  }
  msg += 'Oldest: ' + targets[0].getName();

  if (ui.alert('Archive Old Reports', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const result = archiveOldReports(KEEP);
  ui.alert('Archived ' + result.moved + ' tab' + (result.moved !== 1 ? 's' : '') + '.' +
    (result.skipped > 0 ? '\n\nSkipped ' + result.skipped + ':\n' + result.errors.join('\n') : '') +
    '\n\nReload this spreadsheet - it should open much faster now.');
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
    .addItem('Generate This Month - All Reports', 'menuCurrentMonthAll')
    .addItem('Generate Last Month - All Reports', 'menuLastMonthAll')
    .addSeparator()
    .addItem('Master Summary (All Clinics) - This Month', 'menuCurrentMonthMaster')
    .addItem('Master Summary (All Clinics) - Last Month', 'menuLastMonthMaster')
    .addSeparator()
    .addItem('Generate This Month - Billing Only', 'menuCurrentMonthBilling')
    .addItem('Generate This Month - A/S Summary Only', 'menuCurrentMonthAS')
    .addSeparator()
    .addItem('Generate Last Month - Billing Only', 'menuLastMonthBilling')
    .addItem('Generate Last Month - A/S Summary Only', 'menuLastMonthAS')
    .addSeparator()
    .addItem('Update Families List', 'menuUpdateFamilies')
    .addItem('Sync Contacts from Sign-Up Sheet', 'menuSyncContacts')
    .addItem('Set Up Collections Sheet (one time)', 'menuSetupCollections')
    .addItem('Refresh Collections List', 'menuRefreshCollections')
    .addItem('Send Uncharged Billing Digest Now', 'menuSendUnchargedDigest')
    .addSeparator()
    .addItem('Archive Old Reports', 'menuArchiveOldReports')
    .addToUi();

  ui.createMenu('Roll Reminders')
    .addItem('Send Me a Test Alert', 'menuTestReminder')
    .addItem('Check for Missing Rolls Now', 'menuCheckMissingNow')
    .addItem("Show What's Missing", 'menuShowMissing')
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

  // Create a new monthly trigger - runs on the 1st of each month
  ScriptApp.newTrigger('generateLastMonthAllReports')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();

  Logger.log('Monthly trigger set up - billing + A/S summary will run on the 1st of each month');
}

// ============================================================
// MISSING-ROLL REMINDERS
// ============================================================
// Flags clinics that were scheduled but have no roll logged,
// emails alerts (8pm same day + 7am next morning), and exposes
// the list to the app for an in-app warning badge.
//
// EVERYTHING is managed from tabs in the ROSTER spreadsheet
// (same place as Coaches & Clinic Config) - no code changes needed:
//   "Clinic Schedule"   - Clinic Name | Days | Owner (coach name)
//                         (alerts for a missing roll go to that clinic's
//                          owner, with the email looked up on the Coaches tab)
//   "Coaches"           - add an "Email" column so names resolve to emails
//   "Alert Recipients"  - Name | Email (admins: get EVERY alert; also the
//                         fallback when a clinic has no owner set)
//   "Reminder Settings" - "Reminders On?" | Yes/No  (master switch)
//
// ONE-TIME SETUP (run each once from the editor's Run button):
//   1. setupReminderTabs()         - creates the three tabs, pre-filled
//   2. setupMissingRollTriggers()  - schedules the 8pm + 7am checks
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

  // Locate the optional Owner / Starts / Ends columns by header text
  let ownerCol = -1, startsCol = -1, endsCol = -1;
  for (let c = 0; c < data[0].length; c++) {
    const h = (data[0][c] || '').toString().toLowerCase();
    if (ownerCol === -1 && h.indexOf('owner') !== -1) ownerCol = c;
    else if (startsCol === -1 && h.indexOf('start') !== -1) startsCol = c;
    else if (endsCol === -1 && h.indexOf('end') !== -1) endsCol = c;
  }

  // Accepts a real Date cell, "8/24/2026", "8/24/26" (2-digit year), or
  // "2026-08-24". A 2-digit year is treated as 20xx - without this, text
  // like "8/24/26" would parse as the year 26 AD and silently disable the
  // whole start-date filter.
  const toDate = (v) => {
    if (!v) return null;
    if (v instanceof Date) { const d = new Date(v); d.setHours(0, 0, 0, 0); return d; }
    const s = v.toString().trim();
    if (!s) return null;

    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);   // ISO
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);

    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/); // M/D/YY or M/D/YYYY
    if (m) {
      let year = parseInt(m[3], 10);
      if (m[3].length === 2) year += 2000;
      return new Date(year, +m[1] - 1, +m[2]);
    }

    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) { parsed.setHours(0, 0, 0, 0); return parsed; }
    return null;
  };

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
    // Optional date range: blank means "always". Lets a clinic start
    // meeting a new day mid-season without back-flagging every earlier
    // week as a missing roll.
    const starts = startsCol >= 0 ? toDate(data[i][startsCol]) : null;
    const ends = endsCol >= 0 ? toDate(data[i][endsCol]) : null;

    if (!schedule[clinic]) schedule[clinic] = { days: [], owners: [], rows: [] };
    schedule[clinic].rows.push({ days: days, starts: starts, ends: ends });
    days.forEach(d => {
      if (schedule[clinic].days.indexOf(d) === -1) schedule[clinic].days.push(d);
    });
    if (owner && schedule[clinic].owners.indexOf(owner) === -1) {
      schedule[clinic].owners.push(owner);
    }
  }
  return schedule;
}

// Does this clinic meet on this date, honouring each schedule row's
// optional Starts/Ends window?
function clinicMeetsOn(scheduleEntry, date) {
  if (!scheduleEntry || !scheduleEntry.rows) return false;
  const wd = date.getDay();
  for (const row of scheduleEntry.rows) {
    if (row.days.indexOf(wd) === -1) continue;
    if (row.starts && date < row.starts) continue;
    if (row.ends && date > row.ends) continue;
    return true;
  }
  return false;
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
        // ANY row counts as logged - including a "No Attendees" row.
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
      if (clinicMeetsOn(schedule[clinic], d) && !logged[dateStr + '|||' + clinic]) {
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
  // Subject must stay plain ASCII - emoji and special dashes get garbled by mail clients
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
  if (isTest) html += '<p style="color:#888;font-size:12px;">This is a test - your reminder system is working.</p>';
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
    Logger.log('Missing rolls found but no recipients resolved - check Alert Recipients and the Coaches Email column.');
  }
  return sent;
}

function checkMissingRollsEvening() { emailMissingRolls(true); }   // 8pm - includes today
function checkMissingRollsMorning() { emailMissingRolls(false); }  // 7am - through yesterday

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
// already exist). Fill them in afterward - the owner name must match
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

  Logger.log('Owner/Email columns ready - fill them in on the roster spreadsheet.');
}

// One-time upgrade: adds optional "Starts" and "Ends" date columns to the
// Clinic Schedule tab. Leave them blank for a day the clinic has always
// met. Fill "Starts" when a clinic BEGINS meeting a new day mid-season, so
// earlier weeks aren't retroactively flagged as missing rolls. A clinic can
// have several rows - e.g. "Red Ball | Wed" (blank) plus
// "Red Ball | Mon | Starts 8/24/2026".
function setupScheduleDateColumns() {
  const sheet = SpreadsheetApp.openById(ROSTER_SHEET_ID).getSheetByName('Clinic Schedule');
  if (!sheet) {
    Logger.log('No "Clinic Schedule" tab found - run setupReminderTabs() first.');
    return;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const has = (word) => headers.some(h => (h || '').toString().toLowerCase().indexOf(word) !== -1);

  let col = sheet.getLastColumn() + 1;
  if (!has('start')) {
    sheet.getRange(1, col).setValue('Starts (optional)')
      .setFontWeight('bold').setBackground('#021f3d').setFontColor('white');
    sheet.setColumnWidth(col, 150);
    col++;
  }
  if (!has('end')) {
    sheet.getRange(1, col).setValue('Ends (optional)')
      .setFontWeight('bold').setBackground('#021f3d').setFontColor('white');
    sheet.setColumnWidth(col, 150);
  }
  Logger.log('Starts/Ends columns ready on the Clinic Schedule tab.');
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
    ui.alert('\u2705 All caught up - no missing rolls.');
    return;
  }
  const sent = sendTargetedAlerts(missing, false);
  if (sent === 0) {
    ui.alert('Found ' + missing.length + ' missing roll(s), but no recipients could be resolved. Check the "Alert Recipients" tab and the Email column on the Coaches tab.');
  } else {
    ui.alert('Sent ' + sent + ' alert email(s) covering ' + missing.length + ' missing roll(s) - each person only gets their own clinics.');
  }
}

function menuShowMissing() {
  const ui = SpreadsheetApp.getUi();
  const missing = getCurrentMissingRolls();
  if (missing.length === 0) {
    ui.alert('\u2705 All caught up - no missing rolls.');
    return;
  }
  let txt = 'Missing rolls (scheduled but not logged):\n\n';
  for (const m of missing) txt += '\u2022 ' + m.day + ' ' + m.date + ' - ' + m.clinic + '\n';
  ui.alert(txt);
}
