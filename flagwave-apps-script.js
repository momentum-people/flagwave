// ═══════════════════════════════════════════════════════════════════════
// FLAG WAVE SYSTEM — Google Apps Script Backend
// Launch North Attleboro · World Cup Kickoff Party · June 12, 2026
// Drone Shot: 6:00 PM
// ═══════════════════════════════════════════════════════════════════════
//
// ── SETUP INSTRUCTIONS ──────────────────────────────────────────────────
//
// STEP 1: Create a new Google Sheet
//   - Go to sheets.google.com → create a blank sheet
//   - Copy the Sheet ID from the URL:
//     https://docs.google.com/spreadsheets/d/[THIS_PART_HERE]/edit
//   - Paste it as SPREADSHEET_ID below
//
// STEP 2: Paste this file into Apps Script
//   - In your Google Sheet: Extensions → Apps Script
//   - Delete the default code, paste this entire file
//   - Save (Ctrl+S)
//
// STEP 3: Run setup once
//   - In the Apps Script editor, select "setupSheet" from the function dropdown
//   - Click Run — this creates all sheets with correct headers and formatting
//   - Accept any permissions prompts
//
// STEP 4: Deploy as a Web App
//   - Click Deploy → New Deployment
//   - Type: Web App
//   - Execute as: Me
//   - Who has access: Anyone
//   - Click Deploy → copy the Web App URL
//
// STEP 5: Add the Web App URL to your HTML files
//   - In flagwave-signup.html:  replace https://script.google.com/macros/s/AKfycbyr3B8Xw5gQWnRvSCb-weOnRYWVC7hEznaX0ChM1HBPGTGhVBP3dywity6tBhsCUx9BiA/exec
//   - In flagwave-captain.html: replace https://script.google.com/macros/s/AKfycbyr3B8Xw5gQWnRvSCb-weOnRYWVC7hEznaX0ChM1HBPGTGhVBP3dywity6tBhsCUx9BiA/exec
//   - Also replace YOUR_ZONE_MAP_URL_HERE with a link to your zone map image
//
// STEP 6: Test
//   - Open flagwave-signup.html, fill out the form, submit
//   - Check your Google Sheet — a new row should appear in "Signups"
//   - Open flagwave-captain.html, log in to Zone 1, the guest should appear
//
// ── REDEPLOYING AFTER CHANGES ────────────────────────────────────────────
//   Any time you edit this script, you must redeploy:
//   Deploy → Manage Deployments → Edit → New Version → Deploy
//   The URL stays the same — no need to update the HTML files.
//
// ═══════════════════════════════════════════════════════════════════════

// ── CONFIGURATION ───────────────────────────────────────────────────────

const SPREADSHEET_ID = '1vwF3u84AMhtuCat5vRIpHxG-WP_jRluCguCLG32Vhu0';

// Sheet names — do not change these after running setupSheet()
const SHEET_SIGNUPS   = 'Signups';
const SHEET_DASHBOARD = 'Dashboard';
const SHEET_RAFFLE    = 'Raffle Export';

// Zone setup
const TOTAL_ZONES    = 6;
const BEARERS_PER_ZONE = 8;  // first 8 check-ins per zone = flag bearers (3x entries)
const MAX_PER_ZONE   = 40;   // soft cap per zone for the optimizer — adjust based on expected turnout

// Zone names
const ZONE_NAMES = {
  1: 'Zone 1 - Trampoline Courts',
  2: 'Zone 2 - Arcade',
  3: 'Zone 3 - Main Entrance / Lobby',
  4: 'Zone 4 - Lucky Putt / Bowling',
  5: 'Zone 5 - Krave Cafe',
  6: 'Zone 6 - Bar Hops'
};

// Event details
const DRONE_SHOT_TIME = '6:00 PM';
const CHECKIN_DEADLINE = '5:50 PM';
const EVENT_DATE = 'June 12, 2026';
const EVENT_NAME = 'World Cup Kickoff Party';
const VENUE_NAME = 'Launch North Attleboro';
const ZONE_MAP_URL = 'https://drive.google.com/file/d/14Mye2sQQXogOX_ASgKY3d_7KqhlFJ1HJ/view?usp=drive_link';

// ── COLUMN INDEXES (0-based, matches sheet headers) ─────────────────────
const COL = {
  TIMESTAMP:    0,
  FIRST_NAME:   1,
  LAST_NAME:    2,
  PHONE:        3,
  EMAIL:        4,
  ZONE:         5,
  STATUS:       6,   // pending | present | bearer
  CAPTAIN:      7,
  CHECKIN_TIME: 8,
  SIGNUP_TIME:  9,
  // Team preference columns removed — zone assignment is round-robin only
};

// ═══════════════════════════════════════════════════════════════════════
// ENTRY POINTS
// ═══════════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'signup')       return handleSignup(data);
    if (action === 'updateStatus') return handleStatusUpdate(data);
    if (action === 'getGuests')    return handleGetGuests(data);

    return jsonResponse({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return jsonResponse({ success: false, error: err.message });
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'getGuests')    return handleGetGuests(e.parameter);
    if (action === 'getRaffleList') return handleGetRaffleList();
    if (action === 'getDashboard') return handleGetDashboard();
    return jsonResponse({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ═══════════════════════════════════════════════════════════════════════
// SIGNUP HANDLER
// ═══════════════════════════════════════════════════════════════════════

function handleSignup(data) {
  const sheet = getOrCreateSheet(SHEET_SIGNUPS);
  ensureSignupHeaders(sheet);

  // Prevent duplicate signups by phone number
  if (data.phone) {
    const existing = findRowByPhone(sheet, data.phone);
    if (existing > 0) {
      const rows = sheet.getDataRange().getValues();
      const existingZone = rows[existing - 1][COL.ZONE];
      return jsonResponse({
        success: true,
        duplicate: true,
        zone: existingZone,
        zoneDesc: ZONE_NAMES[existingZone],
        message: 'Already registered'
      });
    }
  }

  // Assign zone — round-robin by current total count for even distribution
  const currentCount = Math.max(0, sheet.getLastRow() - 1);
  const zone = (currentCount % TOTAL_ZONES) + 1;

  // Is there still flag bearer capacity in this zone?
  const zoneBearer = getZoneBearerCount(sheet, zone);
  const potentialBearer = zoneBearer < BEARERS_PER_ZONE;

  const timestamp = new Date();
  const timeDisplay = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'h:mm a');

  // Build the row — 10 columns
  const newRow = new Array(10).fill('');
  newRow[COL.TIMESTAMP]   = timestamp;
  newRow[COL.FIRST_NAME]  = data.firstName  || '';
  newRow[COL.LAST_NAME]   = data.lastName   || '';
  newRow[COL.PHONE]       = data.phone      || '';
  newRow[COL.EMAIL]       = data.email      || '';
  newRow[COL.ZONE]        = zone;
  newRow[COL.STATUS]      = 'pending';
  newRow[COL.CAPTAIN]     = '';
  newRow[COL.CHECKIN_TIME]= '';
  newRow[COL.SIGNUP_TIME] = timeDisplay;


  sheet.appendRow(newRow);
  updateDashboard();

  // Send confirmation email if provided
  const confirmationSent = sendConfirmationEmail(data, zone, potentialBearer);

  return jsonResponse({
    success: true,
    zone: zone,
    zoneName: ZONE_NAMES[zone],
    potentialBearer: potentialBearer,
    confirmationSent: confirmationSent
  });
}

// ═══════════════════════════════════════════════════════════════════════
// STATUS UPDATE HANDLER (Zone Captain check-in app)
// ═══════════════════════════════════════════════════════════════════════

function handleStatusUpdate(data) {
  const sheet = getOrCreateSheet(SHEET_SIGNUPS);
  const rows = sheet.getDataRange().getValues();

  // Find by row ID (preferred — sent by captain app) or phone as fallback
  const targetId    = parseInt(data.id)  || 0;
  const targetPhone = normalizePhone(data.phone || '');

  for (let i = 1; i < rows.length; i++) {
    const rowId    = i;
    const rowPhone = normalizePhone(String(rows[i][COL.PHONE]));
    const matched  = (targetId && rowId === targetId) || (targetPhone && rowPhone === targetPhone);

    if (matched) {
      const newStatus = data.status || 'present';
      sheet.getRange(i + 1, COL.STATUS + 1).setValue(newStatus);
      sheet.getRange(i + 1, COL.CAPTAIN + 1).setValue(data.captain || '');
      if (newStatus !== 'pending') {
        const checkinTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'h:mm a');
        sheet.getRange(i + 1, COL.CHECKIN_TIME + 1).setValue(checkinTime);
      }
      updateDashboard();
      return jsonResponse({ success: true, updated: true, row: i + 1 });
    }
  }

  return jsonResponse({ success: false, error: 'Guest not found' });
}

// ═══════════════════════════════════════════════════════════════════════
// GET GUESTS FOR ZONE (Zone Captain check-in app)
// ═══════════════════════════════════════════════════════════════════════

function handleGetGuests(params) {
  const zone = parseInt(params.zone);
  if (!zone) return jsonResponse({ success: false, error: 'Zone required' });

  const sheet = getOrCreateSheet(SHEET_SIGNUPS);
  const rows  = sheet.getDataRange().getValues();
  const guests = [];

  for (let i = 1; i < rows.length; i++) {
    if (parseInt(rows[i][COL.ZONE]) === zone) {
      guests.push({
        id:         i,
        firstName:  rows[i][COL.FIRST_NAME],
        lastName:   rows[i][COL.LAST_NAME],
        name:       rows[i][COL.FIRST_NAME] + ' ' + rows[i][COL.LAST_NAME],
        phone:      rows[i][COL.PHONE],
        zone:       rows[i][COL.ZONE],
        status:     rows[i][COL.STATUS] || 'pending',
        captain:    rows[i][COL.CAPTAIN],
        checkinTime:rows[i][COL.CHECKIN_TIME],
        signupTime: rows[i][COL.SIGNUP_TIME]
      });
    }
  }

  // Sort: bearers first, then present, then pending; alphabetical within each
  guests.sort((a, b) => {
    const order = { bearer: 0, present: 1, pending: 2 };
    const diff = (order[a.status] || 2) - (order[b.status] || 2);
    if (diff !== 0) return diff;
    return a.lastName.localeCompare(b.lastName);
  });

  return jsonResponse({ success: true, guests: guests, zone: zone, zoneName: ZONE_NAMES[zone] });
}

// ═══════════════════════════════════════════════════════════════════════
// RAFFLE LIST EXPORT
// Returns checked-in names formatted for the raffle app.
// Flag bearers are marked with * (app weights them 3x automatically).
// Zone is included in parentheses so the raffle app can display it.
// ═══════════════════════════════════════════════════════════════════════

function handleGetRaffleList() {
  const sheet = getOrCreateSheet(SHEET_SIGNUPS);
  const rows  = sheet.getDataRange().getValues();
  const names = [];

  for (let i = 1; i < rows.length; i++) {
    const status = rows[i][COL.STATUS];
    if (status === 'present' || status === 'bearer') {
      const first    = rows[i][COL.FIRST_NAME];
      const last     = rows[i][COL.LAST_NAME];
      const zoneName = ZONE_NAMES[rows[i][COL.ZONE]] || ('Zone ' + rows[i][COL.ZONE]);
      const bearer   = status === 'bearer';
      // Format: "First Last* (Zone N - Location)"
      names.push(`${first} ${last}${bearer ? '*' : ''} (${zoneName})`);
    }
  }

  const bearerCount  = names.filter(n => n.includes('*')).length;
  const regularCount = names.length - bearerCount;

  // Also write to the Raffle Export sheet for easy copy-paste
  writeRaffleExportSheet(names);

  return jsonResponse({
    success:  true,
    names:    names,
    total:    names.length,
    bearers:  bearerCount,
    regular:  regularCount,
    entries:  regularCount + (bearerCount * 3)
  });
}

// ═══════════════════════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════════════════════

function handleGetDashboard() {
  updateDashboard();
  const sheet = getOrCreateSheet(SHEET_DASHBOARD);
  const data  = sheet.getDataRange().getValues();
  return jsonResponse({ success: true, dashboard: data });
}

function updateDashboard() {
  try {
    const signups  = getOrCreateSheet(SHEET_SIGNUPS);
    const dash     = getOrCreateSheet(SHEET_DASHBOARD);
    const rows     = signups.getDataRange().getValues();

    // Aggregate per zone
    const stats = {};
    for (let z = 1; z <= TOTAL_ZONES; z++) {
      stats[z] = { zone: z, name: ZONE_NAMES[z], signups: 0, present: 0, bearers: 0, pending: 0 };
    }

    for (let i = 1; i < rows.length; i++) {
      const z      = parseInt(rows[i][COL.ZONE]);
      const status = rows[i][COL.STATUS] || 'pending';
      if (!stats[z]) continue;
      stats[z].signups++;
      if (status === 'bearer')  { stats[z].bearers++; stats[z].present++; }
      else if (status === 'present') stats[z].present++;
      else stats[z].pending++;
    }

    // Write dashboard
    dash.clearContents();
    dash.getRange(1, 1).setValue('FLAG WAVE — LIVE DASHBOARD');
    dash.getRange(1, 1).setFontWeight('bold').setFontSize(14);
    dash.getRange(2, 1).setValue('Last updated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a'));

    const headers = ['Zone', 'Location', 'Total Signed Up', 'Checked In', 'Flag Bearers', 'Still Pending', 'Bearer Slots Left'];
    dash.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

    const dataRows = [];
    let totalSignups = 0, totalPresent = 0, totalBearers = 0;

    for (let z = 1; z <= TOTAL_ZONES; z++) {
      const s = stats[z];
      const bearerSlotsLeft = Math.max(0, BEARERS_PER_ZONE - s.bearers);
      dataRows.push([z, s.name, s.signups, s.present, s.bearers, s.pending, bearerSlotsLeft]);
      totalSignups += s.signups;
      totalPresent += s.present;
      totalBearers += s.bearers;
    }

    dash.getRange(5, 1, dataRows.length, headers.length).setValues(dataRows);

    // Totals row
    const totalRow = 5 + dataRows.length;
    dash.getRange(totalRow, 1, 1, headers.length).setValues([[
      'TOTAL', '', totalSignups, totalPresent, totalBearers,
      totalSignups - totalPresent,
      Math.max(0, (BEARERS_PER_ZONE * TOTAL_ZONES) - totalBearers)
    ]]).setFontWeight('bold');

    // Raffle entries count
    dash.getRange(totalRow + 2, 1).setValue('Total raffle entries: ' + (( totalSignups - totalBearers) + (totalBearers * 3)));
    dash.getRange(totalRow + 3, 1).setValue('  Regular (1x): ' + (totalPresent - totalBearers));
    dash.getRange(totalRow + 4, 1).setValue('  Flag Bearers (3x): ' + totalBearers + ' × 3 = ' + (totalBearers * 3) + ' entries');

    // Auto-size columns
    dash.autoResizeColumns(1, headers.length);

  } catch (err) {
    Logger.log('updateDashboard error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// EMAIL CONFIRMATIONS
// ═══════════════════════════════════════════════════════════════════════

function sendConfirmationEmail(data, zone, potentialBearer) {
  try {
    if (!data.email) return false;

    const zoneName  = ZONE_NAMES[zone];
    const bearerMsg = potentialBearer
      ? '⭐ BONUS: Get to your zone early to become an official Flag Bearer and earn 3× raffle entries!'
      : 'Show up to your zone and wave your flag to earn your raffle entry.';

    const subject = `🚩 You\'re in the Flag Wave! ${zoneName} — ${EVENT_DATE}`;

    const body =
`Hey ${data.firstName}!

You\'re officially registered for the Launch Flag Wave — here\'s everything you need:

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
YOUR ZONE ASSIGNMENT
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📍 ${zoneName}
⏰ Be there by: 5:50 PM (10 min before the shot)
📅 ${EVENT_DATE} at ${VENUE_NAME}

${bearerMsg}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
HOW IT WORKS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. Head to ${zoneName} before 5:50 PM
2. Check in with your Zone Captain
3. Grab your flag and wave it with everything you've got for the drone shot at ${DRONE_SHOT_TIME}
4. Stay after the shot — we're drawing the raffle winner live on the big screen!

🏆 PRIZE: 2 Front Row tickets to a World Cup match

📍 VIEW THE ZONE MAP: ${ZONE_MAP_URL}
(Open this to see exactly where your zone is located in the venue)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
IMPORTANT: You must be present at your zone during the 6:00 PM drone shot to be entered in the raffle. No-shows will not be included.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

See you out there!

— The Launch Team
${VENUE_NAME}
${EVENT_NAME} · ${EVENT_DATE}
`;

    GmailApp.sendEmail(data.email, subject, body);
    return true;
  } catch (err) {
    Logger.log('sendConfirmationEmail error: ' + err.message);
    return false;
  }
}

// ═══════════════════════════════════════════════════════════════════════
// RAFFLE EXPORT SHEET
// Writes a ready-to-copy list to the "Raffle Export" sheet.
// At 6:05 PM: open this sheet, copy column A, paste into the raffle app.
// ═══════════════════════════════════════════════════════════════════════

function writeRaffleExportSheet(names) {
  try {
    const sheet = getOrCreateSheet(SHEET_RAFFLE);
    sheet.clearContents();

    sheet.getRange(1, 1).setValue('RAFFLE EXPORT — copy column A into the raffle app').setFontWeight('bold');
    sheet.getRange(2, 1).setValue('Generated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d h:mm a'));
    sheet.getRange(3, 1).setValue('Total entries: ' + names.length + '  (* = flag bearer, 3x entries)');
    sheet.getRange(5, 1).setValue('Name (paste this column into the raffle app)').setFontWeight('bold');

    if (names.length > 0) {
      sheet.getRange(6, 1, names.length, 1).setValues(names.map(n => [n]));
    }

    sheet.autoResizeColumn(1);
  } catch (err) {
    Logger.log('writeRaffleExportSheet error: ' + err.message);
  }
}

// Run this manually just before the raffle to generate the export sheet
function generateRaffleExport() {
  handleGetRaffleList();
  try {
    SpreadsheetApp.getUi().alert('Raffle Export Generated!\n\nOpen the "Raffle Export" sheet, copy column A starting from row 6, and paste into the raffle app.');
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════════
// SETUP — run once after pasting this script
// ═══════════════════════════════════════════════════════════════════════

function setupSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── Signups sheet ──────────────────────────────────────────────────
  let signups = ss.getSheetByName(SHEET_SIGNUPS);
  if (!signups) signups = ss.insertSheet(SHEET_SIGNUPS);
  else signups.clearContents();

  const signupHeaders = [
    'Timestamp', 'First Name', 'Last Name', 'Phone', 'Email',
    'Zone', 'Status', 'Captain', 'Checked In Time', 'Signup Time',
    'Team Pref 1', 'Team Pref 2', 'Optimizer Run'
  ];
  signups.getRange(1, 1, 1, signupHeaders.length).setValues([signupHeaders]);
  signups.getRange(1, 1, 1, signupHeaders.length).setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#ffffff');
  signups.setFrozenRows(1);
  signups.setColumnWidth(1, 160);   // Timestamp
  signups.setColumnWidth(2, 100);   // First Name
  signups.setColumnWidth(3, 100);   // Last Name
  signups.setColumnWidth(4, 130);   // Phone
  signups.setColumnWidth(5, 180);   // Email
  signups.setColumnWidth(6, 60);    // Zone
  signups.setColumnWidth(7, 80);    // Status
  signups.setColumnWidth(8, 120);   // Captain
  signups.setColumnWidth(9, 110);   // Checked In Time
  signups.setColumnWidth(10, 90);   // Signup Time


  // ── Dashboard sheet ────────────────────────────────────────────────
  let dash = ss.getSheetByName(SHEET_DASHBOARD);
  if (!dash) dash = ss.insertSheet(SHEET_DASHBOARD);

  // ── Raffle Export sheet ────────────────────────────────────────────
  let raffle = ss.getSheetByName(SHEET_RAFFLE);
  if (!raffle) raffle = ss.insertSheet(SHEET_RAFFLE);

  // Move Signups to first tab
  ss.setActiveSheet(signups);
  ss.moveActiveSheet(1);

  // Trigger initial dashboard
  updateDashboard();

  Logger.log('✅ Setup complete! All sheets created.');
  try {
    SpreadsheetApp.getUi().alert('✅ Setup complete!\n\nSheets created:\n• Signups\n• Dashboard\n• Raffle Export\n\nNext step: Deploy as a Web App (Deploy → New Deployment).');
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════════════════

function getOrCreateSheet(name) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureSignupHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    const headers = [
      'Timestamp', 'First Name', 'Last Name', 'Phone', 'Email',
      'Zone', 'Status', 'Captain', 'Checked In Time', 'Signup Time',
      'Team Pref 1', 'Team Pref 2', 'Optimizer Run'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, signupHeaders.length).setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}

function getZoneBearerCount(sheet, zone) {
  const rows = sheet.getDataRange().getValues();
  return rows.filter((r, i) => i > 0 && parseInt(r[COL.ZONE]) === zone && r[COL.STATUS] === 'bearer').length;
}

function findRowByPhone(sheet, phone) {
  const normalized = normalizePhone(phone);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (normalizePhone(String(rows[i][COL.PHONE])) === normalized) return i + 1;
  }
  return -1;
}

function normalizePhone(phone) {
  return String(phone).replace(/\D/g, '');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
