/**
 * FNRG Portal — Self-Initializing Google Apps Script Backend (v5)
 * Page 1: Thursday Calling — admin, assignments, dated rounds & call counts
 *
 * HOW TO USE:
 * 1. Extensions -> Apps Script -> replace all code with this file. Save.
 * 2. Deploy -> Manage deployments -> edit pencil -> Version: New version -> Deploy.
 *    (Execute as: Me, Who has access: Anyone)
 *
 * Sheet layout ("Thursday Calling"):
 *   A Name | B Phone Number | C W/S | D Sessions Attended | E No. of Calls |
 *   F Assigned To | G, H, I ... date columns (e.g. "Jul 10", "Jul 17")
 *
 *   - The LAST date column is the active round; status submissions write there.
 *   - "No. of Calls" (E) is a live formula: counts date cells that are filled
 *     and not "Not Done" — stays correct even when you edit the sheet by hand.
 *   - The script AUTO-MIGRATES older sheets: inserts the E column, adds Role /
 *     Assigned To, creates the first date column, styles headers.
 *
 * Access rules:
 *   - Admin: sees ALL contacts; edits W/S, Status, Assigned To.
 *   - User:  sees ONLY contacts assigned to them; edits W/S and Status.
 *   - Sessions and Phone Number can ONLY be edited directly in the sheet.
 *   - Status accepts list values or free text (the portal's "Others" option).
 */

// ---------------------------------------------------------------
// Configuration & Seed Data
// ---------------------------------------------------------------

var SHEET_USERS = "Users";
var SHEET_CALLING = "Thursday Calling";
var SHEET_FESTIVAL = "Festival Promotions";

var CACHE_KEY_CALLING = "calling_v10";
var CACHE_KEY_USERS = "users_v3";
var CACHE_KEY_SETUP = "setup_v14";
var CACHE_KEY_VALIDATION = "validation_sig_v10";
var CACHE_SECONDS = 60;

var WS_OPTIONS = ["W", "S", "NA"];
var ROLE_ADMIN = "admin";
var ROLE_USER = "user";

// Columns: 1 Name, 2 Phone, 3 W/S, 4 Sessions, 5 No. of Calls, 6 Cultivated By, 7 Assigned To
var FIXED_COLS = 7; // date columns start at column 8 (H)
var COL_WS = 3;
var COL_CALLS = 5;
var COL_CULTIVATED = 6;
var COL_ASSIGNED = 7;

var STATUS_DEFAULT = "Not Done";
var STATUS_OPTIONS = [
  "Not Done",
  "Don't Call him again",
  "Joining the session",
  "Next Week will join",
  "Out of station",
  "evening Shift",
  "Busy",
  "Will come for Saturday",
  "Wrong Number",
  "Sunday Available",
  "Will try to attend",
  "Yet To Call",
  "Didn't Receive, Sent in WhatsApp",
  "Out of Network Coverage",
  "Shifted to Home town",
  "Only Online session",
  "Others"
];

var SAMPLE_USERS = [
  ["Sai Vardhan", "1111111111", "user"],
  ["Dushmanth", "9876543211", "user"],
  ["Srinivas", "9876543212", "user"],
  ["SNKD", "9999999999", "admin"]
];

var ADMIN_SEED = { name: "SNKD", phone: "9999999999" };

var SAMPLE_CONTACTS = [
  ["Aarav Sharma",   "9812345601", "W",  5, "", "", ""],
  ["Ravi Teja",      "9812345602", "S",  2, "", "", ""],
  ["Karthik Reddy",  "9812345603", "NA", 0, "", "", ""],
  ["Manoj Kumar",    "9812345604", "W",  7, "", "", ""],
  ["Pranav Varma",   "9812345605", "S",  3, "", "", ""],
  ["Harsha Vardhan", "9812345606", "W",  1, "", "", ""],
  ["Nikhil Rao",     "9812345607", "NA", 0, "", "", ""],
  ["Abhinav Gupta",  "9812345608", "S",  4, "", "", ""],
  ["Vamsi Krishna",  "9812345609", "W",  6, "", "", ""],
  ["Rohan Das",      "9812345610", "S",  2, "", "", ""]
];

// ---------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Keep only digits, compare by last 10 (handles +91, spaces, dashes, etc.)
function normalizePhone(value) {
  var digits = String(value === null || value === undefined ? "" : value).replace(/\D/g, "");
  return digits.length > 10 ? digits.slice(-10) : digits;
}

function styleHeader(range, color) {
  range.setFontWeight("bold").setBackground(color).setFontColor("#ffffff");
}

/**
 * Sheets auto-converts headers like "11 July" into real Date values.
 * Always show them as a clean "Jul 11" instead of the raw date string.
 */
function headerLabel(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "MMM d");
  }
  return String(value).trim();
}

/**
 * Finds the active (last) date column of the calling sheet.
 * Returns { activeCol, activeDate } — activeCol is -1 if none exists.
 */
function getActiveDateColumn(sheet) {
  var lastCol = sheet.getLastColumn();
  var result = { activeCol: -1, activeDate: "" };
  var sheetName = sheet.getName();
  var fixed = (sheetName === SHEET_FESTIVAL) ? 8 : 7;
  if (lastCol <= fixed) return result;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var c = fixed; c < headers.length; c++) {
    var title = headerLabel(headers[c]);
    if (title !== "") {
      result.activeCol = c + 1; // 1-indexed
      result.activeDate = title;
    }
  }
  return result;
}

/**
 * Creates tabs on a blank sheet, or migrates an older sheet in place.
 * Header checks only run once per cache window, so requests stay fast.
 */
function ensureSetup() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var cache = CacheService.getScriptCache();
  if (cache.get(CACHE_KEY_SETUP)) return doc;

  // Everything below is best-effort sheet setup/cosmetics. It must NEVER
  // crash a request (e.g. Sheets "typed columns"/Tables reject some writes),
  // so the whole block is guarded and the API keeps serving data regardless.
  try {

  // --- Users tab: create or upgrade ---
  var users = doc.getSheetByName(SHEET_USERS);
  if (!users) {
    users = doc.insertSheet(SHEET_USERS);
    try { users.getRange(1, 1, 1, 8).setValues([["Name", "Phone Number", "Role", "Call Limit", "Auto Assign", "Festival", "From Date", "To Date"]]); } catch(e){}
    try { styleHeader(users.getRange(1, 1, 1, 8), "#0f766e"); } catch(e){}
    try { users.setFrozenRows(1); } catch(e){}
    try { users.getRange(2, 2, 1000, 1).setNumberFormat("@"); } catch(e){}
    try { users.getRange(2, 4, 1000, 1).setNumberFormat("#"); } catch(e){}
    try { users.getRange(2, 1, SAMPLE_USERS.length, 3).setValues(SAMPLE_USERS); } catch(e){}
  } else {
    try {
      var roleHeader = String(users.getRange(1, 3).getValue()).trim();
      if (roleHeader !== "Role") {
        users.getRange(1, 3).setValue("Role");
        try { styleHeader(users.getRange(1, 3), "#0f766e"); } catch(e){}
        var uLast = users.getLastRow();
        if (uLast > 1) {
          var roleRange = users.getRange(2, 3, uLast - 1, 1);
          var roleValues = roleRange.getValues();
          for (var r = 0; r < roleValues.length; r++) {
            if (String(roleValues[r][0]).trim() === "") roleValues[r][0] = ROLE_USER;
          }
          roleRange.setValues(roleValues);
        }
      }
    } catch(e){}
    try {
      // Migrate: check D header is Call Limit
      var limitHeader = String(users.getRange(1, 4).getValue()).trim();
      if (limitHeader !== "Call Limit") {
        users.getRange(1, 4).setValue("Call Limit");
        try { styleHeader(users.getRange(1, 4), "#0f766e"); } catch(e){}
        try { users.getRange(2, 4, 1000, 1).setNumberFormat("#"); } catch(e){}
      }
    } catch(e){}
    try {
      // Migrate: check E header is Auto Assign
      var autoAssignHeader = String(users.getRange(1, 5).getValue()).trim();
      if (autoAssignHeader !== "Auto Assign") {
        users.getRange(1, 5).setValue("Auto Assign");
        try { styleHeader(users.getRange(1, 5), "#0f766e"); } catch(e){}
      }
    } catch(e){}
    try {
      // Migrate: check F header is Festival
      var festivalHeader = String(users.getRange(1, 6).getValue()).trim();
      if (festivalHeader !== "Festival") {
        users.getRange(1, 6).setValue("Festival");
        try { styleHeader(users.getRange(1, 6), "#0f766e"); } catch(e){}
      }
    } catch(e){}
    try {
      // Migrate: check G header is From Date
      var fromHeader = String(users.getRange(1, 7).getValue()).trim();
      if (fromHeader !== "From Date") {
        users.getRange(1, 7).setValue("From Date");
        try { styleHeader(users.getRange(1, 7), "#0f766e"); } catch(e){}
      }
    } catch(e){}
    try {
      // Migrate: check H header is To Date
      var toHeader = String(users.getRange(1, 8).getValue()).trim();
      if (toHeader !== "To Date") {
        users.getRange(1, 8).setValue("To Date");
        try { styleHeader(users.getRange(1, 8), "#0f766e"); } catch(e){}
      }
    } catch(e){}
    try {
      // Ensure the admin user exists and has the admin role
      var adminRowIndex = -1;
      var uLast2 = users.getLastRow();
      if (uLast2 > 1) {
        var phones = users.getRange(2, 2, uLast2 - 1, 1).getValues();
        for (var p = 0; p < phones.length; p++) {
          if (normalizePhone(phones[p][0]) === ADMIN_SEED.phone) {
            adminRowIndex = p + 2;
            break;
          }
        }
      }
      if (adminRowIndex === -1) {
        users.appendRow([ADMIN_SEED.name, ADMIN_SEED.phone, ROLE_ADMIN, ""]);
        try { users.getRange(users.getLastRow(), 2).setNumberFormat("@").setValue(ADMIN_SEED.phone); } catch(e){}
      } else {
        users.getRange(adminRowIndex, 3).setValue(ROLE_ADMIN);
      }
    } catch(e){}
  }

  // --- Thursday Calling tab: create or migrate ---
  var calling = doc.getSheetByName(SHEET_CALLING);
  if (!calling) {
    calling = doc.insertSheet(SHEET_CALLING);
    try { calling.getRange(1, 1, 1, FIXED_COLS)
      .setValues([["Name", "Phone Number", "W/S", "Sessions Attended", "No. of Calls", "Cultivated By", "Assigned To"]]); } catch(e){}
    try { styleHeader(calling.getRange(1, 1, 1, FIXED_COLS), "#1e1b4b"); } catch(e){}
    try { calling.setFrozenRows(1); } catch(e){}
    try { calling.getRange(2, 2, 1000, 1).setNumberFormat("@"); } catch(e){}
    try { calling.getRange(2, 1, SAMPLE_CONTACTS.length, FIXED_COLS).setValues(SAMPLE_CONTACTS); } catch(e){}
  } else {
    try {
      // Migrate: insert "No. of Calls" as column E if it isn't there yet
      var e1 = String(calling.getRange(1, COL_CALLS).getValue()).trim();
      if (e1 !== "No. of Calls") {
        calling.insertColumnBefore(COL_CALLS);
        try { calling.getRange(1, COL_CALLS, calling.getMaxRows(), 1).clearDataValidations(); } catch(e){}
        calling.getRange(1, COL_CALLS).setValue("No. of Calls");
        try { styleHeader(calling.getRange(1, COL_CALLS), "#1e1b4b"); } catch(e){}
      }
    } catch(e){}
    try {
      // Migrate: insert "Cultivated By" as column F (6) if it isn't there yet
      var f1 = String(calling.getRange(1, COL_CULTIVATED).getValue()).trim();
      if (f1 !== "Cultivated By") {
        calling.insertColumnBefore(COL_CULTIVATED);
        try { calling.getRange(1, COL_CULTIVATED, calling.getMaxRows(), 1).clearDataValidations(); } catch(e){}
        calling.getRange(1, COL_CULTIVATED).setValue("Cultivated By");
        try { styleHeader(calling.getRange(1, COL_CULTIVATED), "#1e1b4b"); } catch(e){}
      }
    } catch(e){}
    try {
      // Older sheets: make sure the Assigned To header exists (now column G)
      var g1 = String(calling.getRange(1, COL_ASSIGNED).getValue()).trim();
      if (g1 !== "Assigned To") {
        calling.getRange(1, COL_ASSIGNED).setValue("Assigned To");
        try { styleHeader(calling.getRange(1, COL_ASSIGNED), "#1e1b4b"); } catch(e){}
      }
    } catch(e){}
  }

  // --- Session History tab: create or seed ---
  try {
    var sessionHistory = doc.getSheetByName("Session History");
    if (!sessionHistory) {
      sessionHistory = doc.insertSheet("Session History");
      try { sessionHistory.getRange(1, 1, 1, 5)
        .setValues([["Time", "Session Name", "Contact Name", "Phone Number", "Attended"]]); } catch(e){}
      try { styleHeader(sessionHistory.getRange(1, 1, 1, 5), "#1e1b4b"); } catch(e){}
      try { sessionHistory.setFrozenRows(1); } catch(e){}
      try { sessionHistory.getRange(2, 4, 1000, 1).setNumberFormat("@"); } catch(e){}
      
      // Seed sample sessions (match starting record counts: Aarav has 3, Ravi has 2)
      var SAMPLE_SESSIONS = [
        ["1 Jul 2026, 7:00 PM", "Gita Study Course — Session 1", "Aarav Sharma", "9812345601", "Yes"],
        ["2 Jul 2026, 7:00 PM", "Gita Study Course — Session 2", "Aarav Sharma", "9812345601", "Yes"],
        ["8 Jul 2026, 7:00 PM", "Gita Study Course — Session 3", "Aarav Sharma", "9812345601", "Yes"],
        ["1 Jul 2026, 7:00 PM", "Gita Study Course — Session 1", "Ravi Teja", "9812345602", "Yes"],
        ["8 Jul 2026, 7:00 PM", "Gita Study Course — Session 3", "Ravi Teja", "9812345602", "Yes"]
      ];
      try { sessionHistory.getRange(2, 1, SAMPLE_SESSIONS.length, 5).setValues(SAMPLE_SESSIONS); } catch(e){}
    }
  } catch(e){}

  // --- Festival Promotions tab: create or migrate ---
  try {
    var festival = doc.getSheetByName(SHEET_FESTIVAL);
    if (!festival) {
      festival = doc.insertSheet(SHEET_FESTIVAL);
      try { festival.getRange(1, 1, 1, 8)
        .setValues([["Name", "Phone Number", "W/S", "Sessions Attended", "No. of Calls", "Cultivated By", "Assigned To", "Event"]]); } catch(e){}
      try { styleHeader(festival.getRange(1, 1, 1, 8), "#4338ca"); } catch(e){}
      try { festival.setFrozenRows(1); } catch(e){}
      try { festival.getRange(2, 2, 1000, 1).setNumberFormat("@"); } catch(e){}
      
      var SAMPLE_FESTIVAL = [
        ["Krishna Das", "9812345610", "W", 0, 0, "", "", "Ratha Yatra"],
        ["Radha Dasi", "9812345611", "S", 0, 0, "", "", "Ratha Yatra"],
        ["Balaram Dev", "9812345612", "NA", 0, 0, "", "", "Ratha Yatra"]
      ];
      try { festival.getRange(2, 1, SAMPLE_FESTIVAL.length, 8).setValues(SAMPLE_FESTIVAL); } catch(e){}
    } else {
      try {
        var e1 = String(festival.getRange(1, COL_CALLS).getValue()).trim();
        if (e1 !== "No. of Calls") {
          festival.insertColumnBefore(COL_CALLS);
          try { festival.getRange(1, COL_CALLS, festival.getMaxRows(), 1).clearDataValidations(); } catch(e){}
          festival.getRange(1, COL_CALLS).setValue("No. of Calls");
          try { styleHeader(festival.getRange(1, COL_CALLS), "#4338ca"); } catch(e){}
        }
      } catch(e){}
      try {
        var f1 = String(festival.getRange(1, COL_CULTIVATED).getValue()).trim();
        if (f1 !== "Cultivated By") {
          festival.insertColumnBefore(COL_CULTIVATED);
          try { festival.getRange(1, COL_CULTIVATED, festival.getMaxRows(), 1).clearDataValidations(); } catch(e){}
          festival.getRange(1, COL_CULTIVATED).setValue("Cultivated By");
          try { styleHeader(festival.getRange(1, COL_CULTIVATED), "#4338ca"); } catch(e){}
        }
      } catch(e){}
      try {
        var g1 = String(festival.getRange(1, COL_ASSIGNED).getValue()).trim();
        if (g1 !== "Assigned To") {
          festival.getRange(1, COL_ASSIGNED).setValue("Assigned To");
          try { styleHeader(festival.getRange(1, COL_ASSIGNED), "#4338ca"); } catch(e){}
        }
      } catch(e){}
      try {
        // Dynamic Event Column Migration:
        var currentHeaders = festival.getRange(1, 1, 1, Math.max(8, festival.getLastColumn())).getValues()[0];
        if (currentHeaders.indexOf("Event") === -1) {
          festival.insertColumnAfter(7); // Insert Event at Column 8
          festival.getRange(1, 8).setValue("Event");
          try { styleHeader(festival.getRange(1, 8), "#4338ca"); } catch(e){}
          SpreadsheetApp.flush();
        }
      } catch(e){}
    }
  } catch(e){}

  // First date column for festival promotions
  try {
    var festival = festival || doc.getSheetByName(SHEET_FESTIVAL);
    if (festival) {
      var activeFestival = getActiveDateColumn(festival);
      if (activeFestival.activeCol === -1) {
        var todayName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d");
        festival.getRange(1, 9).setValue(todayName);
        try { styleHeader(festival.getRange(1, 9), "#4338ca"); } catch(e){}
      }
    }
  } catch(e){}

  // First date column: create one named with today's date if none exists yet.
  // (Afterwards, start each new round by adding a column header manually.)
  try {
    var active = getActiveDateColumn(calling);
    if (active.activeCol === -1) {
      var todayName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d");
      calling.getRange(1, FIXED_COLS + 1).setValue(todayName);
      try { styleHeader(calling.getRange(1, FIXED_COLS + 1), "#0f766e"); } catch(e){}
    }
  } catch(e){}

  // --- Settings tab: create or seed ---
  try {
    var settingsSheet = doc.getSheetByName("Settings");
    if (!settingsSheet) {
      settingsSheet = doc.insertSheet("Settings");
      try { settingsSheet.getRange(1, 1, 1, 2).setValues([["Key", "Value"]]); } catch(e){}
      try { styleHeader(settingsSheet.getRange(1, 1, 1, 2), "#0f766e"); } catch(e){}
      try { settingsSheet.setFrozenRows(1); } catch(e){}
      
      var DEFAULT_SETTINGS = [
        ["calling_message", "Hare Krishna {name}, please attend Thursday Session."],
        ["festival_message", "Hare Krishna {name}, you are invited to our upcoming festival."],
        ["calling_poster", ""],
        ["festival_poster", ""]
      ];
      try { settingsSheet.getRange(2, 1, DEFAULT_SETTINGS.length, 2).setValues(DEFAULT_SETTINGS); } catch(e){}
    }
  } catch(e){}

  try { SpreadsheetApp.flush(); } catch(e){}

  } catch (setupError) {
    // Skip cosmetics this round; data reads/writes below still work.
  }

  // ALWAYS cache setup result — even if some operations failed on typed columns.
  // This prevents the setup from retrying and crashing on every single request.
  cache.put(CACHE_KEY_SETUP, "1", 21600); // re-check headers at most every 6h

  return doc;
}

/**
 * Keeps the sheet tidy and self-maintaining:
 *   - dropdowns only on rows that hold data and date columns that exist
 *   - "No. of Calls" formulas on every data row
 * Re-applies automatically when rows or date columns are added/removed
 * (the row/column signature changes, so the next data read refreshes them).
 */
function applyValidations(doc, sheetName) {
  try {

  sheetName = sheetName || SHEET_CALLING;
  var cache = CacheService.getScriptCache();
  var calling = doc.getSheetByName(sheetName);
  var users = doc.getSheetByName(SHEET_USERS);
  if (!calling || !users) return;

  var lastRow = calling.getLastRow();
  var active = getActiveDateColumn(calling);
  var signature = sheetName + ":" + lastRow + ":" + active.activeCol;
  if (cache.get(CACHE_KEY_VALIDATION) === signature) return;

  var maxRows = calling.getMaxRows();
  var maxCols = calling.getMaxColumns();

  var fixed = (sheetName === SHEET_FESTIVAL) ? 8 : 7;

  // Color headers dynamically
  var headerColor = (sheetName === SHEET_FESTIVAL) ? "#4338ca" : "#1e1b4b";
  try { styleHeader(calling.getRange(1, 1, 1, fixed), headerColor); } catch(e){}

  // Clear every dropdown first, then re-apply only where data actually exists
  try { calling.getRange(2, COL_WS, maxRows - 1, 1).clearDataValidations(); } catch(e){}
  try { calling.getRange(2, COL_CULTIVATED, maxRows - 1, 1).clearDataValidations(); } catch(e){}
  try { calling.getRange(2, COL_ASSIGNED, maxRows - 1, maxCols - COL_ASSIGNED + 1).clearDataValidations(); } catch(e){}

  var dataRows = lastRow - 1;
  if (dataRows > 0) {
    try {
      var wsRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(WS_OPTIONS, true).setAllowInvalid(true).build();
      calling.getRange(2, COL_WS, dataRows, 1).setDataValidation(wsRule);
    } catch(e){}

    try {
      var assignRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(users.getRange("A2:A1000"), true).setAllowInvalid(true).build();
      calling.getRange(2, COL_CULTIVATED, dataRows, 1).setDataValidation(assignRule);
      calling.getRange(2, COL_ASSIGNED, dataRows, 1).setDataValidation(assignRule);
    } catch(e){}

    if (active.activeCol !== -1) {
      try {
        var statusRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(STATUS_OPTIONS, true).setAllowInvalid(true).build();
        calling.getRange(2, fixed + 1, dataRows, active.activeCol - fixed).setDataValidation(statusRule);
      } catch(e){}
      // Date column headers keep the same styling as the rest of the header row
      try { styleHeader(calling.getRange(1, fixed + 1, 1, active.activeCol - fixed), "#0f766e"); } catch(e){}
    }

    // No. of Sessions: live formula counting Yes values in Session History matching Phone
    try {
      calling.getRange(2, 4, maxRows - 1, 1).clearDataValidations();
      var sessionFormulas = [];
      for (var r = 2; r <= lastRow; r++) {
        sessionFormulas.push([
          '=COUNTIFS(\'Session History\'!D:D, $B' + r + ', \'Session History\'!E:E, "Yes")'
        ]);
      }
      calling.getRange(2, 4, dataRows, 1).setFormulas(sessionFormulas);
    } catch (sessionFormulaError) {
      // fallback
    }

    // No. of Calls: live formula — counts date cells filled with anything
    // other than "Not Done" across ALL date columns of that row.
    try {
      calling.getRange(2, COL_CALLS, maxRows - 1, 1).clearDataValidations();
      var formulas = [];
      for (var r = 2; r <= lastRow; r++) {
        formulas.push([
          '=IF($A' + r + '="","",COUNTA($H' + r + ':' + r + ')-COUNTIF($H' + r + ':' + r + ',"Not Done"))'
        ]);
      }
      calling.getRange(2, COL_CALLS, dataRows, 1).setFormulas(formulas);
    } catch (formulaError) {
      // Leave the column as-is; getCallingData() falls back to a computed count.
    }
  }

  cache.put(CACHE_KEY_VALIDATION, signature, 21600);

  } catch (validationError) {
    // Dropdowns/formulas are conveniences — never let them break the API.
    // getCallingData() computes call counts itself, so data stays correct.
  }
}

/**
 * Append-only log: every status submission adds one row here with the exact
 * time, so admin can see that a contact was called at 11 AM and again at 5 PM.
 * The main tab keeps only the latest status per round (count stays 1).
 */
function logCallHistory(doc, round, contactName, phone, status, byName) {
  try {
    var sheet = doc.getSheetByName("Call History");
    if (!sheet) {
      sheet = doc.insertSheet("Call History");
      sheet.getRange(1, 1, 1, 6)
        .setValues([["Time", "Round", "Contact Name", "Phone Number", "Status", "Submitted By"]]);
      styleHeader(sheet.getRange(1, 1, 1, 6), "#1e1b4b");
      sheet.setFrozenRows(1);
      sheet.getRange(2, 4, 1000, 1).setNumberFormat("@");
    }
    var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "d MMM yyyy, h:mm a");
    sheet.appendRow([ts, "'" + round, contactName, phone, status, byName]);
  } catch (historyError) {
    // History is a convenience — never let it break a submission.
  }
}

/**
 * All users: [{name, phone, role}] — cached briefly so logins stay fast
 * but newly added users can log in within ~30s.
 */
function getUsers(skipCache) {
  var cache = CacheService.getScriptCache();
  if (!skipCache) {
    var cached = cache.get(CACHE_KEY_USERS);
    if (cached) {
      try { return JSON.parse(cached); } catch (ignored) {}
    }
  }

  var doc = ensureSetup();
  var sheet = doc.getSheetByName(SHEET_USERS);
  var lastRow = sheet.getLastRow();
  var result = [];

  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (var i = 0; i < values.length; i++) {
      var name = String(values[i][0]).trim();
      var phone = normalizePhone(values[i][1]);
      var rawRole = String(values[i][2]).trim().toLowerCase();
      if (rawRole === "festivals") continue;
      if (name === "" || phone === "") continue;
      var role = rawRole === ROLE_ADMIN ? ROLE_ADMIN : ROLE_USER;
      var limit = Number(values[i][3]);
      var limitVal = (isNaN(limit) || values[i][3] === "") ? 999 : limit;
      var autoAssign = String(values[i][4] || "").trim().toLowerCase() === "yes";
      var festival = String(values[i][5] || "").trim();
      result.push({ 
        name: name, 
        phone: phone, 
        role: role, 
        limit: limitVal, 
        autoAssign: autoAssign,
        festival: festival
      });
    }
  }

  cache.put(CACHE_KEY_USERS, JSON.stringify(result), 30);
  return result;
}

function getFestivals() {
  var doc = ensureSetup();
  var sheet = doc.getSheetByName(SHEET_USERS);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var festivals = [];
  var seen = {};
  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues(); // Read columns A to H
    for (var i = 0; i < values.length; i++) {
      var role = String(values[i][2]).trim().toLowerCase();
      if (role === "festivals") {
        var name = String(values[i][0]).trim(); // Column A (Name)
        if (name !== "" && !seen.hasOwnProperty(name)) {
          festivals.push(name);
          seen[name] = true;
        }
      }
    }
  }
  return festivals;
}

function findUser(phone) {
  var target = normalizePhone(phone);
  if (target.length < 10) return null;
  
  // Try reading from cache first
  var users = getUsers(false);
  for (var i = 0; i < users.length; i++) {
    if (users[i].phone === target) return users[i];
  }
  
  // Cache-busting fallback: if not found, pull fresh from sheet to handle instant spreadsheet updates
  users = getUsers(true);
  for (var i = 0; i < users.length; i++) {
    if (users[i].phone === target) return users[i];
  }
  
  return null;
}

/**
 * Filter the Call History sheet to find all logs matching the target contact phone.
 */
function getCallHistoryForContact(phone) {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName("Call History");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var history = [];
  var target = normalizePhone(phone);
  for (var i = 0; i < values.length; i++) {
    if (normalizePhone(values[i][3]) === target) {
      history.push({
        time: values[i][0],
        round: values[i][1],
        name: values[i][2],
        status: values[i][4],
        by: values[i][5]
      });
    }
  }
  return history;
}

/**
 * Filter the Session History sheet to find all logs matching the target contact phone.
 */
function getSessionHistoryForContact(phone) {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName("Session History");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var history = [];
  var target = normalizePhone(phone);
  for (var i = 0; i < values.length; i++) {
    if (normalizePhone(values[i][3]) === target) {
      history.push({
        time: values[i][0],
        session: values[i][1],
        name: values[i][2],
        attended: values[i][4]
      });
    }
  }
  return history;
}

/**
 * The calling sheet as data:
 *   { activeDate, contacts: [{name, phone, ws, sessions, calls, cultivatedBy, assignedTo, status}] }
 * status comes from the ACTIVE (last) date column; empty cells read as "Not Done".
 */
function getCallingData(skipCache, sheetName) {
  sheetName = sheetName || SHEET_CALLING;
  
  if (sheetName === SHEET_FESTIVAL) {
    try {
      processFormResponses();
    } catch(formErr) {
      // safe fallback
    }
    try {
      runAutoAssignment();
    } catch(autoAssignErr) {
      // safe fallback
    }
  }

  var cache = CacheService.getScriptCache();
  var cacheKey = CACHE_KEY_CALLING + "_" + sheetName.replace(/\s+/g, "_");
  if (!skipCache) {
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (ignored) {}
    }
  }

  var doc = ensureSetup();
  applyValidations(doc, sheetName); // keep dropdowns & call-count formulas matched to data size
  var sheet = doc.getSheetByName(sheetName);
  if (!sheet) {
    return { activeDate: "", contacts: [] };
  }
  var active = getActiveDateColumn(sheet);
  var fixed = (sheetName === SHEET_FESTIVAL) ? 8 : 7;
  var lastRow = sheet.getLastRow();
  var width = Math.max(sheet.getLastColumn(), fixed);
  var contacts = [];

  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    for (var i = 0; i < values.length; i++) {
      var name = String(values[i][0]).trim();
      if (name === "") continue;
      var status = "";
      if (active.activeCol !== -1) {
        status = String(values[i][active.activeCol - 1]).trim();
      }

      var callsCount = 0;
      for (var d = fixed; d < width; d++) {
        var cell = String(values[i][d]).trim();
        if (cell !== "" && cell.toLowerCase() !== STATUS_DEFAULT.toLowerCase()) callsCount++;
      }

      var contactObj = {
        name: name,
        phone: normalizePhone(values[i][1]),
        ws: String(values[i][2]).trim(),
        sessions: Number(values[i][3]) || 0,
        calls: callsCount,
        cultivatedBy: String(values[i][COL_CULTIVATED - 1]).trim(),
        assignedTo: String(values[i][COL_ASSIGNED - 1]).trim(),
        status: status === "" ? STATUS_DEFAULT : status
      };
      if (sheetName === SHEET_FESTIVAL) {
        contactObj.event = String(values[i][7] || "").trim();
      }
      contacts.push(contactObj);
    }
  }

  var data = { activeDate: active.activeDate, contacts: contacts };
  cache.put(cacheKey, JSON.stringify(data), CACHE_SECONDS);
  return data;
}

/**
 * Admin sees everything; a user sees only contacts assigned to them.
 */
function contactsForUser(user, contacts, type) {
  if (user.role === ROLE_ADMIN) return contacts;
  var target = user.name.toLowerCase();
  return contacts.filter(function (c) {
    if (type === "cultivation") {
      return c.cultivatedBy.toLowerCase() === target;
    }
    return c.assignedTo.toLowerCase() === target;
  });
}

/**
 * The full page payload for a logged-in user (login, refresh, and post-save).
 */
function getActiveWeekCallCounts(doc, activeDate) {
  var counts = {};
  if (!activeDate) return counts;
  var sheet = doc.getSheetByName("Call History");
  if (!sheet) return counts;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return counts;
  var values = sheet.getRange(2, 2, lastRow - 1, 3).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var round = String(values[i][0]).trim();
    var phone = normalizePhone(values[i][2]);
    if (round === activeDate) {
      counts[phone] = (counts[phone] || 0) + 1;
    }
  }
  return counts;
}

function buildDataPayload(user, sheetName, campaignType, skipCache) {
  sheetName = sheetName || SHEET_CALLING;
  campaignType = campaignType || "calling";
  var data = getCallingData(!!skipCache, sheetName);
  var userContacts = contactsForUser(user, data.contacts, campaignType);

  // Users see call counts for the active week only, admins see total call history count.
  if (user.role === ROLE_USER) {
    var doc = ensureSetup();
    var activeWeekCounts = getActiveWeekCallCounts(doc, data.activeDate);
    userContacts = userContacts.map(function (c) {
      var copy = Object.assign({}, c);
      copy.calls = activeWeekCounts[c.phone] || 0;
      return copy;
    });
  }

  var payload = {
    status: "success",
    user: user,
    page: "thursday-calling",
    activeDate: data.activeDate,
    contacts: userContacts,
    wsOptions: WS_OPTIONS,
    statusOptions: STATUS_OPTIONS,
    festivals: getFestivals(),
    settings: getSettings()
  };
  if (user.role === ROLE_ADMIN) {
    var allUsers = getUsers(false);
    payload.userNames = allUsers.map(function (u) { return u.name; });
    payload.userLimits = allUsers.reduce(function (acc, u) { acc[u.name] = u.limit; return acc; }, {});
    payload.userPhones = allUsers.map(function (u) { return normalizePhone(u.phone); });

    // Auto-assign state for Festival Promotions admin panel
    var aaUsers = [];
    var aaUserFestivals = {};
    allUsers.forEach(function(u) {
      if (u.autoAssign) {
        aaUsers.push(u.name);
      }
      aaUserFestivals[u.name] = u.festival || ""; // Include for all users so dropdown works for any checkbox toggle
    });
    payload.autoAssignActive = aaUsers.length > 0;
    payload.autoAssignUsers = aaUsers;
    payload.autoAssignUserFestivals = aaUserFestivals;
  }
  return payload;
}

// ---------------------------------------------------------------
// GET
//   ?action=login&phone=98xxxxxxxx -> validate + user's page data
//   ?action=data&phone=98xxxxxxxx  -> refresh (filtered per role)
//   ?action=history&phone=98xxxxxxxx -> fetch call history of a contact
// ---------------------------------------------------------------

function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = params.action || "data";

    if (action === "history") {
      var contactPhone = params.phone || "";
      var history = getCallHistoryForContact(contactPhone);
      return jsonResponse({ status: "success", history: history });
    }

    if (action === "sessions") {
      var contactPhone = params.phone || "";
      var sessions = getSessionHistoryForContact(contactPhone);
      return jsonResponse({ status: "success", sessions: sessions });
    }

    if (action === "searchContact") {
      var searchPhone = normalizePhone(params.phone || "");
      var doc = SpreadsheetApp.getActiveSpreadsheet();
      var foundContact = null;
      var sheets = [SHEET_CALLING, SHEET_FESTIVAL];
      
      for (var s = 0; s < sheets.length; s++) {
        var sheet = doc.getSheetByName(sheets[s]);
        if (!sheet) continue;
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) continue;
        var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (var i = 0; i < values.length; i++) {
          if (normalizePhone(values[i][1]) === searchPhone) {
            foundContact = {
              name: String(values[i][0]).trim(),
              phone: searchPhone
            };
            break;
          }
        }
        if (foundContact) break;
      }
      
      if (foundContact) {
        var sessionsCount = 0;
        var sessionHistory = doc.getSheetByName("Session History");
        if (sessionHistory) {
          var sLast = sessionHistory.getLastRow();
          if (sLast > 1) {
            var sValues = sessionHistory.getRange(2, 4, sLast - 1, 2).getValues();
            for (var j = 0; j < sValues.length; j++) {
              if (normalizePhone(sValues[j][0]) === searchPhone && String(sValues[j][1]).trim().toLowerCase() === "yes") {
                sessionsCount++;
              }
            }
          }
        }
        foundContact.sessions = sessionsCount;
        return jsonResponse({ status: "success", found: true, contact: foundContact });
      } else {
        return jsonResponse({ status: "success", found: false });
      }
    }

    if (action === "login" || action === "data") {
      var user = findUser(params.phone || "");
      if (!user) {
        return jsonResponse({
          status: "error",
          code: "USER_NOT_FOUND",
          message: "This phone number is not registered. Please contact your coordinator."
        });
      }
      var sheetName = params.sheet || SHEET_CALLING;
      var campaignType = params.campaignType || "calling";
      var skipCache = params.skipCache === "true";
      return jsonResponse(buildDataPayload(user, sheetName, campaignType, skipCache));
    }

    return jsonResponse({ status: "error", code: "UNKNOWN_ACTION", message: "Unknown action: " + action });

  } catch (error) {
    return jsonResponse({ status: "error", code: "SERVER_ERROR", message: error.toString() });
  }
}

// ---------------------------------------------------------------
// POST (JSON body sent as text/plain from the frontend)
//   {
//     action: "updateContact",
//     requesterPhone: "99…",       // who is making the change
//     phone: "98…",                // which contact row to update
//     updates: { ws, status, cultivatedBy, assignedTo }
//   }
//   status is written into the ACTIVE date column. It may be a list value or
//   free text (the portal's "Others" option). assignedTo/cultivatedBy: admin only.
//   Sessions and Phone Number are NOT editable through the API at all.
// ---------------------------------------------------------------

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: "error", code: "NO_PAYLOAD", message: "No payload received" });
    }

    var body;
    try {
      var rawContent = e.postData.contents || "";
      if (rawContent.trim().charAt(0) !== "{") {
        try {
          rawContent = decodeURIComponent(rawContent);
        } catch(e){}
      }
      body = JSON.parse(rawContent);
    } catch (parseError) {
      return jsonResponse({ status: "error", code: "BAD_JSON", message: "Invalid JSON payload: " + parseError.toString() });
    }

    if (body.action === "saveSetting") {
      var requester = findUser(body.requesterPhone || "");
      if (!requester || requester.role !== ROLE_ADMIN) {
        return jsonResponse({ status: "error", code: "FORBIDDEN", message: "Only the admin can configure messages." });
      }
      saveSetting(body.key, body.value);
      return jsonResponse({ status: "success", settings: getSettings() });
    }

    if (body.action === "saveSettingImage") {
      var requester = findUser(body.requesterPhone || "");
      if (!requester || requester.role !== ROLE_ADMIN) {
        return jsonResponse({ status: "error", code: "FORBIDDEN", message: "Only the admin can configure messages." });
      }
      try {
        var directUrl = saveSettingImage(body.key, body.base64Data, body.mimeType, body.fileName);
        return jsonResponse({ status: "success", url: directUrl, settings: getSettings() });
      } catch (uploadErr) {
        return jsonResponse({ status: "error", code: "UPLOAD_FAILED", message: "Failed to upload image to Google Drive: " + uploadErr.toString() });
      }
    }

    if (body.action === "updateContactsBulk") {
      var requester = findUser(body.requesterPhone || "");
      if (!requester) {
        return jsonResponse({ status: "error", code: "AUTH", message: "Your session is invalid. Please log in again." });
      }
      var isAdmin = requester.role === ROLE_ADMIN;
      if (!isAdmin) {
        return jsonResponse({ status: "error", code: "FORBIDDEN", message: "Only the admin can perform bulk updates." });
      }

      var locked = lock.tryLock(30000);
      if (!locked) {
        return jsonResponse({ status: "error", code: "BUSY", message: "Server is busy, please try again." });
      }

      var doc = ensureSetup();
      var sheetName = body.sheet || SHEET_CALLING;
      var sheet = doc.getSheetByName(sheetName);
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return jsonResponse(buildDataPayload(requester, sheetName));
      }

      var bulkUpdates = body.updates || []; 
      var fixed = (sheetName === SHEET_FESTIVAL) ? 8 : 7;
      var rows = sheet.getRange(2, 1, lastRow - 1, fixed).getValues();
      var active = getActiveDateColumn(sheet);
      var allUsers = getUsers(true);

      for (var uIndex = 0; uIndex < bulkUpdates.length; uIndex++) {
        var item = bulkUpdates[uIndex];
        var targetPhone = normalizePhone(item.phone || "");
        var updates = item.updates || {};

        // Find rowIndex
        var rowIndex = -1;
        var rowData = null;
        for (var i = 0; i < rows.length; i++) {
          if (normalizePhone(rows[i][1]) === targetPhone) {
            rowIndex = i + 2;
            rowData = rows[i];
            break;
          }
        }
        if (rowIndex === -1) continue;

        // Apply updates
        if (updates.hasOwnProperty("name")) {
          var newName = String(updates.name).trim();
          if (newName !== "") {
            sheet.getRange(rowIndex, 1).setValue(newName);
          }
        }
        var contactName = updates.hasOwnProperty("name") ? String(updates.name).trim() : String(rowData[0]).trim();

        if (updates.hasOwnProperty("ws")) {
          var ws = String(updates.ws).trim().toUpperCase();
          if (WS_OPTIONS.indexOf(ws) !== -1) {
            sheet.getRange(rowIndex, COL_WS).setValue(ws);
          }
        }

        if (updates.hasOwnProperty("status")) {
          var callStatus = String(updates.status).trim();
          if (callStatus !== "" && callStatus.toLowerCase() !== "others" && active.activeCol !== -1) {
            sheet.getRange(rowIndex, active.activeCol).setValue(callStatus);
            if (callStatus !== STATUS_DEFAULT) {
              logCallHistory(doc, active.activeDate, contactName, targetPhone, callStatus, requester.name);
            }
          }
        }
        if (sheetName === SHEET_FESTIVAL && updates.hasOwnProperty("event")) {
          sheet.getRange(rowIndex, 8).setValue(String(updates.event).trim());
        }

        if (updates.hasOwnProperty("cultivatedBy")) {
          var cultivator = String(updates.cultivatedBy).trim();
          if (cultivator !== "") {
            var match = null;
            for (var u = 0; u < allUsers.length; u++) {
              if (allUsers[u].name.toLowerCase() === cultivator.toLowerCase()) { match = allUsers[u].name; break; }
            }
            if (match) cultivator = match;
          }
          sheet.getRange(rowIndex, COL_CULTIVATED).setValue(cultivator);
        }

        if (updates.hasOwnProperty("assignedTo")) {
          var assignee = String(updates.assignedTo).trim();
          if (assignee !== "") {
            var match = null;
            for (var u = 0; u < allUsers.length; u++) {
              if (allUsers[u].name.toLowerCase() === assignee.toLowerCase()) { match = allUsers[u].name; break; }
            }
            if (match) assignee = match;
          }
          sheet.getRange(rowIndex, COL_ASSIGNED).setValue(assignee);
        }
      }

      SpreadsheetApp.flush();
      var cacheKey = CACHE_KEY_CALLING + "_" + sheetName.replace(/\s+/g, "_");
      CacheService.getScriptCache().remove(cacheKey);
      return jsonResponse(buildDataPayload(requester, sheetName));
    }

    if (body.action === "markAttendance") {
      var requester = findUser(body.requesterPhone || "");
      if (!requester) {
        return jsonResponse({ status: "error", code: "AUTH", message: "Your session is invalid. Please log in again." });
      }
      
      var phone = normalizePhone(body.phone || "");
      var name = String(body.name || "").trim();
      var sessionName = String(body.sessionName || "General Session").trim();
      
      if (phone === "" || name === "") {
        return jsonResponse({ status: "error", code: "BAD_VALUE", message: "Name and Phone Number are required." });
      }
      
      var locked = lock.tryLock(30000);
      if (!locked) {
        return jsonResponse({ status: "error", code: "BUSY", message: "Server is busy, please try again." });
      }
      
      try {
        var doc = ensureSetup();
        var sessionHistory = doc.getSheetByName("Session History");
        if (!sessionHistory) {
          sessionHistory = doc.insertSheet("Session History");
          sessionHistory.getRange(1, 1, 1, 5).setValues([["Time", "Session Name", "Contact Name", "Phone Number", "Attended"]]);
          styleHeader(sessionHistory.getRange(1, 1, 1, 5), "#1e1b4b");
          sessionHistory.setFrozenRows(1);
        try {
          sessionHistory.getRange(2, 4, 1000, 1).setNumberFormat("@");
        } catch(e) {}
      }
      
      var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "d MMM yyyy, h:mm a");
      sessionHistory.appendRow([ts, sessionName, name, phone, "Yes"]);
      try {
        sessionHistory.getRange(sessionHistory.getLastRow(), 4).setNumberFormat("@");
      } catch(e) {}
        
        var isRegistered = false;
        var sheets = [SHEET_CALLING, SHEET_FESTIVAL];
        for (var s = 0; s < sheets.length; s++) {
          var sheet = doc.getSheetByName(sheets[s]);
          if (!sheet) continue;
          var lastRow = sheet.getLastRow();
          if (lastRow > 1) {
            var phones = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
            for (var p = 0; p < phones.length; p++) {
              if (normalizePhone(phones[p][0]) === phone) { isRegistered = true; break; }
            }
          }
          if (isRegistered) break;
        }
        
        if (!isRegistered) {
          var targetSheet = doc.getSheetByName(SHEET_CALLING);
          if (targetSheet) {
            var nextRow = targetSheet.getLastRow() + 1;
            var rowValues = [
              name,
              phone,
              "NA",
              '=COUNTIFS(\'Session History\'!D:D, $B' + nextRow + ', \'Session History\'!E:E, "Yes")',
              0,
              "",
              ""
            ];
            targetSheet.appendRow(rowValues);
            try {
              targetSheet.getRange(nextRow, 2).setNumberFormat("@");
            } catch(e) {}
            applyValidations(doc, SHEET_CALLING);
          }
        }
        
        SpreadsheetApp.flush();
        CacheService.getScriptCache().remove(CACHE_KEY_CALLING + "_" + SHEET_CALLING.replace(/\s+/g, "_"));
        CacheService.getScriptCache().remove(CACHE_KEY_CALLING + "_" + SHEET_FESTIVAL.replace(/\s+/g, "_"));
        
        return jsonResponse({ status: "success", message: "Attendance marked successfully" });
      } finally {
        try { lock.releaseLock(); } catch (ignored) {}
      }
    }

    if (body.action === "updateAutoAssignUser") {
      var requester = findUser(body.requesterPhone || "");
      if (!requester || requester.role !== ROLE_ADMIN) {
        return jsonResponse({ status: "error", code: "FORBIDDEN", message: "Only admin can configure auto assignment." });
      }

      var targetUser = body.username || "";
      var enabled = body.enabled === true;

      var locked = lock.tryLock(30000);
      if (!locked) {
        return jsonResponse({ status: "error", code: "BUSY", message: "Server is busy, please try again." });
      }

      try {
        var doc = ensureSetup();
        var sheet = doc.getSheetByName(SHEET_USERS);
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
          for (var i = 0; i < names.length; i++) {
            if (String(names[i][0]).trim() === targetUser) {
              sheet.getRange(i + 2, 5).setValue(enabled ? "Yes" : "");
              break;
            }
          }
        }
        SpreadsheetApp.flush();
        CacheService.getScriptCache().remove(CACHE_KEY_USERS);

        // Run auto assignment immediately in case a user was newly enabled
        if (enabled) {
          try {
            runAutoAssignment();
          } catch(e) {}
        }

        return jsonResponse(buildDataPayload(requester, "Festival Promotions"));
      } finally {
        lock.releaseLock();
      }
    }

    if (body.action === "updateUserFestival") {
      var requester = findUser(body.requesterPhone || "");
      if (!requester || requester.role !== ROLE_ADMIN) {
        return jsonResponse({ status: "error", code: "FORBIDDEN", message: "Only admin can configure auto assignment." });
      }

      var targetUser = body.username || "";
      var festivalName = String(body.festival || "").trim();

      var locked = lock.tryLock(30000);
      if (!locked) {
        return jsonResponse({ status: "error", code: "BUSY", message: "Server is busy, please try again." });
      }

      try {
        var doc = ensureSetup();
        var sheet = doc.getSheetByName(SHEET_USERS);
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
          for (var i = 0; i < names.length; i++) {
            if (String(names[i][0]).trim() === targetUser) {
              sheet.getRange(i + 2, 6).setValue(festivalName); // Column F is 6
              break;
            }
          }
        }
        SpreadsheetApp.flush();
        CacheService.getScriptCache().remove(CACHE_KEY_USERS);

        // Run auto assignment immediately in case assignments can now be balanced
        try {
          runAutoAssignment();
        } catch(e) {}

        return jsonResponse(buildDataPayload(requester, "Festival Promotions"));
      } finally {
        lock.releaseLock();
      }
    }

    if (body.action === "importContacts") {
      var requester = findUser(body.requesterPhone || "");
      if (!requester) {
        return jsonResponse({ status: "error", code: "AUTH", message: "Your session is invalid. Please log in again." });
      }
      if (requester.role !== ROLE_ADMIN) {
        return jsonResponse({ status: "error", code: "FORBIDDEN", message: "Only the admin can import contacts." });
      }

      var sheetName = body.sheet || SHEET_CALLING;
      var rowsToImport = body.contacts || [];
      if (!Array.isArray(rowsToImport) || rowsToImport.length === 0) {
        return jsonResponse({ status: "error", code: "BAD_VALUE", message: "No contacts to import." });
      }

      var locked = lock.tryLock(60000);
      if (!locked) {
        return jsonResponse({ status: "error", code: "BUSY", message: "Server is busy, please try again." });
      }

      try {
        var doc = ensureSetup();
        var targetSheet = doc.getSheetByName(sheetName);
        if (!targetSheet) {
          return jsonResponse({ status: "error", code: "SHEET_NOT_FOUND", message: "Sheet not found: " + sheetName });
        }

        // Fetch existing phone numbers in this sheet
        var existingPhones = {};
        var lastRow = targetSheet.getLastRow();
        if (lastRow > 1) {
          var phones = targetSheet.getRange(2, 2, lastRow - 1, 1).getValues();
          for (var p = 0; p < phones.length; p++) {
            existingPhones[normalizePhone(phones[p][0])] = true;
          }
        }

        var importedCount = 0;
        var skippedCount = 0;
        var defaultEvent = body.event || "";

        for (var i = 0; i < rowsToImport.length; i++) {
          var row = rowsToImport[i];
          var name = String(row.name || "").trim();
          var phone = normalizePhone(row.phone || "");
          if (name === "" || phone === "" || phone.length !== 10) {
            skippedCount++;
            continue;
          }

          if (existingPhones[phone]) {
            skippedCount++;
            continue; // skip duplicate
          }

          var nextRow = targetSheet.getLastRow() + 1;
          var rowValues = [
            name,
            phone,
            "NA",
            '=COUNTIFS(\'Session History\'!D:D, $B' + nextRow + ', \'Session History\'!E:E, "Yes")',
            0,
            "",
            ""
          ];
          if (sheetName === SHEET_FESTIVAL) {
            rowValues.push(row.event || defaultEvent);
          }
          targetSheet.appendRow(rowValues);
          try {
            targetSheet.getRange(nextRow, 2).setNumberFormat("@");
          } catch(e) {}
          
          existingPhones[phone] = true;
          importedCount++;
        }

        applyValidations(doc, sheetName);
        SpreadsheetApp.flush();

        // Run auto assignment immediately if this is Festival Promotions
        if (sheetName === SHEET_FESTIVAL) {
          try {
            runAutoAssignment();
          } catch(e) {}
        }

        var cacheKey = CACHE_KEY_CALLING + "_" + sheetName.replace(/\s+/g, "_");
        CacheService.getScriptCache().remove(cacheKey);

        var payload = buildDataPayload(requester, sheetName);
        payload.importedCount = importedCount;
        payload.skippedCount = skippedCount;
        return jsonResponse(payload);

      } finally {
        try { lock.releaseLock(); } catch (ignored) {}
      }
    }

    if (body.action === "addContact") {
      var requester = findUser(body.requesterPhone || "");
      if (!requester) {
        return jsonResponse({ status: "error", code: "AUTH", message: "Your session is invalid. Please log in again." });
      }
      if (requester.role !== ROLE_ADMIN) {
        return jsonResponse({ status: "error", code: "FORBIDDEN", message: "Only the admin can add contacts directly." });
      }

      var phone = normalizePhone(body.phone || "");
      var name = String(body.name || "").trim();
      var sheetName = body.sheet || SHEET_CALLING;
      if (phone === "" || name === "") {
        return jsonResponse({ status: "error", code: "BAD_VALUE", message: "Name and Phone Number are required." });
      }

      var locked = lock.tryLock(30000);
      if (!locked) {
        return jsonResponse({ status: "error", code: "BUSY", message: "Server is busy, please try again." });
      }

      try {
        var doc = ensureSetup();
        var targetSheet = doc.getSheetByName(sheetName);
        if (!targetSheet) {
          return jsonResponse({ status: "error", code: "SHEET_NOT_FOUND", message: "Sheet not found: " + sheetName });
        }

        // Check if phone number is already registered in this sheet
        var isRegistered = false;
        var lastRow = targetSheet.getLastRow();
        if (lastRow > 1) {
          var phones = targetSheet.getRange(2, 2, lastRow - 1, 1).getValues();
          for (var p = 0; p < phones.length; p++) {
            if (normalizePhone(phones[p][0]) === phone) { isRegistered = true; break; }
          }
        }

        if (isRegistered) {
          return jsonResponse({ status: "error", code: "ALREADY_REGISTERED", message: "This phone number is already registered." });
        }

        var nextRow = targetSheet.getLastRow() + 1;
        var rowValues = [
          name,
          phone,
          "NA",
          '=COUNTIFS(\'Session History\'!D:D, $B' + nextRow + ', \'Session History\'!E:E, "Yes")',
          0,
          "",
          ""
        ];
        if (sheetName === SHEET_FESTIVAL) {
          rowValues.push(body.event || "");
        }
        targetSheet.appendRow(rowValues);
        try {
          targetSheet.getRange(nextRow, 2).setNumberFormat("@");
        } catch(formatErr) {
          // ignore typed columns constraint
        }
        applyValidations(doc, sheetName);

        SpreadsheetApp.flush();

        // Run auto assignment immediately if this is Festival Promotions
        if (sheetName === SHEET_FESTIVAL) {
          try {
            runAutoAssignment();
          } catch(e) {}
        }

        var cacheKey = CACHE_KEY_CALLING + "_" + sheetName.replace(/\s+/g, "_");
        CacheService.getScriptCache().remove(cacheKey);

        return jsonResponse(buildDataPayload(requester, sheetName));
      } finally {
        try { lock.releaseLock(); } catch (ignored) {}
      }
    }

    if ((body.action || "") !== "updateContact") {
      return jsonResponse({ status: "error", code: "UNKNOWN_ACTION", message: "Unknown action: " + body.action });
    }

    // Identify the requester — their role decides what they may edit
    var requester = findUser(body.requesterPhone || "");
    if (!requester) {
      return jsonResponse({ status: "error", code: "AUTH", message: "Your session is invalid. Please log in again." });
    }
    var isAdmin = requester.role === ROLE_ADMIN;

    var updates = body.updates || {};
    if (updates.hasOwnProperty("phone") || updates.hasOwnProperty("sessions")) {
      return jsonResponse({ status: "error", code: "FORBIDDEN", message: "Phone number and sessions can only be edited in the Google Sheet." });
    }
    if (!isAdmin && (updates.hasOwnProperty("assignedTo") || updates.hasOwnProperty("cultivatedBy"))) {
      return jsonResponse({ status: "error", code: "FORBIDDEN", message: "Only the admin can change assignments or cultivators." });
    }

    // Serialize writes: two people can save at the same moment
    var locked = lock.tryLock(30000);
    if (!locked) {
      return jsonResponse({ status: "error", code: "BUSY", message: "Server is busy, please try again." });
    }

    var doc = ensureSetup();
    var sheetName = body.sheet || SHEET_CALLING;
    var sheet = doc.getSheetByName(sheetName);
    var target = normalizePhone(body.phone || "");
    var lastRow = sheet.getLastRow();

    if (target.length < 10 || lastRow < 2) {
      return jsonResponse({ status: "error", code: "CONTACT_NOT_FOUND", message: "Contact not found" });
    }

    var fixed = (sheetName === SHEET_FESTIVAL) ? 8 : 7;
    var rows = sheet.getRange(2, 1, lastRow - 1, fixed).getValues();
    var rowIndex = -1;
    var rowData = null;
    for (var i = 0; i < rows.length; i++) {
      if (normalizePhone(rows[i][1]) === target) {
        rowIndex = i + 2; // account for header row
        rowData = rows[i];
        break;
      }
    }

    if (rowIndex === -1) {
      return jsonResponse({ status: "error", code: "CONTACT_NOT_FOUND", message: "Contact not found" });
    }

    // Non-admins may only touch contacts assigned to them
    if (!isAdmin && String(rowData[COL_ASSIGNED - 1]).trim().toLowerCase() !== requester.name.toLowerCase()) {
      return jsonResponse({ status: "error", code: "FORBIDDEN", message: "This contact is not assigned to you." });
    }

    // --- Apply updates (validated field by field) ---

    if (updates.hasOwnProperty("name")) {
      var newName = String(updates.name).trim();
      if (newName === "") {
        return jsonResponse({ status: "error", code: "BAD_VALUE", message: "Name cannot be empty." });
      }
      if (newName.length > 100) {
        return jsonResponse({ status: "error", code: "BAD_VALUE", message: "Name is too long (max 100 characters)." });
      }
      sheet.getRange(rowIndex, 1).setValue(newName);
    }
    var contactName = updates.hasOwnProperty("name")
      ? String(updates.name).trim()
      : String(rowData[0]).trim();

    if (updates.hasOwnProperty("ws")) {
      var ws = String(updates.ws).trim().toUpperCase();
      if (WS_OPTIONS.indexOf(ws) === -1) {
        return jsonResponse({ status: "error", code: "BAD_VALUE", message: "W/S must be one of: " + WS_OPTIONS.join(", ") });
      }
      sheet.getRange(rowIndex, COL_WS).setValue(ws);
    }

    if (updates.hasOwnProperty("status")) {
      var callStatus = String(updates.status).trim();
      // "Others" must arrive as the typed text, never as the bare word
      if (callStatus === "" || callStatus.toLowerCase() === "others") {
        return jsonResponse({ status: "error", code: "BAD_VALUE", message: "Please type the details for the Others status." });
      }
      if (callStatus.length > 200) {
        return jsonResponse({ status: "error", code: "BAD_VALUE", message: "Status text is too long (max 200 characters)." });
      }
      var active = getActiveDateColumn(sheet);
      if (active.activeCol === -1) {
        return jsonResponse({ status: "error", code: "NO_DATE_COLUMN", message: "No date column found in the sheet. Ask the admin to add one." });
      }
      sheet.getRange(rowIndex, active.activeCol).setValue(callStatus);
      // Every real call attempt is logged with its exact time ("Not Done" isn't a call)
      if (callStatus !== STATUS_DEFAULT) {
        logCallHistory(doc, active.activeDate, contactName, target, callStatus, requester.name);
      }
    }
    if (sheetName === SHEET_FESTIVAL && updates.hasOwnProperty("event")) {
      sheet.getRange(rowIndex, 8).setValue(String(updates.event).trim());
    }
    if (updates.hasOwnProperty("cultivatedBy")) {
      var cultivator = String(updates.cultivatedBy).trim();
      if (cultivator !== "") {
        var allUsers = getUsers(true);
        var match = null;
        for (var u = 0; u < allUsers.length; u++) {
          if (allUsers[u].name.toLowerCase() === cultivator.toLowerCase()) { match = allUsers[u].name; break; }
        }
        if (!match) {
          return jsonResponse({ status: "error", code: "BAD_VALUE", message: '"' + cultivator + '" is not a registered user.' });
        }
        cultivator = match;
      }
      sheet.getRange(rowIndex, COL_CULTIVATED).setValue(cultivator);
    }

    if (updates.hasOwnProperty("assignedTo")) {
      var assignee = String(updates.assignedTo).trim();
      if (assignee !== "") {
        // Must match a registered user; store the canonical spelling
        var allUsers = getUsers(true);
        var match = null;
        for (var u = 0; u < allUsers.length; u++) {
          if (allUsers[u].name.toLowerCase() === assignee.toLowerCase()) { match = allUsers[u].name; break; }
        }
        if (!match) {
          return jsonResponse({ status: "error", code: "BAD_VALUE", message: '"' + assignee + '" is not a registered user.' });
        }
        assignee = match;
      }
      sheet.getRange(rowIndex, COL_ASSIGNED).setValue(assignee);
    }

    SpreadsheetApp.flush();
    var cacheKey = CACHE_KEY_CALLING + "_" + sheetName.replace(/\s+/g, "_");
    CacheService.getScriptCache().remove(cacheKey);

    // Return the requester's fresh view of the data (includes updated call counts)
    return jsonResponse(buildDataPayload(requester, sheetName));

  } catch (error) {
    return jsonResponse({ status: "error", code: "SERVER_ERROR", message: error.toString() });
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function getCurrentFestival() {
  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName(SHEET_USERS);
    if (!sheet) return "";
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return "";
    
    // Read F (Festival), G (From Date), H (To Date) -> columns 6, 7, 8
    var values = sheet.getRange(2, 6, lastRow - 1, 3).getValues();
    var today = new Date();
    today.setHours(0,0,0,0);
    
    for (var i = 0; i < values.length; i++) {
      var fest = String(values[i][0]).trim();
      var fromVal = values[i][1];
      var toVal = values[i][2];
      
      if (fest !== "" && fromVal && toVal) {
        var fromDate = new Date(fromVal);
        var toDate = new Date(toVal);
        if (!isNaN(fromDate.getTime()) && !isNaN(toDate.getTime())) {
          fromDate.setHours(0,0,0,0);
          toDate.setHours(23,59,59,999);
          if (today >= fromDate && today <= toDate) {
            return fest;
          }
        }
      }
    }
  } catch(e) {}
  return "";
}

function runAutoAssignment() {
  var allUsers = getUsers(false);
  var activeAssignees = allUsers.filter(function(u) { return u.autoAssign; });
  if (activeAssignees.length === 0) return;

  var userPhones = {};
  allUsers.forEach(function(u) {
    if (u.phone) {
      userPhones[normalizePhone(u.phone)] = true;
    }
  });

  var lock = LockService.getScriptLock();
  var locked = lock.tryLock(10000);
  if (!locked) return;

  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName("Festival Promotions");
    if (!sheet) return;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // Count current assignments for all active auto-assign users
    var userCounts = {};
    activeAssignees.forEach(function(u) { userCounts[u.name] = 0; });

    var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var r = 0; r < values.length; r++) {
      var ass = String(values[r][6]).trim(); // Column G = Assigned To
      if (userCounts.hasOwnProperty(ass)) {
        userCounts[ass]++;
      }
    }

    var assignedAny = false;
    for (var r = 0; r < values.length; r++) {
      var ass = String(values[r][6]).trim();
      if (ass === "") {
        // Skip auto-assignment if the contact is a registered caller (user)
        var contactPhone = normalizePhone(values[r][1]);
        if (userPhones.hasOwnProperty(contactPhone)) {
          continue;
        }

        var eventName = String(values[r][7]).trim(); // Column H = Event
        if (eventName === "") {
          var curFest = getCurrentFestival();
          if (curFest !== "") {
            eventName = curFest;
            sheet.getRange(r + 2, 8).setValue(curFest);
          }
        }

        // Filter eligible users who match this contact's event
        var eligibleUsers = [];
        activeAssignees.forEach(function(u) {
          // A user is eligible if they are assigned to this event specifically, or if their festival filter is blank/All
          if (u.festival === eventName || u.festival === "") {
            eligibleUsers.push(u.name);
          }
        });

        if (eligibleUsers.length === 0) continue; // Skip assignment if no callers are mapped to this festival

        // Pick eligible user with minimum assigned count
        var minUser = eligibleUsers[0];
        var minCount = userCounts[minUser] || 0;
        for (var u = 1; u < eligibleUsers.length; u++) {
          var currUser = eligibleUsers[u];
          var currCount = userCounts[currUser] || 0;
          if (currCount < minCount) {
            minUser = currUser;
            minCount = currCount;
          }
        }
        sheet.getRange(r + 2, 7).setValue(minUser);
        userCounts[minUser] = (userCounts[minUser] || 0) + 1;
        assignedAny = true;
      }
    }

    if (assignedAny) {
      SpreadsheetApp.flush();
      // Clear cache
      var cacheKey = CACHE_KEY_CALLING + "_" + "Festival Promotions".replace(/\s+/g, "_");
      CacheService.getScriptCache().remove(cacheKey);
    }
  } finally {
    lock.releaseLock();
  }
}
function getFestivalForDate(targetDate) {
  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName(SHEET_USERS);
    if (!sheet) return "";
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return "";
    
    // Read columns A to H (1 to 8) to parse festival name from Column A and From/To Dates from Columns G & H
    var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var checkDate = new Date(targetDate);
    checkDate.setHours(0,0,0,0);
    
    for (var i = 0; i < values.length; i++) {
      var role = String(values[i][2]).trim().toLowerCase();
      if (role === "festivals") {
        var fest = String(values[i][0]).trim(); // Column A (Name)
        var fromVal = values[i][6]; // Column G (7)
        var toVal = values[i][7]; // Column H (8)
        
        if (fest !== "" && fromVal && toVal) {
          var fromDate = new Date(fromVal);
          var toDate = new Date(toVal);
          if (!isNaN(fromDate.getTime()) && !isNaN(toDate.getTime())) {
            fromDate.setHours(0,0,0,0);
            toDate.setHours(23,59,59,999);
            if (checkDate >= fromDate && checkDate <= toDate) {
              return fest;
            }
          }
        }
      }
    }
  } catch(e) {}
  return "";
}

function processFormResponses() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = doc.getSheetByName("Form Responces 1") || doc.getSheetByName("Form Responses 1");
  if (!formSheet) return;
  
  var festivalSheet = doc.getSheetByName(SHEET_FESTIVAL);
  if (!festivalSheet) return;
  
  var lastRow = formSheet.getLastRow();
  if (lastRow < 2) return;
  
  // Find Name, Phone and Timestamp columns by scanning header
  var headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
  var nameCol = 2; // default B
  var phoneCol = 3; // default C
  var timestampCol = 1; // default A
  var foundYourName = false;
  
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase().trim();
    if (h === "your name" || h.indexOf("your name") !== -1) {
      nameCol = i + 1;
      foundYourName = true;
    } else if (h.indexOf("name") !== -1 && !foundYourName) {
      nameCol = i + 1;
    } else if (h.indexOf("phone") !== -1 || h.indexOf("number") !== -1 || h.indexOf("mobile") !== -1) {
      phoneCol = i + 1;
    } else if (h.indexOf("timestamp") !== -1 || h.indexOf("time") !== -1) {
      timestampCol = i + 1;
    }
  }
  
  // Get all registered user phones to skip processing them as contacts
  var allUsers = getUsers(true);
  var userPhones = {};
  allUsers.forEach(function(u) {
    if (u.phone) {
      userPhones[normalizePhone(u.phone)] = true;
    }
  });

  // Get all existing phones in Festival Promotions to skip duplicates
  var existingPhones = {};
  var fLast = festivalSheet.getLastRow();
  if (fLast > 1) {
    var fPhones = festivalSheet.getRange(2, 2, fLast - 1, 1).getValues();
    for (var j = 0; j < fPhones.length; j++) {
      existingPhones[normalizePhone(fPhones[j][0])] = true;
    }
  }
  
  // Read all form responses
  var formValues = formSheet.getRange(2, 1, lastRow - 1, formSheet.getLastColumn()).getValues();
  var copiedAny = false;
  
  for (var r = 0; r < formValues.length; r++) {
    var name = String(formValues[r][nameCol - 1]).trim();
    var phone = normalizePhone(formValues[r][phoneCol - 1]);
    if (name === "" || phone === "" || phone.length !== 10) continue;
    
    // Skip if the submitter's phone number belongs to one of our registered callers
    if (userPhones.hasOwnProperty(phone)) continue;

    // Skip if already in Festival Promotions
    if (existingPhones.hasOwnProperty(phone)) continue;
    
    // Parse timestamp to find correct festival
    var tsVal = formValues[r][timestampCol - 1];
    var tsDate = new Date(tsVal);
    if (isNaN(tsDate.getTime())) {
      tsDate = new Date(); // fallback to today
    }
    
    // Find matching festival based on the timestamp date
    var eventName = getFestivalForDate(tsDate);
    
    // Append to Festival Promotions
    var nextRow = festivalSheet.getLastRow() + 1;
    var rowValues = [
      name,
      phone,
      "NA",
      '=COUNTIFS(\'Session History\'!D:D, $B' + nextRow + ', \'Session History\'!E:E, "Yes")',
      0,
      "",
      "", // Assigned To (left blank for auto-assignment to pick up)
      eventName
    ];
    festivalSheet.appendRow(rowValues);
    try {
      festivalSheet.getRange(nextRow, 2).setNumberFormat("@");
    } catch(e) {}
    
    existingPhones[phone] = true;
    copiedAny = true;
  }
  
  if (copiedAny) {
    SpreadsheetApp.flush();
    // Clear cache for Festival Promotions
    var cacheKey = CACHE_KEY_CALLING + "_" + SHEET_FESTIVAL.replace(/\s+/g, "_");
    CacheService.getScriptCache().remove(cacheKey);
  }
}

function getSettings() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("app_settings");
  if (cached) {
    try { return JSON.parse(cached); } catch(e){}
  }

  var doc = ensureSetup();
  var sheet = doc.getSheetByName("Settings");
  var settings = {
    calling_message: "Hare Krishna {name}, please attend Thursday Session.",
    festival_message: "Hare Krishna {name}, you are invited to our upcoming festival.",
    calling_poster: "",
    festival_poster: ""
  };

  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var i = 0; i < values.length; i++) {
        var key = String(values[i][0]).trim();
        var val = String(values[i][1]).trim();
        if (key) {
          settings[key] = val;
        }
      }
    }
  }

  try {
    cache.put("app_settings", JSON.stringify(settings), 21600); // cache for 6 hours
  } catch(e){}
  return settings;
}

function saveSetting(key, value) {
  var doc = ensureSetup();
  var sheet = doc.getSheetByName("Settings");
  if (!sheet) {
    sheet = doc.insertSheet("Settings");
    sheet.appendRow(["Key", "Value"]);
    try { styleHeader(sheet.getRange(1, 1, 1, 2), "#0f766e"); } catch(e){}
  }

  var decodedValue = value;
  try {
    decodedValue = decodeURIComponent(value);
  } catch(e){}

  var lastRow = sheet.getLastRow();
  var foundRow = -1;
  if (lastRow > 1) {
    var keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === key) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow !== -1) {
    sheet.getRange(foundRow, 2).setValue(decodedValue);
  } else {
    sheet.appendRow([key, decodedValue]);
  }

  try {
    SpreadsheetApp.flush();
  } catch(e){}
  
  // Clear settings cache
  try {
    CacheService.getScriptCache().remove("app_settings");
  } catch(e){}
}

function saveSettingImage(key, base64Data, mimeType, fileName) {
  var doc = ensureSetup();
  var sheet = doc.getSheetByName("Settings");
  if (!sheet) {
    sheet = doc.insertSheet("Settings");
    sheet.appendRow(["Key", "Value"]);
    try { styleHeader(sheet.getRange(1, 1, 1, 2), "#0f766e"); } catch(e){}
  }

  // 1. Decode base64 bytes to blob
  var base64String = base64Data.split(",")[1] || base64Data;
  var bytes = Utilities.base64Decode(base64String);
  var blob = Utilities.newBlob(bytes, mimeType, fileName);

  // 2. Create file in user's Google Drive
  var file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 3. Generate direct download link
  var directUrl = "https://drive.google.com/uc?export=download&id=" + file.getId();

  // 4. Find key to update and delete older file if applicable
  var lastRow = sheet.getLastRow();
  var foundRow = -1;
  var oldFileId = "";
  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === key) {
        foundRow = i + 2;
        var oldUrl = String(values[i][1]).trim();
        if (oldUrl.indexOf("id=") !== -1) {
          oldFileId = oldUrl.split("id=")[1];
        }
        break;
      }
    }
  }

  if (oldFileId) {
    try {
      var oldFile = DriveApp.getFileById(oldFileId);
      oldFile.setTrashed(true);
    } catch(err){}
  }

  // 5. Save setting row value
  if (foundRow !== -1) {
    sheet.getRange(foundRow, 2).setValue(directUrl);
  } else {
    sheet.appendRow([key, directUrl]);
  }

  // Clear settings cache
  try {
    CacheService.getScriptCache().remove("app_settings");
  } catch(e){}

  return directUrl;
}
