/* ============================================================================
   IFAR DIGITAL · Google Apps Script backend  (Code.gs)
   ----------------------------------------------------------------------------
   CHANGELOG — user-login activation
   ----------------------------------------------------------------------------
   • EVERY account in the "Users" sheet can now log in:
       - usernames are normalised (trim / lower-case, spaces -> dots, duplicates
         automatically suffixed so nothing collides)
       - missing usernames are generated from the Full Name
       - plain-text passwords typed by hand into the sheet are detected and
         hashed automatically at login time
       - blank passwords get a readable TEMPORARY password (forced change on
         first login)
       - Role / Status / Email are normalised ("admin" -> "Admin", "active " ->
         "Active", unknown status -> Active)
       - legacy column names (Password, Name, E-mail, ...) are accepted
   • New self-service + admin credential tools:
       apiChangePassword · apiResetPassword · apiSetUserStatus · apiUnlockUser
       apiLogout · apiGetCredentialReport
       Spreadsheet menu "IFAR Digital" with Activate-all / Email-credentials /
       Add-user / Reset-password / Unlock actions.
   • Brute-force protection: 8 failed attempts -> 10 minute lock (auto-clears,
     admin can unlock instantly). Expired sessions are pruned automatically.
   • Existing behaviour (records, options, exports, audits) is unchanged.
   ============================================================================ */

var SS = SpreadsheetApp.getActiveSpreadsheet();
var DB_SHEET = 'Database';
var USERS_SHEET = 'Users';
var LOG_SHEET = 'AuditLogs';
var OPTIONS_SHEET = 'Options';
var SESSIONS_SHEET = 'Sessions';
var CLAIMS_SHEET = 'EditClaims';
var CRED_SHEET = 'CredentialReport';
var SESSION_HOURS = 12;
var CLAIM_HOURS = 2;        // how long a record stays "open for editing"
var MAX_SNAPSHOT_CHARS = 45000;

var VALID_ROLES = ['Admin', 'User', 'Viewer'];
var VALID_STATUS = ['Active', 'Inactive'];
var MAX_LOGIN_ATTEMPTS = 8;      // wrong passwords before a temporary lock
var LOCK_MINUTES = 10;           // how long the lock lasts
var MIN_PASSWORD_LENGTH = 6;     // policy for passwords users set themselves

/* Emergency administrator login. This account is authenticated directly here,
   never read from or written to the Users sheet. The requested password (6382)
   is stored as a SHA-256 hash rather than plain text. */
var EMERGENCY_ADMIN_USERNAME = 'a';
var EMERGENCY_ADMIN_PASSWORD_HASH = '8bceed0f1d74b8cd45d26c44973772d6ff9d850f86e74a4c7b0d561243d597a8';
var EMERGENCY_ADMIN_FULLNAME = 'Emergency Administrator';

/* Canonical Users-sheet columns. Older sheets keep their existing columns and
   simply get the missing ones appended on the right. */
var USERS_HEADERS = [
  'Username', 'Password Hash', 'Full Name', 'Role', 'Status',
  'Created At', 'Updated At', 'Email', 'Must Change Password',
  'Last Login', 'Failed Attempts', 'Locked Until'
];

/* Alternative header names that are accepted for each canonical column. */
var USER_HEADER_ALIASES = {
  'Username': ['Username', 'User Name', 'User', 'Login', 'Account'],
  'Password Hash': ['Password Hash', 'Password', 'Pass', 'Pwd'],
  'Full Name': ['Full Name', 'Fullname', 'Name', 'Employee Name'],
  'Role': ['Role', 'Access Level', 'User Role'],
  'Status': ['Status', 'Account Status', 'Active'],
  'Created At': ['Created At', 'Created', 'Date Created'],
  'Updated At': ['Updated At', 'Updated', 'Date Updated'],
  'Email': ['Email', 'E-mail', 'Email Address'],
  'Must Change Password': ['Must Change Password', 'Must Change', 'Change Password'],
  'Last Login': ['Last Login', 'Last Login At'],
  'Failed Attempts': ['Failed Attempts', 'Attempts'],
  'Locked Until': ['Locked Until', 'Locked']
};

/* Demo accounts created only when the Users sheet is completely empty. */
var DEFAULT_USERS = [
  ['admin', 'admin123', 'System Administrator', 'Admin', 'Active'],
  ['user1', 'user123', 'Field Officer 1', 'User', 'Active'],
  ['viewer', 'viewer123', 'Quality Auditor', 'Viewer', 'Active']
];

/* ============================================================================
   WEB APP ENTRY POINT
   ============================================================================ */
function doGet() {
  initializeIFAR();
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('IFAR DIGITAL')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initializeIFAR() {
  ensureUsers_();
  ensureLogs_();
  ensureSessions_();
  ensureClaims_();
  ensureOptions_();
  ensureDatabase_();
  return { ok: true, message: 'IFAR Database Initialized' };
}

/* ---------------------------------------------------------------------------
   Spreadsheet menu (visible when the file is opened in Google Sheets)
   --------------------------------------------------------------------------- */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('IFAR Digital')
      .addItem('🚀  Activate / repair ALL user logins', 'menuActivateAll')
      .addItem('🛡  Re-create default accounts (admin/user1/viewer)', 'menuSeedDefaults')
      .addItem('➕  Add / update one user', 'menuAddUser')
      //.addItem('🧱  Rebuild Database headers (backup + re-map)', 'menuRebuildSchema')
      //.addItem('🧹  Normalize Database (remove stale headers)', 'menuNormalizeDatabaseSchema')
      .addItem('🔑  Reset a user password', 'menuResetPassword')
      .addItem('🔓  Unlock a user', 'menuUnlockUser')
      .addSeparator()
      .addItem('📧  Email credentials to users with an e-mail', 'menuEmailCredentials')
      .addItem('👁  Show credential report sheet', 'menuShowReport')
      //.addItem('🧹  Clear credential report', 'menuClearReport')
      .addToUi();
  } catch (e) { /* onOpen also fires for API calls where there is no UI */ }
}

/* ============================================================================
   GENERIC SHEET HELPERS
   ============================================================================ */
function getOrCreate_(name) {
  var s = SS.getSheetByName(name);
  if (!s) s = SS.insertSheet(name);
  return s;
}

function formatHeader_(sheet, colCount) {
  if (!colCount || colCount < 1) return;
  sheet.getRange(1, 1, 1, colCount)
    .setFontWeight('bold')
    .setBackground('#16324f')
    .setFontColor('#ffffff')
    .setWrap(false);
  try { sheet.setFrozenRows(1); } catch (e) {}
}

/* ============================================================================
   COMPLETE DYNAMIC SCHEMA GENERATOR FOR ALL INPUT FIELDS (FORMS 01 - 04)
   ----------------------------------------------------------------------------
   GENERATED FILE - do not hand-edit the active UI column list.
   Source of truth = the web app's own combinedFlatten() output, extracted from
   ifarindex.html (see tools/extract_ui_headers.js + tools/gen_schema.py).
   Land Preparation labor columns were removed from the active schema; old
   values are migrated into FULL_RECORD_JSON before physical removal.
   Re-run:  python3 tools/gen_schema.py
   ============================================================================ */
function buildFullSchema_() {
  var H = [];
  /* --- Identity / record meta --- */
  H.push('FC Code', 'Farmer Name', 'Last Updated', 'Updated By');
  H.push('Revision', 'Crop Year', 'Home Address', 'Contract No.');
  H.push('Date Signed', 'FCV Variety', 'Burley Variety', 'Registered Area (ha)');
  H.push('Home Longitude', 'Home Latitude', 'Home Remark', 'Farm Lot 1 Longitude');
  H.push('Farm Lot 1 Latitude', 'Farm Lot 1 Remark', 'Farm Lot 2 Longitude', 'Farm Lot 2 Latitude');
  H.push('Farm Lot 2 Remark', 'Farm Lot 3 Longitude', 'Farm Lot 3 Latitude', 'Farm Lot 3 Remark');
  H.push('Barn 1 Longitude', 'Barn 1 Latitude', 'Barn 1 Remark', 'Barn 2 Longitude');
  H.push('Barn 2 Latitude', 'Barn 2 Remark', 'Barn 3 Longitude', 'Barn 3 Latitude');
  H.push('Barn 3 Remark', 'Barn 4 Longitude', 'Barn 4 Latitude', 'Barn 4 Remark');
  H.push('Barn 5 Longitude', 'Barn 5 Latitude', 'Barn 5 Remark');
  
  /* --- FORM 02: Seedling Care and Maintenance --- */
  H.push('Source of Seeds', 'Seedbed Type', 'Number of Packs', 'Gram per Pack', 'Pack No.');
  H.push('Lot No.', 'Number of Beds', 'Seedbed Size (m)', 'Number of Trays', 'Estimated Seedlings');
  H.push('SW Date 1', 'SW Date 2', 'SW Remarks');
  H.push('RS Date 1', 'RS Date 2', 'RS Remarks');
  H.push('BF Date 1', 'BF Date 2', 'BF Remarks');
  H.push('TF Date 1', 'TF Date 2', 'TF Remarks');
  H.push('CL Date 1', 'CL Date 2', 'CL Remarks');
  var colsFs = ['Date', 'Qty (kg/bags)', 'Unit', 'Remarks'];
  ['FS1', 'FS2'].forEach(function (p) {
    for (var c = 0; c < colsFs.length; c++) H.push(p + ' ' + colsFs[c]);
  });
  var colsCpa = ['Date', 'CPA Used', 'Dosage', 'PPE', 'Fam M No', 'Fam M Age', 'Fam F No', 'Fam F Age', 'FT M No', 'FT M Age', 'FT F No', 'FT F Age', 'Tmp M No', 'Tmp M Age', 'Tmp F No', 'Tmp F Age', 'Laborer No', 'Laborer Age', 'Total Labor'];
  ['C1', 'C2', 'C3'].forEach(function (p) {
    for (var c = 0; c < colsCpa.length; c++) H.push(p + ' ' + colsCpa[c]);
  });
  var colsWat = ['Date', 'PPE', 'Fam M No', 'Fam M Age', 'Fam F No', 'Fam F Age', 'FT M No', 'FT M Age', 'FT F No', 'FT F Age', 'Tmp M No', 'Tmp M Age', 'Tmp F No', 'Tmp F Age', 'Laborer No', 'Laborer Age', 'Total Labor'];
  ['WA1', 'WA2', 'WA3'].forEach(function (p) {
    for (var c = 0; c < colsWat.length; c++) H.push(p + ' ' + colsWat[c]);
  });
  var colsPrk = ['Date of Pricking', 'No. of Seedlings Pricked'];
  ['PRK1', 'PRK2', 'PRK3'].forEach(function (p) {
    for (var c = 0; c < colsPrk.length; c++) H.push(p + ' ' + colsPrk[c]);
  });
  ['WE1', 'WE2', 'WE3'].forEach(function (p) {
    for (var c = 0; c < colsWat.length; c++) H.push(p + ' ' + colsWat[c]);
  });
  var colsSbPrk = ['Seedbed', 'No. of Seedling Pricked'];
  ['SBP1', 'SBP2', 'SBP3'].forEach(function (p) {
    for (var c = 0; c < colsSbPrk.length; c++) H.push(p + ' ' + colsSbPrk[c]);
  });
  var colsPull = ['Date', 'No. of Seedlings', 'Disposal Method'];
  ['SP1', 'SP2'].forEach(function (p) {
    for (var c = 0; c < colsPull.length; c++) H.push(p + ' ' + colsPull[c]);
  });
  H.push('Destruction of Seed Bed After TP Date', 'Destruction of Seed Bed After TP Remarks');

  /* --- FORM 03: Land Prep, Transplanting, Volume, Sidedress, Hilling & Irrigation --- */
  /* Land Preparation: 3 initial rows (LP1, LP2, LP3), PPE removed, no labor */
  var colsLp = ['Date', 'Method', 'Model', 'HP', 'Fuel L'];
  ['LP1', 'LP2', 'LP3'].forEach(function (p) {
    for (var c = 0; c < colsLp.length; c++) H.push(p + ' ' + colsLp[c]);
  });
  var colsFert = ['Date', '18-46-0 Bags', '10-18-24 Bags', '21-0-0 Bags', '0-0-50 Bags', 'PPE', 'Fam M No', 'Fam M Age', 'Fam F No', 'Fam F Age', 'FT M No', 'FT M Age', 'FT F No', 'FT F Age', 'Tmp M No', 'Tmp M Age', 'Tmp F No', 'Tmp F Age', 'Total Labor'];
  ['BA1', 'BA2', 'BA3', 'SD1', 'SD2', 'SD3'].forEach(function (p) {
    for (var c = 0; c < colsFert.length; c++) H.push(p + ' ' + colsFert[c]);
  });
  var colsTp = ['Area Planted', 'Date', 'Seedlings Transplanted', 'Method', 'Spacing', 'PPE', 'Fam M No', 'Fam M Age', 'Fam F No', 'Fam F Age', 'FT M No', 'FT M Age', 'FT F No', 'FT F Age', 'Tmp M No', 'Tmp M Age', 'Tmp F No', 'Tmp F Age', 'Total Labor'];
  ['P1', 'P2', 'P3'].forEach(function (p) {
    for (var c = 0; c < colsTp.length; c++) H.push(p + ' ' + colsTp[c]);
  });
  /* Hilling: 3 initial rows (H1, H2, H3) */
  var colsHl = ['Date', 'Method', 'Model', 'HP', 'Fuel L', 'PPE', 'Fam M No', 'Fam M Age', 'Fam F No', 'Fam F Age', 'FT M No', 'FT M Age', 'FT F No', 'FT F Age', 'Tmp M No', 'Tmp M Age', 'Tmp F No', 'Tmp F Age', 'Total Labor'];
  ['H1', 'H2', 'H3'].forEach(function (p) {
    for (var c = 0; c < colsHl.length; c++) H.push(p + ' ' + colsHl[c]);
  });
  var colsIr = ['Date', 'Method', 'Water Source', 'Pump Model', 'HP', 'Gas L', 'Diesel L', 'Est Water Volume (cu.m)', 'PPE', 'Fam M No', 'Fam M Age', 'Fam F No', 'Fam F Age', 'FT M No', 'FT M Age', 'FT F No', 'FT F Age', 'Tmp M No', 'Tmp M Age', 'Tmp F No', 'Tmp F Age', 'Total Labor'];
  for (var n = 1; n <= 8; n++) {
    for (var c = 0; c < colsIr.length; c++) H.push('IR' + n + ' ' + colsIr[c]);
  }
  var colsVol = ['Plant Count', 'Effective Area (ha)', 'Estimated Yield (kg/ha)', 'Estimated Volume (kg)', 'Score Robustness', 'Score Disease/Insects', 'Score Cultural Practices', 'Score Weather', 'Total Score', 'Remarks'];
  ['V1', 'V2'].forEach(function (p) {
    for (var c = 0; c < colsVol.length; c++) H.push(p + ' ' + colsVol[c]);
  });
  H.push('V3 Plant Count', 'V3 Equivalent Area (ha)', 'V3 Score Robustness', 'V3 Score Disease/Insects');
  H.push('V3 Score Cultural Practices', 'V3 Score Weather', 'V3 Total Score', 'V3 Volume Estimate (kg)');
  H.push('V3 Bad Area', 'V3 Bad Area Est Yield (kg/ha)', 'V3 Estimated Salvage Volume (kg)', 'V3 Total 3rd Volume Estimate (kg)');

  /* --- FORM 04: Integrated Pest Management, Topping & Suckercide, Disposal --- */
  H.push('Plots Logged');
  var colsWk = ['Date', 'Plant Stage', 'Insect Scouting', 'Insect Found', 'ETL Reached', 'Manual Picking', 'CPA Used', 'Dosage', 'PPE', 'Warning Sign', 'Fam M No', 'Fam M Age', 'Fam F No', 'Fam F Age', 'FT M No', 'FT M Age', 'FT F No', 'FT F Age', 'Tmp M No', 'Tmp M Age', 'Tmp F No', 'Tmp F Age', 'Total Labor'];
  for (var n = 1; n <= 12; n++) {
    for (var c = 0; c < colsWk.length; c++) H.push('W' + n + ' ' + colsWk[c]);
  }
  /* Topping / Suckercide: 'Suckercide Used' (renamed from 'CPA Used') */
  var colsRd = ['Date', 'Chemical Y/N', 'Manual Y/N', 'Suckercide Used', 'Dosage', 'PPE', 'Fam M No', 'Fam M Age', 'Fam F No', 'Fam F Age', 'FT M No', 'FT M Age', 'FT F No', 'FT F Age', 'Tmp M No', 'Tmp M Age', 'Tmp F No', 'Tmp F Age', 'Total Labor'];
  for (var n = 1; n <= 12; n++) {
    for (var c = 0; c < colsRd.length; c++) H.push('R' + n + ' ' + colsRd[c]);
  }
  /* Section C: CPA Container and Fertilizer Bag Disposal/Retrieval */
  H.push('CPA: Punch hole then rinse and bury', 'CPA: Punch hole then rinse and bury - Remark', 'CPA: Punch hole then rinse await collection', 'CPA: Punch hole then rinse await collection - Remark');
  H.push('CPA: Deliver to buying station', 'CPA: Deliver to buying station - Remark', 'CPA: Collected by supplier', 'CPA: Collected by supplier - Remark');
  H.push('Fert Bag: Retained for own use', 'Fert Bag: Retained for own use - Remark', 'Fert Bag: Send to buying station for redemption', 'Fert Bag: Send to buying station for redemption - Remark');
  H.push('Fert Bag: Sold to third party', 'Fert Bag: Sold to third party - Remark');
  H.push('FULL_RECORD_JSON');
  H.push('SP3 Date', 'SP3 No. of Seedlings', 'SP3 Disposal Method');
  H.push('H1 Animal Y/N', 'H1 Mechanized Y/N', 'H2 Animal Y/N', 'H2 Mechanized Y/N');
  return H;
}

/* ============================================================================
   DATABASE HEADER REBUILD / AUDIT  (opt-in - nothing runs by itself)
   ----------------------------------------------------------------------------
   "Rebuild Database headers" in the IFAR Digital menu (or apiRebuildSchema(token)
   for an Admin) rewrites the Database header row into exactly the order the web
   app produces, without losing a single value:

     1. a timestamped backup copy of the Database sheet is made first,
     2. the existing header row is read and SCHEMA_RENAMES are applied,
     3. every data row is re-mapped BY NAME into the new column order,
     4. columns the UI no longer produces are kept at the end as LEGACY
        (pass { dropLegacy: true } to remove them instead),
     5. the header row is re-formatted and any surplus columns are trimmed.

   Saved records themselves are never touched: the app loads a record from the
   FULL_RECORD_JSON column, which is carried across like any other value.
   ============================================================================ */

/* old column name -> new column name (values are carried across on rebuild) */
var SCHEMA_RENAMES = {
  'Rounds Logged': 'Plots Logged',
  'FZ Date 1': 'BF Date 1',
  'FZ Date 2': 'BF Date 2',
  'FZ Remarks': 'BF Remarks',
  'R1 CPA Used': 'R1 Suckercide Used',
  'R2 CPA Used': 'R2 Suckercide Used',
  'R3 CPA Used': 'R3 Suckercide Used',
  'R4 CPA Used': 'R4 Suckercide Used',
  'R5 CPA Used': 'R5 Suckercide Used',
  'R6 CPA Used': 'R6 Suckercide Used',
  'R7 CPA Used': 'R7 Suckercide Used',
  'R8 CPA Used': 'R8 Suckercide Used',
  'R9 CPA Used': 'R9 Suckercide Used',
  'R10 CPA Used': 'R10 Suckercide Used',
  'R11 CPA Used': 'R11 Suckercide Used',
  'R12 CPA Used': 'R12 Suckercide Used'
};

function reverseRenames_() {
  var out = {};
  Object.keys(SCHEMA_RENAMES).forEach(function (old) { out[SCHEMA_RENAMES[old]] = old; });
  return out;
}

var LP_LABOR_SUFFIXES_ = [
  'Fam M No', 'Fam M Age', 'Fam F No', 'Fam F Age', 'FT M No', 'FT M Age',
  'FT F No', 'FT F Age', 'Tmp M No', 'Tmp M Age', 'Tmp F No', 'Tmp F Age',
  'Total Labor'
];
var LP_LABOR_KEYS_ = [
  ['famM', 'Fam M'], ['famF', 'Fam F'], ['ftM', 'FT M'], ['ftF', 'FT F'],
  ['tmpM', 'Tmp M'], ['tmpF', 'Tmp F']
];

/* Before a dropLegacy rebuild, move any physical LP labor values into the
   record JSON. Valid snapshots receive the values in s03.landPrep; invalid or
   incomplete snapshots receive a small marker that apiGetRecord() merges back
   into the flat legacy reconstruction. The timestamped backup still contains
   the complete original row. */
function migrateLandPreparationLaborIntoJson_(oldHeaders, oldRows) {
  var indices = Object.create(null);
  oldHeaders.forEach(function (header, index) {
    if (!header) return;
    if (!indices[header]) indices[header] = [];
    indices[header].push(index);
  });
  var jsonIndices = indices['FULL_RECORD_JSON'] || [];
  if (!jsonIndices.length) return { rows: 0, values: 0 };
  var fcIndex = (indices['FC Code'] || [])[0];
  var farmerIndex = (indices['Farmer Name'] || [])[0];
  var updatedIndex = (indices['Last Updated'] || [])[0];
  var migratedRows = 0, migratedValues = 0;

  function valueAt(row, name) {
    var list = indices[name] || [];
    for (var i = list.length - 1; i >= 0; i--) {
      var value = row[list[i]];
      if (value !== undefined && value !== null && String(value) !== '') return value;
    }
    return '';
  }
  function hasValue(value) { return value !== undefined && value !== null && String(value) !== ''; }

  for (var r = 0; r < oldRows.length; r++) {
    var labor = {}, any = false;
    ['LP1', 'LP2', 'LP3'].forEach(function (prefix) {
      LP_LABOR_SUFFIXES_.forEach(function (suffix) {
        var value = valueAt(oldRows[r], prefix + ' ' + suffix);
        labor[prefix + ' ' + suffix] = value;
        if (hasValue(value)) { any = true; migratedValues++; }
      });
    });
    if (!any) continue;

    var jsonIndex = jsonIndices[jsonIndices.length - 1];
    var parsed = null;
    var raw = oldRows[r][jsonIndex];
    if (raw) {
      try { parsed = JSON.parse(raw); } catch (e) { parsed = null; }
    }
    if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
      if (parsed.s03 && Array.isArray(parsed.s03.landPrep)) {
        while (parsed.s03.landPrep.length < 3) parsed.s03.landPrep.push({});
        for (var lp = 0; lp < 3; lp++) {
          var target = parsed.s03.landPrep[lp] || {};
          LP_LABOR_KEYS_.forEach(function (keyInfo) {
            var prefix = 'LP' + (lp + 1) + ' ';
            target[keyInfo[0] + '_n'] = valueAt(oldRows[r], prefix + keyInfo[1] + ' No');
            target[keyInfo[0] + '_a'] = valueAt(oldRows[r], prefix + keyInfo[1] + ' Age');
          });
          parsed.s03.landPrep[lp] = target;
        }
      } else {
        parsed.__legacyLpLabor = labor;
      }
      oldRows[r][jsonIndex] = JSON.stringify(parsed);
    } else {
      oldRows[r][jsonIndex] = JSON.stringify({
        v: 1,
        fc: fcIndex === undefined ? '' : String(oldRows[r][fcIndex] || ''),
        farmer: farmerIndex === undefined ? '' : String(oldRows[r][farmerIndex] || ''),
        updatedAt: updatedIndex === undefined ? '' : String(oldRows[r][updatedIndex] || ''),
        __legacyLpLabor: labor
      });
    }
    migratedRows++;
  }
  return { rows: migratedRows, values: migratedValues };
}

function rebuildDatabaseSchema(options) {
  options = options || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = getOrCreate_(DB_SHEET);
  var lastRow = s.getLastRow(), lastCol = s.getLastColumn();
  if (!lastRow) return { ok: false, message: 'The ' + DB_SHEET + ' sheet is empty - nothing to rebuild.' };

  var oldHeaders = s.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v || '').trim(); });
  var oldRows = lastRow > 1 ? s.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  /* ---- 1. backup ---------------------------------------------------------- */
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  var backupName = DB_SHEET + ' backup ' + stamp;
  var backup = ss.getSheetByName(backupName);
  if (!backup && !options.skipBackup) backup = s.copyTo(ss).setName(backupName);

  /* Migrate LP labor values before dropLegacy removes their physical columns. */
  var lpMigration = { rows: 0, values: 0 };
  if (options.dropLegacy) lpMigration = migrateLandPreparationLaborIntoJson_(oldHeaders, oldRows);

  /* ---- 2. new header list ------------------------------------------------- */
  var seen = {}, newHeaders = [];
  buildFullSchema_().forEach(function (h) {
    if (!seen[h]) { seen[h] = true; newHeaders.push(h); }
  });

  var legacy = [];
  if (!options.dropLegacy) {
    oldHeaders.forEach(function (h) {
      if (!h) return;
      if (seen[h] || seen[SCHEMA_RENAMES[h]]) return;
      if (legacy.indexOf(h) < 0) legacy.push(h);
    });
  }
  var finalHeaders = newHeaders.concat(legacy);

  /* ---- 3. map every old value across by column name -----------------------
     The live audit found duplicate header blocks. The current save code targets
     the last copy of a duplicate header, so rebuild from the RIGHTMOST nonblank
     duplicate value; this avoids throwing away the active copy. The backup made
     above still contains every original cell if a conflict needs review. */
  var oldIndices = {};
  oldHeaders.forEach(function (h, i) {
    if (!h) return;
    if (!oldIndices[h]) oldIndices[h] = [];
    oldIndices[h].push(i);
  });
  var back = reverseRenames_();
  var duplicateHeaders = Object.keys(oldIndices).filter(function (h) { return oldIndices[h].length > 1; });
  var duplicatePhysicalColumns = 0;
  duplicateHeaders.forEach(function (h) { duplicatePhysicalColumns += oldIndices[h].length - 1; });

  function hasValue_(value) {
    return value !== undefined && value !== null && String(value) !== '';
  }
  function indicesFor_(name) {
    var groups = [];
    if (oldIndices[name]) groups.push(oldIndices[name]);
    if (back[name] !== undefined && oldIndices[back[name]]) groups.push(oldIndices[back[name]]);
    return groups;
  }
  function valueFor(row, name) {
    var groups = indicesFor_(name);
    for (var g = 0; g < groups.length; g++) {
      var indices = groups[g];
      for (var i = indices.length - 1; i >= 0; i--) {
        if (hasValue_(row[indices[i]])) return row[indices[i]];
      }
    }
    return '';
  }

  var duplicateValueConflicts = 0, duplicateValueConflictSamples = [];
  for (var rr = 0; rr < oldRows.length; rr++) {
    for (var dh = 0; dh < duplicateHeaders.length; dh++) {
      var dupName = duplicateHeaders[dh], values = Object.create(null), distinct = 0;
      var dupIndices = oldIndices[dupName];
      for (var di = 0; di < dupIndices.length; di++) {
        var candidate = oldRows[rr][dupIndices[di]];
        if (!hasValue_(candidate)) continue;
        var candidateKey = String(candidate);
        if (!values[candidateKey]) { values[candidateKey] = true; distinct++; }
      }
      if (distinct > 1) {
        duplicateValueConflicts++;
        if (duplicateValueConflictSamples.length < 50) {
          duplicateValueConflictSamples.push({ row: rr + 2, header: dupName, positions: dupIndices.map(function (i) { return i + 1; }) });
        }
      }
    }
  }

  var outRows = [];
  for (var r = 0; r < oldRows.length; r++) {
    var src = oldRows[r], out = [];
    for (var c = 0; c < finalHeaders.length; c++) out.push(valueFor(src, finalHeaders[c]));
    outRows.push(out);
  }

  /* ---- 4. write back ------------------------------------------------------ */
  if (finalHeaders.length > s.getMaxColumns()) {
    s.insertColumnsAfter(s.getMaxColumns(), finalHeaders.length - s.getMaxColumns());
  }
  s.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
  var CHUNK = 200;
  for (var w = 0; w < outRows.length; w += CHUNK) {
    var part = outRows.slice(w, w + CHUNK);
    s.getRange(2 + w, 1, part.length, finalHeaders.length).setValues(part);
  }
  if (lastCol > finalHeaders.length) s.deleteColumns(finalHeaders.length + 1, lastCol - finalHeaders.length);
  formatHeader_(s, finalHeaders.length);

  /* ---- 5. report ---------------------------------------------------------- */
  var moved = 0;
  for (var c2 = 0; c2 < Math.min(oldHeaders.length, finalHeaders.length); c2++) {
    if (oldHeaders[c2] !== finalHeaders[c2]) moved++;
  }
  var result = {
    ok: true,
    backup: backup ? backupName : '(skipped)',
    rows: outRows.length,
    oldColumns: oldHeaders.length,
    newColumns: finalHeaders.length,
    legacyColumns: legacy,
    columnsAtANewPosition: moved,
    duplicateHeaderCount: duplicateHeaders.length,
    duplicatePhysicalColumns: duplicatePhysicalColumns,
    duplicateValueConflicts: duplicateValueConflicts,
    duplicateValueConflictSamples: duplicateValueConflictSamples,
    migratedLandPrepRows: lpMigration.rows,
    migratedLandPrepValues: lpMigration.values
  };
  try {
    logEx_(Session.getActiveUser().getEmail() || 'owner', 'REBUILD_SCHEMA', '',
           'Header row rebuilt: ' + oldHeaders.length + ' -> ' + finalHeaders.length + ' columns',
           { 'Backup': result.backup, 'Rows': result.rows, 'Legacy Columns': legacy.join(', ') });
  } catch (e) { /* logging is best-effort */ }
  return result;
}

function apiRebuildSchema(token, options) {
  var session = requireRole_(token, ['Admin']);
  var out = rebuildDatabaseSchema(options || {});
  out.by = session.username;
  return out;
}

/* Explicit cleanup path: normalize to the current schema and remove only
   headers that are not in buildFullSchema_(). A timestamped backup is still
   created first. The seven retained LEGACY columns remain because they are
   deliberately part of the generated schema. */
function apiNormalizeDatabaseSchema(token) {
  var session = requireRole_(token, ['Admin']);
  var out = rebuildDatabaseSchema({ dropLegacy: true });
  out.by = session.username;
  return out;
}

/* ============================================================================
   Spreadsheet menu entries
   ============================================================================ */
function menuRebuildSchema() {
  var ui = SpreadsheetApp.getUi();
  var go = ui.alert('Rebuild ' + DB_SHEET + ' headers?',
    'This rewrites the header row of "' + DB_SHEET + '" into the exact order produced by the web app.\n\n' +
    'A timestamped backup copy of the sheet is made first, and every existing value is re-mapped by column name, so no data is lost.\n\n' +
    'Saved records are unaffected - the app reads them from the FULL_RECORD_JSON column.\n\nContinue?',
    ui.ButtonSet.OK_CANCEL);
  if (go !== ui.Button.OK) return;
  var out = rebuildDatabaseSchema({});
  ui.alert('Database headers rebuilt',
    'Backup sheet : ' + out.backup + '\n' +
    'Rows re-mapped : ' + out.rows + '\n' +
    'Columns : ' + out.oldColumns + ' -> ' + out.newColumns + '\n' +
    'Legacy columns kept : ' + (out.legacyColumns ? out.legacyColumns.length : 0) + '\n' +
    'Columns that changed position : ' + out.columnsAtANewPosition,
    ui.ButtonSet.OK);
}

function menuNormalizeDatabaseSchema() {
  var ui = SpreadsheetApp.getUi();
  var go = ui.alert('Normalize ' + DB_SHEET + ' and remove stale headers?',
    'This makes a timestamped backup first, removes duplicate header copies,\n' +
    'renames Rounds Logged to Plots Logged, and removes headers not in the\n' +
    'current app schema. The seven retained LEGACY columns remain.\n\nContinue?',
    ui.ButtonSet.OK_CANCEL);
  if (go !== ui.Button.OK) return;
  var out = rebuildDatabaseSchema({ dropLegacy: true });
  ui.alert('Database normalized',
    'Backup sheet : ' + out.backup + '\n' +
    'Rows re-mapped : ' + out.rows + '\n' +
    'Columns : ' + out.oldColumns + ' -> ' + out.newColumns + '\n' +
    'Duplicate headers collapsed : ' + (out.duplicatePhysicalColumns || 0) + '\n' +
    'Conflicting duplicate values : ' + (out.duplicateValueConflicts || 0) + '\n' +
    'Migrated LP rows : ' + (out.migratedLandPrepRows || 0),
    ui.ButtonSet.OK);
}


function ensureDatabase_() {
  var s = getOrCreate_(DB_SHEET);
  var fullSchema = buildFullSchema_();

  if (s.getLastRow() === 0) {
    s.getRange(1, 1, 1, fullSchema.length).setValues([fullSchema]);
    formatHeader_(s, fullSchema.length);
    return s;
  }

  /* Append any missing schema headers by name. Never rewrite/reorder an existing
     database header row: this keeps legacy data columns stable while ensuring
     newly added fields (including Form 02 destruction and Hilling Method) exist. */
  var width = Math.max(1, s.getLastColumn());
  var currentHeader = s.getRange(1, 1, 1, width).getValues()[0];
  var present = Object.create(null);
  for (var i = 0; i < currentHeader.length; i++) {
    var existingName = String(currentHeader[i] === undefined || currentHeader[i] === null ? '' : currentHeader[i]).trim();
    if (existingName) present[existingName] = true;
  }
  var missing = [];
  for (var j = 0; j < fullSchema.length; j++) {
    var requiredName = String(fullSchema[j]);
    if (!present[requiredName]) {
      missing.push(requiredName);
      present[requiredName] = true;
    }
  }
  if (missing.length) {
    s.getRange(1, width + 1, 1, missing.length).setValues([missing]);
    formatHeader_(s, width + missing.length);
  }
  return s;
}

/* AuditLogs columns. The first five are the original ones - the rest are
   appended automatically, so an existing log keeps working. */
var AUDIT_HEADERS = ['Timestamp', 'Username', 'Action', 'FC Code', 'Details',
                     'Changed Fields', 'Field Count', 'Revision', 'Snapshot'];

function ensureLogs_() {
  var s = getOrCreate_(LOG_SHEET);
  if (s.getLastRow() === 0) {
    s.getRange(1, 1, 1, AUDIT_HEADERS.length).setValues([AUDIT_HEADERS]);
    formatHeader_(s, AUDIT_HEADERS.length);
    return s;
  }
  var width = Math.max(1, s.getLastColumn());
  var headers = s.getRange(1, 1, 1, width).getValues()[0];
  var have = {};
  for (var i = 0; i < headers.length; i++) have[String(headers[i]).trim().toLowerCase()] = i;
  var missing = [];
  for (var c = 0; c < AUDIT_HEADERS.length; c++) {
    if (have[AUDIT_HEADERS[c].toLowerCase()] === undefined) missing.push(AUDIT_HEADERS[c]);
  }
  if (missing.length) {
    s.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    formatHeader_(s, headers.length + missing.length);
  }
  return s;
}

function ensureSessions_() {
  var s = getOrCreate_(SESSIONS_SHEET);
  if (s.getLastRow() === 0) {
    var headers = ['Token Hash', 'Username', 'Role', 'Expires At', 'Created At'];
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader_(s, headers.length);
  }
  try { if (!s.isSheetHidden()) s.hideSheet(); } catch (e) {}
  return s;
}

function ensureOptions_() {
  var s = getOrCreate_(OPTIONS_SHEET);
  if (s.getLastRow() === 0) {
    var headers = ['Group', 'Value', 'Created By', 'Created At'];
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader_(s, headers.length);
  }
  return s;
}

/* ============================================================================
   PASSWORD / TOKEN CRYPTOGRAPHY
   ============================================================================ */
function hashPassword_(password) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password),
    Utilities.Charset.UTF_8
  );
  var out = [];
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] + 256) % 256;
    var h = b.toString(16);
    out.push(h.length === 1 ? '0' + h : h);
  }
  return out.join('');
}

function tokenHash_(token) { return hashPassword_(token); }
function randomToken_() { return Utilities.getUuid() + Utilities.getUuid(); }

function isHashed_(value) { return /^[0-9a-f]{64}$/i.test(String(value || '').trim()); }

/* Readable temporary password, e.g. "KXPD-4821" */
function randomTempPassword_() {
  var letters = 'ABCDEFGHJKLMNPQRSTUVWXYZ';   // no I / O
  var digits = '23456789';
  var a = '', b = '';
  for (var i = 0; i < 4; i++) a += letters.charAt(Math.floor(Math.random() * letters.length));
  for (var j = 0; j < 4; j++) b += digits.charAt(Math.floor(Math.random() * digits.length));
  return a + '-' + b;
}

/* ============================================================================
   USERS SHEET · HEADER RESOLUTION, NORMALISATION & ACTIVATION
   ============================================================================ */
function usersAliasList_(canonical) {
  return USER_HEADER_ALIASES[canonical] || [canonical];
}

/* Header row of the Users sheet, with every canonical column guaranteed to
   exist (missing ones are appended on the right). */
function ensureUsersHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, USERS_HEADERS.length).setValues([USERS_HEADERS]);
    formatHeader_(sheet, USERS_HEADERS.length);
    return USERS_HEADERS.slice();
  }
  var width = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  for (var i = 0; i < headers.length; i++) headers[i] = String(headers[i] || '').trim();

  var added = false;
  for (var c = 0; c < USERS_HEADERS.length; c++) {
    var canonical = USERS_HEADERS[c];
    if (findHeaderIndex_(headers, canonical) >= 0) continue;
    headers.push(canonical);
    added = true;
  }
  if (added) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader_(sheet, headers.length);
  }
  return headers;
}

function findHeaderIndex_(headers, canonical) {
  var aliases = usersAliasList_(canonical);
  for (var a = 0; a < aliases.length; a++) {
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim().toLowerCase() === aliases[a].toLowerCase()) return i;
    }
  }
  return -1;
}

/* Column map for the Users sheet: { 'Username': 0, 'Password Hash': 1, ... } */
function usersColumns_(headers) {
  var map = {};
  for (var c = 0; c < USERS_HEADERS.length; c++) {
    var canonical = USERS_HEADERS[c];
    var idx = findHeaderIndex_(headers, canonical);
    if (idx >= 0) map[canonical] = idx;
  }
  return map;
}

function usersTable_() {
  var sheet = getOrCreate_(USERS_SHEET);
  var headers = ensureUsersHeaders_(sheet);
  var width = headers.length;
  var rowCount = Math.max(sheet.getLastRow(), 1);
  var values = sheet.getRange(1, 1, rowCount, width).getValues();
  return { sheet: sheet, headers: headers, col: usersColumns_(headers), values: values, width: width };
}

function cellOf_(table, row, canonical) {
  var idx = table.col[canonical];
  if (idx === undefined) return '';
  var line = table.values[row];
  return line ? line[idx] : '';
}

function setCell_(table, row, canonical, value) {
  var idx = table.col[canonical];
  if (idx === undefined) return;
  if (!table.values[row]) table.values[row] = [];
  table.values[row][idx] = value;
}

function saveUsersTable_(table) {
  var rows = table.values;
  var last = rows.length;
  while (last > 1 && isBlankRow_(rows[last - 1])) last--;   // never leave thousands of empty rows
  var out = rows.slice(0, last);
  while (out.length < 1) out.push([]);
  table.sheet.getRange(1, 1, out.length, table.width).setValues(out);
}

function isBlankRow_(row) {
  if (!row) return true;
  for (var i = 0; i < row.length; i++) {
    if (String(row[i] === null || row[i] === undefined ? '' : row[i]).trim() !== '') return false;
  }
  return true;
}

/* --- normalisation helpers ------------------------------------------------- */
function normUsername_(value) {
  return String(value === null || value === undefined ? '' : value)
    .trim().toLowerCase().replace(/\s+/g, '.').replace(/[^a-z0-9._-]/g, '');
}

function usernameFromName_(fullName) {
  var base = String(fullName || '').trim().toLowerCase()
    .replace(/[^a-z0-9\s.]/g, ' ').replace(/\s+/g, '.').replace(/\.+/g, '.').replace(/^\.|\.$/g, '');
  return base || '';
}

function normRole_(value) {
  var v = String(value || '').trim().toLowerCase();
  if (v === 'admin' || v === 'administrator' || v === 'superadmin' || v === 'super admin') return 'Admin';
  if (v === 'viewer' || v === 'view' || v === 'read only' || v === 'readonly' || v === 'auditor') return 'Viewer';
  if (v === 'user' || v === 'staff' || v === 'encoder' || v === 'field officer') return 'User';
  if (v === '') return 'User';
  return 'User';
}

function normStatus_(value) {
  var v = String(value || '').trim().toLowerCase();
  if (v === '' || v === 'active' || v === 'enabled' || v === 'yes' || v === 'y' || v === 'true' || v === '1' || v === 'ok') return 'Active';
  if (v === 'inactive' || v === 'disabled' || v === 'no' || v === 'n' || v === 'false' || v === '0' || v === 'suspended' || v === 'locked') return 'Inactive';
  return 'Active';   // unknown wording -> activate instead of silently blocking
}

function normEmail_(value) {
  var v = String(value || '').trim().toLowerCase();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) ? v : '';
}

function truthy_(value) {
  var v = String(value === null || value === undefined ? '' : value).trim().toLowerCase();
  return v === 'true' || v === 'yes' || v === 'y' || v === '1' || v === 'x' || v === 'must change';
}

function toDate_(value) {
  if (!value) return null;
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  var d = new Date(String(value));
  return isNaN(d.getTime()) ? null : d;
}

/* --- the Users sheet, guaranteed usable ------------------------------------ */
function ensureUsers_() {
  var sheet = getOrCreate_(USERS_SHEET);
  ensureUsersHeaders_(sheet);
  if (sheet.getLastRow() < 2) seedDefaultUsers_(sheet);
  repairUsers_(sheet, { issueTempPasswords: false, quiet: true });
  return sheet;
}

function seedDefaultUsers_(sheet) {
  var table = usersTable_();
  var now = new Date().toISOString();
  for (var i = 0; i < DEFAULT_USERS.length; i++) {
    var d = DEFAULT_USERS[i];
    var row = [];
    for (var w = 0; w < table.width; w++) row.push('');
    setRowValue_(table, row, 'Username', d[0]);
    setRowValue_(table, row, 'Password Hash', hashPassword_(d[1]));
    setRowValue_(table, row, 'Full Name', d[2]);
    setRowValue_(table, row, 'Role', d[3]);
    setRowValue_(table, row, 'Status', d[4]);
    setRowValue_(table, row, 'Created At', now);
    setRowValue_(table, row, 'Updated At', now);
    setRowValue_(table, row, 'Must Change Password', false);
    table.values.push(row);
  }
  saveUsersTable_(table);
  log_('system', 'SEED_USERS', '', 'Default accounts created: admin / user1 / viewer');
  return table;
}

function setRowValue_(table, row, canonical, value) {
  var idx = table.col[canonical];
  if (idx === undefined) return;
  row[idx] = value;
}

/* ---------------------------------------------------------------------------
   ACTIVATION / REPAIR
   Every row that carries a person gets: a usable username, hashed password,
   a valid role and an Active status.  Runs on its own before every login
   (quiet mode) and from the menu (full mode, which also issues temporary
   passwords for rows that have none and writes the credential report).
   --------------------------------------------------------------------------- */
function repairUsers_(sheet, options) {
  options = options || {};
  var issueTempPasswords = options.issueTempPasswords !== false;
  var table = usersTable_();
  var now = new Date().toISOString();
  var report = [];
  var usedNames = {};
  var changed = false;

  for (var r = 1; r < table.values.length; r++) {
    var rowIsEmpty = isBlankRow_(table.values[r]);
    if (rowIsEmpty) continue;

    var before = snapshotUser_(table, r);

    var fullName = String(cellOf_(table, r, 'Full Name') || '').trim();
    var username = normUsername_(cellOf_(table, r, 'Username'));
    var fromName = false;
    if (!username) {
      username = usernameFromName_(fullName) || ('user' + (r + 1));
      fromName = true;
    }
    if (usedNames[username]) {
      var n = 2;
      while (usedNames[username + n]) n++;
      username = username + n;
      fromName = true;
    }
    usedNames[username] = true;

    var role = normRole_(cellOf_(table, r, 'Role'));
    var status = normStatus_(cellOf_(table, r, 'Status'));
    var email = normEmail_(cellOf_(table, r, 'Email'));
    var stored = String(cellOf_(table, r, 'Password Hash') || '').trim();
    var mustChange = truthy_(cellOf_(table, r, 'Must Change Password'));
    var tempPassword = '';
    var note = [];

    var plainLength = stored ? String(stored).length : 0;
    if (stored && isHashed_(stored)) {
      /* already hashed - keep the existing password */
    } else if (stored && plainLength >= MIN_PASSWORD_LENGTH) {
      stored = hashPassword_(stored);                     // plain text typed in the sheet -> hash it
      note.push('plain-text password hashed');
    } else if (issueTempPasswords) {
      tempPassword = randomTempPassword_();
      stored = hashPassword_(tempPassword);
      mustChange = true;
      note.push(plainLength ? 'weak password replaced by a temporary one' : 'temporary password issued');
    } else {
      if (!stored) mustChange = true;
      note.push(plainLength ? 'short password left as-is' : 'no password yet - use the IFAR Digital menu to issue one');
    }

    if (fromName) note.push('username generated');
    if (normStatus_(cellOf_(table, r, 'Status')) !== String(cellOf_(table, r, 'Status') || '').trim() && String(cellOf_(table, r, 'Status') || '').trim() !== '') {
      note.push('status normalised');
    }

    setCell_(table, r, 'Username', username);
    setCell_(table, r, 'Full Name', fullName);
    if (email) setCell_(table, r, 'Email', email);
    setCell_(table, r, 'Role', role);
    setCell_(table, r, 'Status', status);
    setCell_(table, r, 'Password Hash', stored);
    setCell_(table, r, 'Must Change Password', mustChange ? true : false);
    if (!cellOf_(table, r, 'Created At')) setCell_(table, r, 'Created At', now);
    setCell_(table, r, 'Updated At', now);
    if (cellOf_(table, r, 'Failed Attempts') === '' || cellOf_(table, r, 'Failed Attempts') === null || cellOf_(table, r, 'Failed Attempts') === undefined) {
      setCell_(table, r, 'Failed Attempts', 0);
    }

    var after = snapshotUser_(table, r);
    if (before !== after) changed = true;

    report.push({
      username: username,
      fullName: fullName,
      role: role,
      status: status,
      email: email,
      tempPassword: tempPassword,
      mustChange: !!mustChange,
      note: note.join('; ')
    });
  }

  if (changed) saveUsersTable_(table);
  if (report.length && !options.quiet) writeCredentialReport_(report);

  return { table: table, report: report, changed: changed };
}

function snapshotUser_(table, r) {
  return [
    cellOf_(table, r, 'Username'), cellOf_(table, r, 'Full Name'), cellOf_(table, r, 'Role'),
    cellOf_(table, r, 'Status'), cellOf_(table, r, 'Email'), cellOf_(table, r, 'Password Hash'),
    cellOf_(table, r, 'Must Change Password')
  ].join('|');
}

/* Temporary passwords already handed out, so they can still be looked up or
   re-sent as long as the user has not replaced them yet. */
function previousTempPasswords_() {
  var map = {};
  var s = SS.getSheetByName(CRED_SHEET);
  if (!s || s.getLastRow() < 2) return map;
  var rows = s.getDataRange().getValues();
  var headers = rows.shift();
  var iU = -1, iT = -1;
  for (var h = 0; h < headers.length; h++) {
    var name = String(headers[h]).trim();
    if (name === 'Username') iU = h;
    if (name === 'Temporary Password') iT = h;
  }
  if (iU < 0 || iT < 0) return map;
  for (var i = 0; i < rows.length; i++) {
    var u = String(rows[i][iU] || '').trim();
    var t = String(rows[i][iT] || '').trim();
    if (u && TEMP_PASSWORD_PATTERN.test(t)) map[u] = t;
  }
  return map;
}

var TEMP_PASSWORD_PATTERN = /^[A-Z]{4}-[0-9]{4}$/;

/* Store a temporary password in the credential report so the administrator can
   look it up or e-mail it later (used by create-user and reset-password). */
function rememberTempPassword_(username, temp, fullName, role, email) {
  var headers = ['Generated At', 'Username', 'Full Name', 'Role', 'Status', 'Email', 'Temporary Password', 'Must Change Password', 'Note'];
  var s = getOrCreate_(CRED_SHEET);
  if (s.getLastRow() === 0) {
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader_(s, headers.length);
  }
  var rows = s.getDataRange().getValues();
  var idx = {};
  for (var i = 0; i < rows[0].length; i++) idx[String(rows[0][i]).trim()] = i;
  var now = new Date().toISOString();
  for (var r = 1; r < rows.length; r++) {
    if (String(rows[r][idx['Username']] || '').trim().toLowerCase() === String(username).toLowerCase()) {
      if (idx['Temporary Password'] !== undefined) s.getRange(r + 1, idx['Temporary Password'] + 1).setValue(temp);
      if (idx['Must Change Password'] !== undefined) s.getRange(r + 1, idx['Must Change Password'] + 1).setValue('YES');
      if (idx['Generated At'] !== undefined) s.getRange(r + 1, idx['Generated At'] + 1).setValue(now);
      if (email && idx['Email'] !== undefined) s.getRange(r + 1, idx['Email'] + 1).setValue(email);
      return s;
    }
  }
  var line = [];
  for (var c = 0; c < headers.length; c++) line.push('');
  line[idx['Generated At']] = now;
  line[idx['Username']] = username;
  line[idx['Full Name']] = fullName || '';
  line[idx['Role']] = role || 'User';
  line[idx['Status']] = 'Active';
  line[idx['Email']] = email || '';
  line[idx['Temporary Password']] = temp;
  line[idx['Must Change Password']] = 'YES';
  line[idx['Note']] = 'temporary password issued';
  s.getRange(s.getLastRow() + 1, 1, 1, headers.length).setValues([line]);
  try { if (!s.isSheetHidden()) s.hideSheet(); } catch (e) {}
  return s;
}

/* Credential report sheet (hidden) - the list you hand out to the staff. */
function writeCredentialReport_(report) {
  var s = getOrCreate_(CRED_SHEET);
  var carried = previousTempPasswords_();
  s.clear();
  var headers = ['Generated At', 'Username', 'Full Name', 'Role', 'Status', 'Email', 'Temporary Password', 'Must Change Password', 'Note'];
  var now = new Date().toISOString();
  var rows = [headers];
  for (var i = 0; i < report.length; i++) {
    var r = report[i];
    var shown = r.tempPassword;
    if (!shown) {
      if (r.mustChange && carried[r.username]) shown = carried[r.username] + '  (still valid)';
      else if (r.mustChange) shown = 'unknown - use “Reset password”';
      else shown = 'personal password (not stored)';
    }
    rows.push([now, r.username, r.fullName, r.role, r.status, r.email || '', shown, r.mustChange ? 'YES' : 'no', r.note || '']);
  }
  s.getRange(1, 1, rows.length, headers.length).setValues(rows);
  formatHeader_(s, headers.length);
  s.autoResizeColumns(1, headers.length);
  try { if (!s.isSheetHidden()) s.hideSheet(); } catch (e) {}
  return s;
}

function setReportVisibility_(visible) {
  var s = SS.getSheetByName(CRED_SHEET);
  if (!s) return;
  try { if (visible) s.showSheet(); else s.hideSheet(); } catch (e) {}
}

/* --- user lookup / reading ------------------------------------------------- */
function readUser_(table, row) {
  return {
    row: row,
    username: String(cellOf_(table, row, 'Username') || '').trim(),
    passwordHash: String(cellOf_(table, row, 'Password Hash') || '').trim(),
    fullname: String(cellOf_(table, row, 'Full Name') || '').trim(),
    role: normRole_(cellOf_(table, row, 'Role')),
    roleRaw: String(cellOf_(table, row, 'Role') || '').trim(),
    status: normStatus_(cellOf_(table, row, 'Status')),
    statusRaw: String(cellOf_(table, row, 'Status') || '').trim(),
    email: normEmail_(cellOf_(table, row, 'Email')),
    mustChange: truthy_(cellOf_(table, row, 'Must Change Password')),
    lastLogin: toDate_(cellOf_(table, row, 'Last Login')),
    failed: Number(cellOf_(table, row, 'Failed Attempts')) || 0,
    lockedUntil: toDate_(cellOf_(table, row, 'Locked Until'))
  };
}

/* A user may log in with the username, the e-mail address or the full name. */
function findUserRow_(table, identifier) {
  var raw = String(identifier || '').trim();
  if (!raw) return -1;
  var wanted = normUsername_(raw);
  var plain = raw.toLowerCase();
  var matchName = -1;
  for (var r = 1; r < table.values.length; r++) {
    if (isBlankRow_(table.values[r])) continue;
    var u = readUser_(table, r);
    if (!u.username && !u.email) continue;
    if (u.username === wanted || u.username === plain) return r;
    if (u.email && u.email === plain) return r;
    if (matchName < 0 && u.fullname && u.fullname.toLowerCase() === plain) matchName = r;
  }
  return matchName;
}

function passwordMatches_(user, password) {
  var stored = String(user.passwordHash || '');
  if (!stored) return false;
  var attempts = [String(password), String(password).trim()];
  for (var i = 0; i < attempts.length; i++) {
    var p = attempts[i];
    if (p === '') continue;
    if (isHashed_(stored)) {
      if (stored.toLowerCase() === hashPassword_(p).toLowerCase()) return true;
    } else if (stored === p) {
      return true;   // legacy plain-text password in the sheet
    }
  }
  return false;
}

/* ============================================================================
   AUTHENTICATION API
   ============================================================================ */
function isEmergencyAdmin_(username) {
  return normUsername_(username) === EMERGENCY_ADMIN_USERNAME;
}

/* Create a normal server-side session without looking up a Users-sheet row. */
function emergencyAdminLogin_() {
  var token = randomToken_();
  var expires = new Date(Date.now() + SESSION_HOURS * 60 * 60 * 1000);
  var sheet = ensureSessions_();
  pruneSessions_(sheet);
  sheet.appendRow([
    tokenHash_(token), EMERGENCY_ADMIN_USERNAME, 'Admin',
    expires.toISOString(), new Date().toISOString()
  ]);
  log_(EMERGENCY_ADMIN_USERNAME, 'LOGIN', '', 'Successful emergency administrator login');
  return {
    ok: true,
    token: token,
    user: {
      username: EMERGENCY_ADMIN_USERNAME,
      fullname: EMERGENCY_ADMIN_FULLNAME,
      role: 'Admin',
      email: '',
      mustChangePassword: false
    }
  };
}

function apiLogin(username, password) {
  initializeIFAR();
  var identifier = String(username || '').trim();
  var pass = String(password === null || password === undefined ? '' : password);
  if (!identifier) throw new Error('Username and password are required.');
  if (!pass) throw new Error('Username and password are required.');

  /* Reserved emergency login: no Users-sheet lookup, status check, password
     length rule, failed-attempt counter or first-login password change. */
  if (isEmergencyAdmin_(identifier)) {
    if (hashPassword_(pass) !== EMERGENCY_ADMIN_PASSWORD_HASH) {
      log_(EMERGENCY_ADMIN_USERNAME, 'LOGIN_FAILED', '', 'Wrong emergency administrator password');
      throw new Error('Invalid username or password.');
    }
    return emergencyAdminLogin_();
  }

  var table = usersTable_();
  var row = findUserRow_(table, identifier);
  if (row < 0) throw new Error('Invalid username or password.');

  var user = readUser_(table, row);

  if (user.lockedUntil && user.lockedUntil.getTime() > Date.now()) {
    var mins = Math.max(1, Math.ceil((user.lockedUntil.getTime() - Date.now()) / 60000));
    throw new Error('Too many failed attempts. This account is locked for ' + mins + ' more minute(s).');
  }

  if (!user.username) {
    user.username = normUsername_(identifier);
    setCell_(table, row, 'Username', user.username);
    saveUsersTable_(table);
  }

  if (user.status !== 'Active') {
    throw new Error('This account is inactive. Please contact the administrator.');
  }

  if (!user.passwordHash) {
    throw new Error('This account has no password yet. Please ask the administrator to run “Activate / repair ALL user logins”.');
  }

  if (user.passwordHash && !isHashed_(user.passwordHash)) {
    /* upgrade a legacy plain-text password the moment it is verified */
    if (passwordMatches_(user, pass)) {
      setCell_(table, row, 'Password Hash', hashPassword_(pass));
      log_(user.username, 'PASSWORD_UPGRADED', '', 'Plain-text password replaced by a hash');
    }
  }

  if (!passwordMatches_(user, pass)) {
    registerFailedAttempt_(table, row, user);
    log_(user.username, 'LOGIN_FAILED', '', 'Wrong password (attempt ' + (user.failed + 1) + ')');
    throw new Error('Invalid username or password.');
  }

  /* --- success --- */
  var token = randomToken_();
  var expires = new Date(Date.now() + SESSION_HOURS * 60 * 60 * 1000);
  var sheet = ensureSessions_();
  pruneSessions_(sheet);
  sheet.appendRow([tokenHash_(token), user.username, user.role, expires.toISOString(), new Date().toISOString()]);

  setCell_(table, row, 'Failed Attempts', 0);
  setCell_(table, row, 'Locked Until', '');
  setCell_(table, row, 'Last Login', new Date().toISOString());
  saveUsersTable_(table);

  log_(user.username, 'LOGIN', '', 'Successful login as ' + user.role);

  return {
    ok: true,
    token: token,
    user: {
      username: user.username,
      fullname: user.fullname || user.username,
      role: user.role,
      email: user.email || '',
      mustChangePassword: !!user.mustChange
    }
  };
}

/* Called by the app when it re-opens on a device that already logged in
   (remember-me). Re-validates the token and refreshes role / flags. */
function apiPing(token) {
  initializeIFAR();
  var session = getSession_(token);
  var fullname = session.username, role = session.role, must = false, email = '';

  if (isEmergencyAdmin_(session.username)) {
    /* Do not let a same-named sheet row change or disable this built-in Admin. */
    fullname = EMERGENCY_ADMIN_FULLNAME;
    role = 'Admin';
  } else {
    var table = usersTable_();
    var row = findUserRow_(table, session.username);
    if (row > 0) {
      var u = readUser_(table, row);
      fullname = u.fullname || fullname;
      role = u.role || role;
      must = !!u.mustChange;
      email = u.email || '';
      if (u.status !== 'Active') throw new Error('This account is inactive. Please contact the administrator.');
    }
  }
  return { ok: true, user: { username: session.username, fullname: fullname, role: role, email: email, mustChangePassword: must } };
}

function registerFailedAttempt_(table, row, user) {
  var failed = (user.failed || 0) + 1;
  setCell_(table, row, 'Failed Attempts', failed);
  if (failed >= MAX_LOGIN_ATTEMPTS) {
    setCell_(table, row, 'Locked Until', new Date(Date.now() + LOCK_MINUTES * 60000).toISOString());
    setCell_(table, row, 'Failed Attempts', 0);
  }
  setCell_(table, row, 'Updated At', new Date().toISOString());
  saveUsersTable_(table);
}

function pruneSessions_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  var values = sheet.getRange(2, 1, last - 1, 5).getValues();
  var keep = [];
  for (var i = 0; i < values.length; i++) {
    var exp = toDate_(values[i][3]);
    if (exp && exp.getTime() > Date.now()) keep.push(values[i]);
    else if (!exp) keep.push(values[i]);
  }
  if (keep.length === values.length) return;
  sheet.getRange(2, 1, values.length, 5).clearContent();
  if (keep.length) sheet.getRange(2, 1, keep.length, 5).setValues(keep);
}

/* ============================================================================
   EDIT CLAIMS - "once saved it is fixed, unless it is loaded for editing"
   ----------------------------------------------------------------------------
   Loading a record issues a short-lived claim token (FC + user + TTL). A save
   that would overwrite an existing record is refused unless the caller proves
   it holds the claim for that FC. Claims are stored hidden and hashed.
   ============================================================================ */
function ensureClaims_() {
  var s = getOrCreate_(CLAIMS_SHEET);
  if (s.getLastRow() === 0) {
    var headers = ['FC Code', 'Username', 'Token Hash', 'Issued At', 'Expires At'];
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader_(s, headers.length);
  }
  try { if (!s.isSheetHidden()) s.hideSheet(); } catch (e) {}
  return s;
}

function readClaims_() {
  var sheet = ensureClaims_();
  var last = sheet.getLastRow();
  if (last < 2) return { sheet: sheet, rows: [] };
  return { sheet: sheet, rows: sheet.getRange(2, 1, last - 1, 5).getValues() };
}

function pruneClaims_(table) {
  table = table || readClaims_();
  if (!table.rows.length) return table;
  var now = Date.now();
  var keep = [];
  for (var i = 0; i < table.rows.length; i++) {
    var exp = toDate_(table.rows[i][4]);
    if (!exp || exp.getTime() > now) keep.push(table.rows[i]);
  }
  if (keep.length === table.rows.length) return table;
  table.sheet.getRange(2, 1, table.rows.length, 5).clearContent();
  if (keep.length) table.sheet.getRange(2, 1, keep.length, 5).setValues(keep);
  table.rows = keep;
  return table;
}

function writeClaims_(table, rows) {
  table.sheet.getRange(2, 1, Math.max(1, table.rows.length), 5).clearContent();
  if (rows.length) table.sheet.getRange(2, 1, rows.length, 5).setValues(rows);
}

function getClaim_(fc) {
  var table = pruneClaims_();
  var target = String(fc || '').trim().toUpperCase();
  for (var i = table.rows.length - 1; i >= 0; i--) {
    if (String(table.rows[i][0]).trim().toUpperCase() !== target) continue;
    return {
      fc: String(table.rows[i][0]),
      username: String(table.rows[i][1]),
      hash: String(table.rows[i][2]),
      expires: toDate_(table.rows[i][4])
    };
  }
  return null;
}

/* Issues (or refreshes) the claim for this FC and returns the raw token. */
function issueClaim_(fc, username) {
  var token = randomToken_();
  var table = pruneClaims_();
  var now = new Date();
  var expires = new Date(now.getTime() + CLAIM_HOURS * 60 * 60 * 1000);
  var fresh = [String(fc).trim(), username, tokenHash_(token), now.toISOString(), expires.toISOString()];
  var target = String(fc || '').trim().toUpperCase();
  var keep = [];
  var placed = false;
  for (var i = 0; i < table.rows.length; i++) {
    if (String(table.rows[i][0]).trim().toUpperCase() === target) {
      if (!placed) { keep.push(fresh); placed = true; }
    } else {
      keep.push(table.rows[i]);
    }
  }
  if (!placed) keep.push(fresh);
  writeClaims_(table, keep);
  return token;
}

function removeClaim_(fc) {
  var table = readClaims_();
  if (!table.rows.length) return;
  var target = String(fc || '').trim().toUpperCase();
  var keep = [];
  for (var i = 0; i < table.rows.length; i++) {
    if (String(table.rows[i][0]).trim().toUpperCase() !== target) keep.push(table.rows[i]);
  }
  if (keep.length === table.rows.length) return;
  writeClaims_(table, keep);
}

/* Does this user hold a valid claim (i.e. did they really load the record)? */
function claimValid_(fc, username, token) {
  var claim = getClaim_(fc);
  if (!claim) return { ok: false, reason: 'none' };
  if (claim.username !== username) return { ok: false, reason: 'other', who: claim.username };
  if (!token || claim.hash !== tokenHash_(String(token))) return { ok: false, reason: 'token' };
  return { ok: true, claim: claim };
}

/* ============================================================================
   FIELD-LEVEL DIFF - "what did that user actually change?"
   ============================================================================ */
var DIFF_SKIP_KEYS = { v: 1, fc: 1, farmer: 1, updatedAt: 1, updatedBy: 1, revision: 1 };

function flattenRecord_(obj, prefix, out, depth) {
  depth = depth || 0;
  if (obj === null || obj === undefined || depth > 12) return out;
  if (typeof obj !== 'object') {
    var val = String(obj);
    if (val.trim() !== '') out[prefix] = val;
    return out;
  }
  if (Array.isArray(obj)) {
    for (var i = 0; i < obj.length; i++) flattenRecord_(obj[i], prefix + '[' + (i + 1) + ']', out, depth + 1);
    return out;
  }
  var keys = Object.keys(obj);
  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    if (depth === 0 && DIFF_SKIP_KEYS[key]) continue;
    flattenRecord_(obj[key], prefix ? prefix + '.' + key : key, out, depth + 1);
  }
  return out;
}

function shortVal_(value, max) {
  var s = String(value === undefined || value === null ? '' : value).replace(/\s+/g, ' ').trim();
  max = max || 48;
  return s.length > max ? s.slice(0, max) + '…' : s;
}

function diffRecords_(oldRecord, newRecord) {
  var a = flattenRecord_(oldRecord || {}, '', {}, 0);
  var b = flattenRecord_(newRecord || {}, '', {}, 0);
  var seen = {};
  Object.keys(a).forEach(function (k) { seen[k] = 1; });
  Object.keys(b).forEach(function (k) { seen[k] = 1; });
  var changes = [];
  Object.keys(seen).sort().forEach(function (k) {
    var before = a[k] === undefined ? '' : a[k];
    var after = b[k] === undefined ? '' : b[k];
    if (String(before).trim() === String(after).trim()) return;
    changes.push({ field: k, from: before, to: after });
  });
  return changes;
}

function describeChanges_(changes, limit) {
  limit = limit || 60;
  var lines = [];
  var n = Math.min(changes.length, limit);
  for (var i = 0; i < n; i++) {
    lines.push(changes[i].field + ': "' + shortVal_(changes[i].from) + '" -> "' + shortVal_(changes[i].to) + '"');
  }
  if (changes.length > limit) lines.push('… and ' + (changes.length - limit) + ' more change(s)');
  return lines.join('\n');
}

function countFilledFields_(record) {
  return Object.keys(flattenRecord_(record || {}, '', {}, 0)).length;
}

function snapshotOf_(record) {
  if (!record) return '';
  var json;
  try { json = JSON.stringify(record); } catch (e) { return ''; }
  return json.length > MAX_SNAPSHOT_CHARS ? json.slice(0, MAX_SNAPSHOT_CHARS) + '…[truncated]' : json;
}

function getSession_(token) {
  token = String(token || '');
  if (!token) throw new Error('Please log in.');
  var hash = tokenHash_(token);
  var sheet = ensureSessions_();
  var last = sheet.getLastRow();
  if (last < 2) throw new Error('Invalid session. Please log in again.');
  var rows = sheet.getRange(2, 1, last - 1, 5).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]) !== hash) continue;
    if (new Date(rows[i][3]).getTime() <= Date.now()) throw new Error('Your session expired. Please log in again.');
    return { username: String(rows[i][1]), role: normRole_(rows[i][2]) };
  }
  throw new Error('Invalid session. Please log in again.');
}

function requireRole_(token, roles) {
  var session = getSession_(token);
  var allowed = false;
  for (var i = 0; i < roles.length; i++) {
    if (roles[i] === session.role) { allowed = true; break; }
  }
  if (!allowed) throw new Error('Your account does not have permission for this action.');
  return session;
}

function apiLogout(token) {
  try {
    var hash = tokenHash_(String(token || ''));
    var sheet = ensureSessions_();
    var last = sheet.getLastRow();
    if (last >= 2) {
      var rows = sheet.getRange(2, 1, last - 1, 5).getValues();
      for (var i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][0]) === hash) {
          sheet.deleteRow(i + 2);
          log_(String(rows[i][1]), 'LOGOUT', '', 'Session closed');
        }
      }
    }
  } catch (e) { /* logging out must never fail loudly */ }
  return { ok: true };
}

/* --- self-service password change ------------------------------------------ */
function apiChangePassword(token, currentPassword, newPassword) {
  var session = getSession_(token);
  if (isEmergencyAdmin_(session.username)) {
    throw new Error('The emergency administrator password is managed in Code.gs and cannot be changed from this screen.');
  }
  var pass = String(newPassword === null || newPassword === undefined ? '' : newPassword);
  var current = String(currentPassword === null || currentPassword === undefined ? '' : currentPassword);
  if (String(pass).length < MIN_PASSWORD_LENGTH) throw new Error('The new password must be at least ' + MIN_PASSWORD_LENGTH + ' characters.');

  var table = usersTable_();
  var row = findUserRow_(table, session.username);
  if (row < 0) throw new Error('Your account was not found in the Users sheet.');
  var user = readUser_(table, row);
  if (!passwordMatches_(user, current)) throw new Error('Your current password is not correct.');
  if (current === pass) throw new Error('The new password must be different from the current one.');

  setCell_(table, row, 'Password Hash', hashPassword_(pass));
  setCell_(table, row, 'Must Change Password', false);
  setCell_(table, row, 'Failed Attempts', 0);
  setCell_(table, row, 'Locked Until', '');
  setCell_(table, row, 'Updated At', new Date().toISOString());
  saveUsersTable_(table);
  log_(session.username, 'CHANGE_PASSWORD', '', 'Password changed by the user');
  return { ok: true };
}

/* --- admin tools ----------------------------------------------------------- */
function apiResetPassword(token, username) {
  var session = requireRole_(token, ['Admin']);
  var table = usersTable_();
  var row = findUserRow_(table, username);
  if (row < 0) throw new Error('User "' + username + '" was not found.');
  var user = readUser_(table, row);

  var temp = randomTempPassword_();
  setCell_(table, row, 'Username', user.username || normUsername_(username));
  setCell_(table, row, 'Password Hash', hashPassword_(temp));
  setCell_(table, row, 'Must Change Password', true);
  setCell_(table, row, 'Status', 'Active');
  setCell_(table, row, 'Failed Attempts', 0);
  setCell_(table, row, 'Locked Until', '');
  setCell_(table, row, 'Updated At', new Date().toISOString());
  saveUsersTable_(table);
  rememberTempPassword_(user.username, temp, user.fullname, user.role, user.email);
  log_(session.username, 'RESET_PASSWORD', user.username, 'Temporary password issued');

  return { ok: true, username: user.username, tempPassword: temp, mustChangePassword: true };
}

function apiSetUserStatus(token, username, status) {
  var session = requireRole_(token, ['Admin']);
  var table = usersTable_();
  var row = findUserRow_(table, username);
  if (row < 0) throw new Error('User "' + username + '" was not found.');
  var user = readUser_(table, row);
  var value = normStatus_(status);
  setCell_(table, row, 'Status', value);
  setCell_(table, row, 'Failed Attempts', 0);
  setCell_(table, row, 'Locked Until', '');
  setCell_(table, row, 'Updated At', new Date().toISOString());
  saveUsersTable_(table);
  log_(session.username, 'SET_USER_STATUS', user.username, value);
  return { ok: true, username: user.username, status: value };
}

function apiUnlockUser(token, username) {
  var session = requireRole_(token, ['Admin']);
  var table = usersTable_();
  var row = findUserRow_(table, username);
  if (row < 0) throw new Error('User "' + username + '" was not found.');
  var user = readUser_(table, row);
  setCell_(table, row, 'Failed Attempts', 0);
  setCell_(table, row, 'Locked Until', '');
  setCell_(table, row, 'Updated At', new Date().toISOString());
  saveUsersTable_(table);
  log_(session.username, 'UNLOCK_USER', user.username, 'Account unlocked by admin');
  return { ok: true, username: user.username };
}

/* Activate / repair every login from the web app (admin only) */
function apiActivateAllUsers(token) {
  var session = requireRole_(token, ['Admin']);
  var res = activateAllUsers_({ issuedBy: session.username });
  return { ok: true, count: res.report.length, withTempPassword: res.tempCount, report: res.report };
}

function apiGetCredentialReport(token) {
  requireRole_(token, ['Admin']);
  var s = SS.getSheetByName(CRED_SHEET);
  if (!s || s.getLastRow() < 2) return [];
  var values = s.getDataRange().getValues();
  var headers = values.shift();
  return values.map(function (row) {
    var o = {};
    for (var i = 0; i < headers.length; i++) o[String(headers[i])] = row[i];
    return o;
  });
}

/* ============================================================================
   USERS ADMIN API (used by the User Access screen)
   ============================================================================ */
function apiGetUsers(token) {
  requireRole_(token, ['Admin']);
  var table = usersTable_();
  var out = [];
  for (var r = 1; r < table.values.length; r++) {
    if (isBlankRow_(table.values[r])) continue;
    var u = readUser_(table, r);
    if (!u.username) continue;
    out.push({
      username: u.username,
      fullname: u.fullname,
      role: u.role,
      status: u.status,
      email: u.email,
      hasPassword: !!u.passwordHash,
      mustChangePassword: !!u.mustChange,
      lastLogin: u.lastLogin ? u.lastLogin.toISOString() : '',
      locked: !!(u.lockedUntil && u.lockedUntil.getTime() > Date.now()),
      createdAt: String(cellOf_(table, r, 'Created At') === null ? '' : cellOf_(table, r, 'Created At') || ''),
      updatedAt: String(cellOf_(table, r, 'Updated At') || '')
    });
  }
  return out;
}

function apiSaveUser(token, user) {
  var session = requireRole_(token, ['Admin']);
  if (!user || !user.username) throw new Error('Username is required.');
  var username = normUsername_(user.username);
  if (!username) throw new Error('Username may only contain letters, numbers, dots, dashes and underscores.');
  if (isEmergencyAdmin_(username)) throw new Error('Username "a" is reserved for the emergency administrator.');

  var table = usersTable_();
  var row = findUserRow_(table, username);
  var now = new Date().toISOString();
  var tempPassword = '';

  var fullname = String(user.fullname || '').trim();
  var role = normRole_(user.role);
  var status = normStatus_(user.status);
  var email = normEmail_(user.email);
  var password = String(user.password || '');

  if (row > 0) {
    var existing = readUser_(table, row);
    setCell_(table, row, 'Username', existing.username || username);
    setCell_(table, row, 'Full Name', fullname || existing.fullname);
    setCell_(table, row, 'Role', role);
    setCell_(table, row, 'Status', status);
    if (email) setCell_(table, row, 'Email', email);
    if (password) {
      if (password.length < MIN_PASSWORD_LENGTH) throw new Error('The password must be at least ' + MIN_PASSWORD_LENGTH + ' characters.');
      setCell_(table, row, 'Password Hash', hashPassword_(password));
      setCell_(table, row, 'Must Change Password', false);
      setCell_(table, row, 'Failed Attempts', 0);
      setCell_(table, row, 'Locked Until', '');
    }
    setCell_(table, row, 'Updated At', now);
    saveUsersTable_(table);
    log_(session.username, 'UPDATE_USER', username, role + ' / ' + status);
    return { ok: true, username: username, tempPassword: '' };
  }

  /* new account: if no password was typed, issue a temporary one */
  var rowValues = [];
  for (var w = 0; w < table.width; w++) rowValues.push('');
  if (!password) {
    tempPassword = randomTempPassword_();
    password = tempPassword;
  } else if (password.length < MIN_PASSWORD_LENGTH) {
    throw new Error('The password must be at least ' + MIN_PASSWORD_LENGTH + ' characters.');
  }

  setRowValue_(table, rowValues, 'Username', username);
  setRowValue_(table, rowValues, 'Password Hash', hashPassword_(password));
  setRowValue_(table, rowValues, 'Full Name', fullname);
  setRowValue_(table, rowValues, 'Role', role);
  setRowValue_(table, rowValues, 'Status', status);
  setRowValue_(table, rowValues, 'Email', email);
  setRowValue_(table, rowValues, 'Must Change Password', !!tempPassword);
  setRowValue_(table, rowValues, 'Created At', now);
  setRowValue_(table, rowValues, 'Updated At', now);
  setRowValue_(table, rowValues, 'Failed Attempts', 0);
  table.values.push(rowValues);
  saveUsersTable_(table);

  if (tempPassword) rememberTempPassword_(username, tempPassword, fullname, role, email);
  log_(session.username, 'CREATE_USER', username, role + (tempPassword ? ' (temporary password issued)' : ''));
  return { ok: true, username: username, tempPassword: tempPassword, mustChangePassword: !!tempPassword };
}

/* ============================================================================
   RECORDS API  (unchanged behaviour)
   ============================================================================ */
function apiEnsureSchema(token, headers) {
  requireRole_(token, ['Admin', 'User', 'Viewer']);
  var s = ensureDatabase_();
  var cols = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  return { ok: true, headers: cols, count: cols.length };
}

function findFcRow_(sheet, fc) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var values = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
  var target = String(fc).trim().toUpperCase();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === target) return i + 2;
  }
  return -1;
}

function apiSaveRecord(token, record, flatObject, schemaHeaders, mode, editToken) {
  var session = requireRole_(token, ['Admin', 'User']);
  if (!record || typeof record !== 'object') throw new Error('Invalid record.');
  var fc = String(record.fc || '').trim();
  if (!fc) throw new Error('FC Code is required.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = ensureDatabase_();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var headerMap = {};
    for (var i = 0; i < headers.length; i++) headerMap[headers[i]] = i;

    var rowIndex = findFcRow_(sheet, fc);
    var isUpdate = rowIndex > 0;

    if (isUpdate && String(mode || '') === 'create') {
      throw new Error('A record with FC Code "' + fc + '" already exists. Press “Load” to open it for editing.');
    }
    if (!isUpdate && String(mode || '') === 'update') {
      throw new Error('Record "' + fc + '" no longer exists — it may have been deleted. Press “Clear” to start a new record.');
    }

    /* ---- THE RULE: a saved record may only be overwritten by a user who
       loaded it for editing (valid edit claim) ---- */
    var previousRecord = null;
    if (isUpdate) {
      var claimCheck = claimValid_(fc, session.username, editToken);
      if (!claimCheck.ok) {
        if (claimCheck.reason === 'other') {
          throw new Error('Record "' + fc + '" is open for editing by ' + claimCheck.who + ' right now. Please try again later.');
        }
        throw new Error('Record "' + fc + '" is locked. Press “Load” to open it for editing before saving changes.');
      }
      if (headerMap['FULL_RECORD_JSON'] !== undefined) {
        var prevRaw = sheet.getRange(rowIndex, headerMap['FULL_RECORD_JSON'] + 1).getValue();
        if (prevRaw) { try { previousRecord = JSON.parse(prevRaw); } catch (e) { previousRecord = null; } }
      }
    }

    var oldRevision = 0;
    if (rowIndex > 0 && headerMap['Revision'] !== undefined) {
      oldRevision = Number(sheet.getRange(rowIndex, headerMap['Revision'] + 1).getValue()) || 0;
    }

    var revision = oldRevision + 1;
    var now = new Date().toISOString();
    record.fc = fc;
    record.updatedAt = now;
    record.updatedBy = session.username;
    record.revision = revision;

    var farmer = record.farmer || '';
    if (!farmer && record.s01 && record.s01.profile && record.s01.profile.farmer) {
      farmer = record.s01.profile.farmer;
    }

    /* Start an update from the existing row so columns the current UI no longer
       writes (legacy SP3/H1/H2 fields, future columns, etc.) are preserved.
       Current UI columns are then overwritten below, including intentional blanks. */
    var row = rowIndex > 0
      ? sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0]
      : [];
    if (rowIndex <= 0) {
      for (var r = 0; r < headers.length; r++) row.push('');
    }

    function put(name, value) {
      if (headerMap[name] !== undefined) {
        row[headerMap[name]] = (value === undefined || value === null) ? '' : value;
      }
    }

    put('FC Code', fc);
    put('Farmer Name', farmer);
    put('Last Updated', now);
    put('Updated By', session.username);
    put('Revision', revision);

    if (flatObject && typeof flatObject === 'object') {
      var keys = Object.keys(flatObject);
      for (var k = 0; k < keys.length; k++) put(keys[k], flatObject[keys[k]]);
    }
    /* Form 01 also emits a "Farmer Name" column, so the flat object can
       overwrite the value above with an empty one - fall back to the record's
       farmer name whenever that cell ended up blank. */
    if (headerMap['Farmer Name'] !== undefined && farmer &&
        String(row[headerMap['Farmer Name']] || '').trim() === '') {
      row[headerMap['Farmer Name']] = farmer;
    }

    put('FULL_RECORD_JSON', JSON.stringify(record));

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
      var changes = diffRecords_(previousRecord, record);
      var readable = describeChanges_(changes);
      logEx_(session.username, 'UPDATE_RECORD', fc,
             'Revision ' + revision + ' · ' + changes.length + ' field(s) changed',
             {
               'Changed Fields': readable,
               'Field Count': changes.length,
               'Revision': revision,
               'Snapshot': snapshotOf_(previousRecord)
             });
      SpreadsheetApp.flush();
      /* the user keeps the lock so they can carry on editing in this session */
      var keepToken = issueClaim_(fc, session.username);
      return {
        ok: true, record: record, revision: revision,
        updated: true, changedCount: changes.length,
        changedFields: describeChanges_(changes, 15),
        editToken: keepToken
      };
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
    removeClaim_(fc);
    logEx_(session.username, 'CREATE_RECORD', fc,
           'New record created · Revision 1 · ' + countFilledFields_(record) + ' field(s) filled',
           { 'Changed Fields': '', 'Field Count': 0, 'Revision': 1, 'Snapshot': '' });
    SpreadsheetApp.flush();
    return {
      ok: true, record: record, revision: revision,
      created: true, changedCount: 0, changedFields: '', editToken: ''
    };
  } finally {
    lock.releaseLock();
  }
}

/* Return a row as a header->display-value object for legacy rows whose
   FULL_RECORD_JSON cell is blank or invalid. The browser rebuilds the nested
   form record from the same module flatten definitions, then a normal Save
   writes a fresh FULL_RECORD_JSON snapshot. */
function flatRowObject_(sheet, row, headers) {
  var values = sheet.getRange(row, 1, 1, headers.length).getDisplayValues()[0];
  var out = {};
  for (var i = 0; i < headers.length; i++) {
    var name = String(headers[i] === undefined || headers[i] === null ? '' : headers[i]).trim();
    if (!name) continue;
    var value = values[i] === undefined || values[i] === null ? '' : values[i];
    /* If duplicate headers exist, prefer the nonblank copy. */
    if (out[name] === undefined || String(value).trim() !== '') out[name] = value;
  }
  return out;
}

function apiGetRecord(token, fc) {
  var session = requireRole_(token, ['Admin', 'User', 'Viewer']);
  fc = String(fc || '').trim();
  if (!fc) throw new Error('FC Code is required.');

  var sheet = ensureDatabase_();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var jsonCol = -1, revCol = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]) === 'FULL_RECORD_JSON') jsonCol = i;
    if (String(headers[i]) === 'Revision') revCol = i;
  }
  if (jsonCol < 0) throw new Error('FULL_RECORD_JSON column is missing.');

  var row = findFcRow_(sheet, fc);
  if (row < 0) return { ok: true, found: false };

  var raw = sheet.getRange(row, jsonCol + 1).getValue();
  var record = null, legacyData = false, legacyFlat = null, migratedLp = null;
  if (raw) {
    try {
      record = JSON.parse(raw);
      if (!record || typeof record !== 'object' || Array.isArray(record)) record = null;
      if (record && record.__legacyLpLabor) {
        migratedLp = record.__legacyLpLabor;
        record = null;
      }
    } catch (e) { record = null; }
  }
  if (!record) {
    legacyData = true;
    legacyFlat = flatRowObject_(sheet, row, headers);
    if (migratedLp) {
      Object.keys(migratedLp).forEach(function (name) { legacyFlat[name] = migratedLp[name]; });
    }
  }
  var revision = revCol >= 0 ? (Number(sheet.getRange(row, revCol + 1).getValue()) || 0) : 0;

  /* Viewer accounts may read but never edit, so they never take the lock */
  if (session.role === 'Viewer') {
    log_(session.username, 'VIEW_RECORD', fc, legacyData
      ? (migratedLp ? 'Read a migrated LP compatibility row (browser reconstruction)' : 'Read a legacy row without FULL_RECORD_JSON (browser reconstruction)')
      : 'Read the record (viewer, no edit lock)');
    return {
      ok: true, found: true, record: record, legacyData: legacyData,
      legacyFlat: legacyFlat, editToken: '', lockedBy: '', revision: revision
    };
  }

  var existingClaim = getClaim_(fc);
  var lockedBy = (existingClaim && existingClaim.username !== session.username) ? existingClaim.username : '';
  var editToken = issueClaim_(fc, session.username);
  logEx_(session.username, 'OPEN_FOR_EDIT', fc,
         lockedBy ? ('Opened the record for editing (was open by ' + lockedBy + ')') : 'Opened the record for editing',
         { 'Revision': revision, 'Field Count': 0, 'Changed Fields': '' });

  return {
    ok: true, found: true, record: record, legacyData: legacyData,
    legacyFlat: legacyFlat,
    editToken: editToken, lockedBy: lockedBy, lockHours: CLAIM_HOURS, revision: revision
  };
}

/* Give the lock back (Clear / Logout / switching to another record). */
function apiReleaseEdit(token, fc) {
  try {
    var session = getSession_(token);
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var claim = getClaim_(fc);
      if (claim && claim.username === session.username) removeClaim_(fc);
    } finally { lock.releaseLock(); }
  } catch (e) { /* releasing must never break the UI */ }
  return { ok: true };
}

function apiGetAll(token) {
  requireRole_(token, ['Admin', 'User', 'Viewer']);
  var sheet = ensureDatabase_();
  if (sheet.getLastRow() < 2) return { ok: true, records: [], count: 0 };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  function idx(name) {
    for (var i = 0; i < headers.length; i++) if (String(headers[i]) === name) return i;
    return -1;
  }

  var iFc = idx('FC Code');
  var iFarmer = idx('Farmer Name');
  var iUpdated = idx('Last Updated');
  var iBy = idx('Updated By');
  var iRev = idx('Revision');
  var iJson = idx('FULL_RECORD_JSON');

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var records = [];
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var fc = iFc >= 0 ? String(row[iFc] || '') : '';
    if (!fc) continue;

    var recObj = null, legacyData = false, legacyFlat = null;
    if (iJson >= 0 && row[iJson]) {
      try {
        recObj = JSON.parse(row[iJson]);
        if (recObj && recObj.__legacyLpLabor) {
          legacyData = true;
          legacyFlat = flatRowObject_(sheet, r + 2, headers);
          Object.keys(recObj.__legacyLpLabor).forEach(function (name) { legacyFlat[name] = recObj.__legacyLpLabor[name]; });
          recObj = null;
        }
      } catch (e) { recObj = null; }
    }
    if (!recObj) {
      legacyData = true;
      legacyFlat = flatRowObject_(sheet, r + 2, headers);
    }

    records.push({
      fc: fc,
      farmer: iFarmer >= 0 ? String(row[iFarmer] || '') : '',
      updatedAt: iUpdated >= 0 ? String(row[iUpdated] || '') : '',
      updatedBy: iBy >= 0 ? String(row[iBy] || '') : '',
      revision: iRev >= 0 ? Number(row[iRev] || 0) : 0,
      record: recObj,
      legacyData: legacyData,
      legacyFlat: legacyFlat
    });
  }
  return { ok: true, records: records, count: records.length };
}

function apiDeleteRecord(token, fc) {
  var session = requireRole_(token, ['Admin']);
  fc = String(fc || '').trim();
  if (!fc) throw new Error('FC Code is required.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = ensureDatabase_();
    var row = findFcRow_(sheet, fc);
    if (row < 0) return { ok: false, message: 'Record was not found.' };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var jsonCol = -1, revCol = -1;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]) === 'FULL_RECORD_JSON') jsonCol = i;
      if (String(headers[i]) === 'Revision') revCol = i;
    }
    var deleted = null;
    if (jsonCol >= 0) {
      var raw = sheet.getRange(row, jsonCol + 1).getValue();
      if (raw) { try { deleted = JSON.parse(raw); } catch (e) { deleted = null; } }
    }
    var revision = revCol >= 0 ? (Number(sheet.getRange(row, revCol + 1).getValue()) || 0) : 0;

    sheet.deleteRow(row);
    removeClaim_(fc);
    logEx_(session.username, 'DELETE_RECORD', fc,
           'Permanent deletion · ' + countFilledFields_(deleted) + ' field(s) lost (full copy kept in the Snapshot column)',
           { 'Changed Fields': '', 'Field Count': 0, 'Revision': revision, 'Snapshot': snapshotOf_(deleted) });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================================
   AUDIT TRAIL API - "who did what, when"
   ============================================================================ */
function apiGetAudit(token, fc, limit) {
  requireRole_(token, ['Admin', 'User', 'Viewer']);
  var sheet = ensureLogs_();
  var last = sheet.getLastRow();
  if (last < 2) return [];

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idx = {};
  for (var h = 0; h < headers.length; h++) idx[String(headers[h]).trim()] = h;
  function cell(row, name) { return idx[name] === undefined ? '' : row[idx[name]]; }

  var values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  var want = String(fc || '').trim().toUpperCase();
  var lim = Math.min(Math.max(Number(limit) || 100, 1), 500);
  var out = [];
  for (var i = values.length - 1; i >= 0 && out.length < lim; i--) {
    var row = values[i];
    var rowFc = String(cell(row, 'FC Code') || '');
    if (want && rowFc.trim().toUpperCase() !== want) continue;
    out.push({
      timestamp: String(cell(row, 'Timestamp') || ''),
      username: String(cell(row, 'Username') || ''),
      action: String(cell(row, 'Action') || ''),
      fc: rowFc,
      details: String(cell(row, 'Details') || ''),
      changes: String(cell(row, 'Changed Fields') || ''),
      fieldCount: Number(cell(row, 'Field Count')) || 0,
      revision: Number(cell(row, 'Revision')) || 0
    });
  }
  return out;
}

/* ============================================================================
   OPTIONS API
   ============================================================================ */
function apiGetOptions(token) {
  requireRole_(token, ['Admin', 'User', 'Viewer']);
  var rows = ensureOptions_().getDataRange().getValues();
  var result = {};
  for (var i = 1; i < rows.length; i++) {
    var group = String(rows[i][0] || '');
    var value = String(rows[i][1] || '');
    if (!group || !value) continue;
    if (!result[group]) result[group] = [];
    var exists = false;
    for (var j = 0; j < result[group].length; j++) if (result[group][j] === value) { exists = true; break; }
    if (!exists) result[group].push(value);
  }
  return result;
}

function apiAddOption(token, group, value) {
  var session = requireRole_(token, ['Admin', 'User']);
  group = String(group || '').trim();
  value = String(value || '').trim();
  if (!group || !value) throw new Error('Option group and value are required.');

  var s = ensureOptions_();
  var rows = s.getDataRange().getValues();
  var exists = false;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === group && String(rows[i][1]).toLowerCase() === value.toLowerCase()) {
      exists = true; break;
    }
  }
  if (!exists) s.appendRow([group, value, session.username, new Date().toISOString()]);
  return { ok: true };
}

/* ============================================================================
   LOGGING
   ============================================================================ */
function log_(username, action, fc, details) {
  logEx_(username, action, fc, details, null);
}

/* extra = { 'Changed Fields': '...', 'Field Count': 7, 'Revision': 3, 'Snapshot': '{...}' } */
function logEx_(username, action, fc, details, extra) {
  var sheet = ensureLogs_();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  for (var i = 0; i < headers.length; i++) row.push('');
  var map = {};
  for (var h = 0; h < headers.length; h++) map[String(headers[h]).trim()] = h;
  function put(name, value) {
    if (map[name] !== undefined) row[map[name]] = (value === undefined || value === null) ? '' : value;
  }
  put('Timestamp', new Date().toISOString());
  put('Username', username || '');
  put('Action', action || '');
  put('FC Code', fc || '');
  put('Details', details || '');
  if (extra) {
    var keys = Object.keys(extra);
    for (var k = 0; k < keys.length; k++) put(keys[k], extra[keys[k]]);
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
}

/* ============================================================================
   SPREADSHEET MENU ACTIONS  &  one-click setup
   ============================================================================ */

/* Run this ONCE from the Apps Script editor (or use the menu). */
function setupIFAR() {
  initializeIFAR();
  var res = activateAllUsers_({ issuedBy: 'owner' });
  return res;
}

/* Issues/repairs the login credentials of every row in the Users sheet. */
function activateAllUsers_(options) {
  options = options || {};
  var sheet = ensureUsers_();                                   // guarantees headers + defaults
  var res = repairUsers_(sheet, { issueTempPasswords: true, quiet: false });
  var withTemp = 0;
  for (var i = 0; i < res.report.length; i++) if (res.report[i].tempPassword) withTemp++;
  if (res.report.length) {
    log_(options.issuedBy || 'system', 'ACTIVATE_USERS', '', res.report.length + ' account(s) verified, ' + withTemp + ' temporary password(s) issued');
  }
  return { ok: true, report: res.report, tempCount: withTemp, sheet: sheet };
}

function credentialReportText_(report) {
  var lines = ['IFAR DIGITAL · LOGIN CREDENTIALS', new Date().toLocaleString(), ''];
  for (var i = 0; i < report.length; i++) {
    var r = report[i];
    lines.push((i + 1) + '. ' + (r.fullName || r.username) + '  [' + r.role + ' / ' + r.status + ']');
    lines.push('   username : ' + r.username);
    if (r.email) lines.push('   e-mail   : ' + r.email);
    lines.push('   password : ' + (r.tempPassword || '(kept as set before)'));
    if (r.note) lines.push('   note     : ' + r.note);
    lines.push('');
  }
  lines.push('Note: temporary passwords must be changed at first login.');
  return lines.join('\n');
}

function menuActivateAll() {
  var res = activateAllUsers_({ issuedBy: 'owner' });
  setReportVisibility_(true);
  var tempCount = 0;
  for (var i = 0; i < res.report.length; i++) if (res.report[i].tempPassword) tempCount++;
  SpreadsheetApp.getUi().alert(
    'User logins activated',
    res.report.length + ' account(s) checked.\n' + tempCount + ' temporary password(s) issued.\n\n' +
    'The full list is in the "' + CRED_SHEET + '" sheet (now visible).\n' +
    'Send each person their username + temporary password, then hide or clear the sheet.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function menuSeedDefaults() {
  var table = usersTable_();
  var existing = [];
  for (var r = 1; r < table.values.length; r++) {
    var u = readUser_(table, r);
    if (u.username) existing.push(u.username);
  }
  var added = [];
  var now = new Date().toISOString();
  for (var i = 0; i < DEFAULT_USERS.length; i++) {
    var d = DEFAULT_USERS[i];
    if (existing.indexOf(d[0]) >= 0) continue;
    var rowValues = [];
    for (var w = 0; w < table.width; w++) rowValues.push('');
    setRowValue_(table, rowValues, 'Username', d[0]);
    setRowValue_(table, rowValues, 'Password Hash', hashPassword_(d[1]));
    setRowValue_(table, rowValues, 'Full Name', d[2]);
    setRowValue_(table, rowValues, 'Role', d[3]);
    setRowValue_(table, rowValues, 'Status', d[4]);
    setRowValue_(table, rowValues, 'Created At', now);
    setRowValue_(table, rowValues, 'Updated At', now);
    setRowValue_(table, rowValues, 'Must Change Password', false);
    table.values.push(rowValues);
    added.push(d[0] + ' / ' + d[1]);
  }
  if (added.length) saveUsersTable_(table);
  SpreadsheetApp.getUi().alert('Default accounts',
    added.length ? ('Created:\n' + added.join('\n')) : 'All default accounts already exist.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuAddUser() {
  var ui = SpreadsheetApp.getUi();
  var fullName = ui.prompt('New user · Full name', 'e.g. Juan Dela Cruz', ui.ButtonSet.OK_CANCEL);
  if (fullName.getSelectedButton() !== ui.Button.OK) return;
  var username = ui.prompt('New user · Username', 'Leave blank to build it from the name', ui.ButtonSet.OK_CANCEL);
  if (username.getSelectedButton() !== ui.Button.OK) return;
  var roleAns = ui.prompt('New user · Role', 'Admin / User / Viewer', ui.ButtonSet.OK_CANCEL);
  if (roleAns.getSelectedButton() !== ui.Button.OK) return;
  var emailAns = ui.prompt('New user · E-mail (optional)', 'name@example.com', ui.ButtonSet.OK_CANCEL);
  if (emailAns.getSelectedButton() !== ui.Button.OK) return;

  var name = fullName.getResponseText().trim();
  var temp = randomTempPassword_();
  var table = usersTable_();
  var rowValues = [];
  for (var w = 0; w < table.width; w++) rowValues.push('');
  var uname = normUsername_(username.getResponseText()) || usernameFromName_(name) || 'user';
  var used = {};
  for (var r = 1; r < table.values.length; r++) {
    var u = readUser_(table, r);
    if (u.username) used[u.username] = true;
  }
  var base = uname, n = 2;
  while (used[uname]) { uname = base + n; n++; }

  var now = new Date().toISOString();
  setRowValue_(table, rowValues, 'Username', uname);
  setRowValue_(table, rowValues, 'Password Hash', hashPassword_(temp));
  setRowValue_(table, rowValues, 'Full Name', name);
  setRowValue_(table, rowValues, 'Role', normRole_(roleAns.getResponseText()));
  setRowValue_(table, rowValues, 'Status', 'Active');
  setRowValue_(table, rowValues, 'Email', normEmail_(emailAns.getResponseText()));
  setRowValue_(table, rowValues, 'Must Change Password', true);
  setRowValue_(table, rowValues, 'Created At', now);
  setRowValue_(table, rowValues, 'Updated At', now);
  setRowValue_(table, rowValues, 'Failed Attempts', 0);
  table.values.push(rowValues);
  saveUsersTable_(table);
  rememberTempPassword_(uname, temp, name, normRole_(roleAns.getResponseText()), normEmail_(emailAns.getResponseText()));
  log_('owner', 'CREATE_USER', uname, 'Created from the spreadsheet menu');

  ui.alert('User created',
    'Username : ' + uname + '\nTemporary password : ' + temp + '\n\n' +
    'The user must set a personal password at first login.',
    ui.ButtonSet.OK);
}

function menuResetPassword() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.prompt('Reset password', 'Username to reset:', ui.ButtonSet.OK_CANCEL);
  if (ans.getSelectedButton() !== ui.Button.OK) return;
  var table = usersTable_();
  var row = findUserRow_(table, ans.getResponseText());
  if (row < 0) { ui.alert('User not found.'); return; }
  var user = readUser_(table, row);
  var temp = randomTempPassword_();
  setCell_(table, row, 'Password Hash', hashPassword_(temp));
  setCell_(table, row, 'Must Change Password', true);
  setCell_(table, row, 'Status', 'Active');
  setCell_(table, row, 'Failed Attempts', 0);
  setCell_(table, row, 'Locked Until', '');
  setCell_(table, row, 'Updated At', new Date().toISOString());
  saveUsersTable_(table);
  rememberTempPassword_(user.username, temp, user.fullname, user.role, user.email);
  log_('owner', 'RESET_PASSWORD', user.username, 'Temporary password issued from the menu');
  ui.alert('Password reset', 'Username : ' + user.username + '\nTemporary password : ' + temp, ui.ButtonSet.OK);
}

function menuUnlockUser() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.prompt('Unlock account', 'Username to unlock:', ui.ButtonSet.OK_CANCEL);
  if (ans.getSelectedButton() !== ui.Button.OK) return;
  var table = usersTable_();
  var row = findUserRow_(table, ans.getResponseText());
  if (row < 0) { ui.alert('User not found.'); return; }
  var user = readUser_(table, row);
  setCell_(table, row, 'Failed Attempts', 0);
  setCell_(table, row, 'Locked Until', '');
  setCell_(table, row, 'Updated At', new Date().toISOString());
  saveUsersTable_(table);
  log_('owner', 'UNLOCK_USER', user.username, 'Unlocked from the menu');
  ui.alert('"' + user.username + '" is unlocked and can log in again.');
}

function menuEmailCredentials() {
  var ui = SpreadsheetApp.getUi();
  var s = SS.getSheetByName(CRED_SHEET);
  if (!s || s.getLastRow() < 2) {
    activateAllUsers_({ issuedBy: 'owner' });            // build the list first
    s = SS.getSheetByName(CRED_SHEET);
  }
  if (!s || s.getLastRow() < 2) { ui.alert('Nothing to send — no users found.'); return; }

  var rows = s.getDataRange().getValues();
  var headers = rows.shift();
  function col(name) {
    for (var i = 0; i < headers.length; i++) if (String(headers[i]).trim() === name) return i;
    return -1;
  }
  var iU = col('Username'), iF = col('Full Name'), iR = col('Role'), iE = col('Email'), iT = col('Temporary Password');
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }

  var sent = 0, skipped = 0, failed = 0;
  for (var i = 0; i < rows.length; i++) {
    var line = rows[i];
    var username = String(line[iU] || '').trim();
    var email = normEmail_(line[iE]);
    var temp = String(line[iT] || '').replace('  (still valid)', '').trim();
    if (!username || !email || !TEMP_PASSWORD_PATTERN.test(temp)) { skipped++; continue; }
    try {
      MailApp.sendEmail({
        to: email,
        subject: 'IFAR DIGITAL · your login credentials',
        body: 'Hello ' + (String(line[iF] || '') || username) + ',\n\n' +
              'Your IFAR DIGITAL account is active.\n\n' +
              'Username: ' + username + '\n' +
              'Temporary password: ' + temp + '\n' +
              (url ? 'Open the app: ' + url + '\n' : '') +
              '\nYou will be asked to choose your own password the first time you sign in.\n\n' +
              'IFAR DIGITAL · Central Cloud Database'
      });
      sent++;
    } catch (e) { failed++; }
  }
  ui.alert('Credential e-mails',
    'Sent: ' + sent + '\nSkipped (no e-mail address, or the user already has a personal password): ' + skipped + '\nFailed: ' + failed,
    ui.ButtonSet.OK);
}

function menuShowReport() {
  var s = SS.getSheetByName(CRED_SHEET);
  if (!s) { SpreadsheetApp.getUi().alert('Run "Activate / repair ALL user logins" first.'); return; }
  setReportVisibility_(true);
  SS.setActiveSheet(s);
}

function menuClearReport() {
  var s = SS.getSheetByName(CRED_SHEET);
  if (s) s.clear();
  SpreadsheetApp.getUi().alert('Credential report cleared.');
}
