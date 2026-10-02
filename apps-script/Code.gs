// ============================================================
//  WGS REFUND REQUEST SYSTEM — CONFIGURATION
//  WGS · Refund Tracker · Google Apps Script
// ============================================================

// The tracker workbook. Filled in so the script works whether it is bound to
// the spreadsheet or running standalone: getActiveSpreadsheet() returns null
// in a standalone project and from some trigger contexts.
const TRACKER_ID = 'YOUR_SPREADSHEET_ID';

var _ss = null;

/**
 * Resolves the spreadsheet lazily. Tries the bound workbook first, then
 * falls back to opening by id. Resolving at call time rather than at file
 * load means a null here can no longer break every other function.
 */
function ss_() {
  if (_ss) return _ss;
  try { _ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { _ss = null; }
  if (!_ss && TRACKER_ID) {
    try { _ss = SpreadsheetApp.openById(TRACKER_ID); } catch (e) { _ss = null; }
  }
  if (!_ss) {
    throw new Error('Could not open the tracker workbook. Set TRACKER_ID at the '
                  + 'top of Code.gs to the spreadsheet id.');
  }
  return _ss;
}

/**
 * Non-fatal toast. Never throws, so a failed notification cannot abort the
 * work that triggered it.
 */
function toast_(message, title, seconds) {
  try { ss_().toast(message, title || 'Refund Tracker', seconds || 6); }
  catch (e) { Logger.log('[toast] ' + title + ': ' + message); }
}

/**
 * Returns a sheet by name, or null if it does not exist.
 * Every accessor below is null-safe: a missing optional sheet
 * degrades gracefully instead of throwing.
 */
function sheet_(name) {
  try { return ss_().getSheetByName(name); } catch (e) { return null; }
}

// ── Context-safe UI ──────────────────────────────────────────
// SpreadsheetApp.getUi() throws "Cannot call ... from this context" whenever
// there is no attached spreadsheet UI: a standalone project, a time-driven
// or installable trigger, or a web app request. Every dialog in this file
// therefore goes through these two helpers, which fall back to the execution
// log instead of aborting the function that was doing the real work.

function ui_() {
  try { return SpreadsheetApp.getUi(); } catch (e) { return null; }
}

/** Shows an alert when a UI exists; always writes to the log. */
function notify_(title, message) {
  Logger.log('[' + title + '] ' + message);
  var ui = ui_();
  if (!ui) return false;
  try { ui.alert(title, message, ui.ButtonSet.OK); return true; }
  catch (e) { return false; }
}

/**
 * Asks a yes/no question. Returns true only on an explicit YES.
 * Without a UI it returns false and logs why, so a destructive action
 * can never proceed unconfirmed from a trigger.
 */
function confirm_(title, message) {
  var ui = ui_();
  if (!ui) {
    Logger.log('[' + title + '] Needs confirmation but no UI is available. '
             + 'Run this from the spreadsheet menu. Message was: ' + message);
    return false;
  }
  try { return ui.alert(title, message, ui.ButtonSet.YES_NO) === ui.Button.YES; }
  catch (e) { return false; }
}

const REQUESTS  = () => sheet_('Requests');
const REFUNDS   = () => sheet_('REFUNDS');
const AIRLINE   = () => sheet_('AIRLINE');
const STAFF     = () => sheet_('Data Validation');
const TICKETS25 = () => sheet_('2025 TICKETS');

/** Throws a clear, human-readable error for sheets that are mandatory. */
function requireSheet_(sheet, name) {
  if (!sheet) {
    throw new Error('Required sheet "' + name + '" was not found. '
                  + 'Check the tab name and try again.');
  }
  return sheet;
}

const REQ_ID_PREFIX = 'RFQ';

// ── Email ────────────────────────────────────────────────────
const EMAIL_REPLY_TO = 'customersuccess@example.com';

// NOTE: this address is 'refund', NOT 'refunds'. The trailing "s" was the
// reason the refund manager never received an alert: mail was going to a
// mailbox that does not exist, and GmailApp does not report that back.
const EMAIL_TO_INTERNAL = [
  'refund@example.com',
  'customersuccess@example.com'
].join(',');

// Copied openly (CC) on the internal alert, so everyone can see who else
// is on the thread and can reply-all. These addresses are therefore visible
// to every recipient of the INTERNAL email. They are never added to the
// customer email, so no client ever sees this list.
const EMAIL_CC = [
  'hob@example.com',
  'financemanager@example.com',
  'treasury@example.com',
  'cfo@holdings.example.com',
  'headrac@holdings.example.com',
  'audit.officer@holdings.example.com',
  'team.lead@example.com',
  'hoo@example.com',
  'operationsmanager@example.com'
].join(',');

// ── Tracker link ─────────────────────────────────────────────
// Included in the internal and requester emails so the team can jump
// straight to the sheet and update the status. Deliberately NOT sent
// to customers.
const TRACKER_URL = 'https://docs.google.com/spreadsheets/d/'
                  + 'YOUR_SPREADSHEET_ID/edit'
                  + '?gid=711786809#gid=711786809';

// ── Branding ─────────────────────────────────────────────────
// WGS logo, stored on Google Drive.
// Paste just the FILE ID from the share link, e.g.
//   https://drive.google.com/file/d/<FILE_ID>/view
const LOGO_FILE_ID = 'YOUR_LOGO_FILE_ID';

// Public fallback URL, used only if the Drive fetch fails.
const LOGO_FALLBACK_URL = 'https://drive.google.com/thumbnail?id='
                        + LOGO_FILE_ID + '&sz=w320';

// Rendered width of the logo in emails and on the form header.
const LOGO_WIDTH_PX = 150;


// ── Built-in IATA code map ───────────────────────────────────
// Used when the AIRLINE sheet is absent, and as a fallback for any
// code that sheet does not cover. Unknown codes pass through as-is.
const IATA_MAP = {
  'ET':'ETHIOPIAN AIRLINES',      'QR':'QATAR AIRWAYS',
  'P4':'AIR PEACE',               'BA':'BRITISH AIRWAYS',
  'TK':'TURKISH AIRLINES',        'EK':'EMIRATES',
  'AF':'AIR FRANCE',              'VS':'VIRGIN ATLANTIC',
  'UN':'UNITED NIGERIA AIRLINES', 'WB':'RWANDAIR',
  'MS':'EGYPTAIR',                'AT':'ROYAL AIR MAROC',
  'LH':'LUFTHANSA',               'KQ':'KENYA AIRWAYS',
  'SA':'SOUTH AFRICAN AIRWAYS',   'KL':'KLM',
  'W3':'ARIK AIR',                'DL':'DELTA AIR LINES',
  'UA':'UNITED AIRLINES',         'SV':'SAUDIA',
  'DT':'TAAG ANGOLA AIRLINES',    'KP':'ASKY AIRLINES',
  'TC':'AIR TANZANIA',            'AW':'AFRICA WORLD AIRLINES',
  'TP':'TAP AIR PORTUGAL',        'AI':'AIR INDIA',
  'ME':'MIDDLE EAST AIRLINES',    'LA':'LATAM AIRLINES',
  'AA':'AMERICAN AIRLINES',       'AC':'AIR CANADA',
  'AM':'AEROMEXICO',              'VY':'VUELING',
  'AH':'AIR ALGERIE',             'W6':'WIZZ AIR',
  'EW':'EUROWINGS',               'CZ':'CHINA SOUTHERN AIRLINES',
  'GF':'GULF AIR',                'GA':'GARUDA INDONESIA',
  'HR':'HAHN AIR',                '6E':'INDIGO',
  'KE':'KOREAN AIR',              'MU':'CHINA EASTERN AIRLINES',
  'MH':'MALAYSIA AIRLINES',       'IB':'IBERIA',
  'TG':'THAI AIRWAYS',            'AV':'AVIANCA',
  'WS':'WESTJET',                 'VA':'VIRGIN AUSTRALIA',
  'D8':'NORWEGIAN AIR',           'TO':'TRANSAVIA',
  'QI':'CEMAIR',                  'HF':'AIR COTE DIVOIRE',
  'OF':'OVERLAND AIRWAYS',        'UY':'CAMAIR-CO',
  '4Z':'AIRLINK'
};

// ── Refund Types ─────────────────────────────────────────────
const REFUND_TYPES = [
  'VOLUNTARY',
  'INVOLUNTARY',
  'SCHEDULE CHANGE',
  'MEDICAL',
  'DEATH',
  'TAX REFUND',
  'WAIVER',
  'OTHER'
];

// ── Disbursement Types (where the money goes) ────────────────
// Read live from Data Validation!D2:D — this is the fallback.
const DISBURSEMENT_TYPES_FALLBACK = [
  'TRANSFER TO CLIENT LEDGER',
  'PAY TO CUSTOMER ACCOUNT',
  'DEBT PAYMENT'
];

// ── Column indices in the Requests sheet (1-based) ───────────
const COL = {
  REQ_ID:              1,   // A   S/N
  REQUEST_DATE:        2,   // B   ← stage 1 timestamp
  REQUESTED_BY:        3,   // C
  TXN_ID:              4,   // D
  TXN_DATE:            5,   // E   DATE OF INITIAL TRANSACTION
  TICKET_NUMBER:       6,   // F
  PAX_NAME:            7,   // G
  CLIENT:              8,   // H
  AIRLINE:             9,   // I
  REFUND_TYPE:        10,   // J
  ISSUED_FROM:        11,   // K   auto-filled from REFUNDS!G
  AMOUNT_ADVICE:      12,   // L   ← form fills: amount advised, in naira
  COMMENT:            13,   // M   ← form fills
  CUSTOMER_EMAIL:     14,   // N   supports multiple, comma separated
  STATUS:             15,   // O   – manager fills
  DISBURSEMENT:       16,   // P   – manager fills
  NOW_PROCESSING:     17,   // Q   ← stage 2 timestamp (auto)
  PROCESSING_DONE:    18,   // R   ← stage 3 timestamp (auto)
  SERVICE_CHARGE_DONE:19,   // S   SERVICE CHARGE COMPLETED (see note below)
  PAYOUT_COMPLETED:   20,   // T   ← stage 5 timestamp (auto)
  SERVICE_CHARGE_AT:  21,   // U   ← stage 4 timestamp (auto), SERVICE CHARGE DATE
  ISSUED_DATE:        22,   // V   REFUND ISSUED DATE (auto)
  BSP_RA_NO:          23,   // W
  MANAGER_REMARK:     24,   // X   REFUND MANAGER / OPM REMARK
  CURRENCY_TYPE:      25,   // Y
  APPROVED_NGN:       26,   // Z   – manager fills
  PROFIT:             27,   // AA  – manager fills
  MODIFIED_BY:        28,   // AB  – edit log
  REQUESTER_EMAIL:    29,   // AC  – hidden, set at submission
  LAST_MODIFIED:      30,   // AD  – edit log timestamp (see note below)
  LAST_COL:           30
};

// NOTE ON TWO SERVICE-CHARGE COLUMNS
// The sheet carries both S "SERVICE CHARGE COMPLETED" (2 rows populated) and
// U "SERVICE CHARGE DATE" (17 rows). U is the one the pipeline stamps and the
// one every turnaround metric reads. S is left untouched by this script; if it
// is genuinely redundant it can be deleted, but nothing here depends on it.

// NOTE ON LAST DATE MODIFIED
// That column was removed from the sheet, leaving the edit log with an editor
// but nowhere to record when. It is re-created at AD by ensureEditLogColumn(),
// which runs from the menu and appends rather than inserting, so no existing
// column shifts.

// ── REFUNDS sheet columns (1-based) ─────────────────────────
const RCOL = {
  TXN_ID:      1,  // A
  DATE:        2,  // B
  PAX:         3,  // C
  TICKET_NO:   4,  // D
  CLIENT:      5,  // E
  AIRLINE:     6,  // F
  ISSUED_FROM: 7   // G  (present in REFUNDS only, not in the 2025 archive)
};

// ── Status pipeline ──────────────────────────────────────────
// Updated to the five-stage flow:
//   PENDING -> PROCESSING -> PROCESS DONE -> SERVICE CHARGE -> POSTED TO TRAVCO
const STATUS = {
  PENDING:        'PENDING',
  PROCESSING:     'PROCESSING',
  PROCESS_DONE:   'PROCESS DONE',
  SERVICE_CHARGE: 'SERVICE CHARGE',
  POSTED:         'POSTED TO TRAVCOM'
};

// Ordered pipeline, used to build the sheet dropdown and to sanity-check
// that a status change moves forward rather than skipping stages.
const STATUS_ORDER = [
  STATUS.PENDING,
  STATUS.PROCESSING,
  STATUS.PROCESS_DONE,
  STATUS.SERVICE_CHARGE,
  STATUS.POSTED
];

// Values still present in historical rows, normalised whenever read.
//   PROCESSING DONE   -> renamed to PROCESS DONE
//   OUT / PAY OUT ... -> the old terminal state, now DONE
// DONE itself is kept as a recognised legacy terminal state: 179 rows carry
// it, and silently rewriting them to POSTED TO TRAVCO would assert something
// about the ledger that the historical data does not actually record.
const LEGACY_STATUS_MAP = {
  'PROCESSING DONE':   'PROCESS DONE',
  'OUT':               'DONE',
  'PAY OUT COMPLETED': 'DONE',
  'POSTED TO TRAVCO':  'POSTED TO TRAVCOM'   // the old spelling, missing its trailing m
};

// Statuses that close a request. DONE is legacy; POSTED TO TRAVCO is current.
const TERMINAL_STATUSES = ['DONE', STATUS.POSTED];

function isTerminalStatus_(s) {
  return TERMINAL_STATUSES.indexOf(normaliseStatus_(s)) !== -1;
}

function normaliseStatus_(raw) {
  var s = String(raw || '').trim().toUpperCase();
  return LEGACY_STATUS_MAP[s] || s;
}


// ============================================================
//  UTILITIES
// ============================================================

/**
 * Converts any currency-ish value to a number.
 * Handles 450000, "450,000", "₦450,000.00", " 450000 ".
 * Returns null when no usable number can be extracted.
 */
function toNumber_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (typeof val === 'number') return isNaN(val) ? null : val;

  var cleaned = String(val).replace(/[^0-9.\-]/g, '');
  if (cleaned === '' || cleaned === '-' || cleaned === '.') return null;

  var num = parseFloat(cleaned);
  return isNaN(num) ? null : num;
}

/**
 * Splits a multi-email string on comma, semicolon, space, or newline,
 * validates each, de-duplicates, and returns a clean comma-joined string.
 * Returns '' when nothing valid is found.
 */
function cleanEmails_(raw) {
  if (!raw) return '';
  var parts = String(raw).split(/[,;\s\n]+/);
  var seen  = {};
  var out   = [];

  for (var i = 0; i < parts.length; i++) {
    var e = parts[i].trim().toLowerCase();
    if (!e) continue;
    // Basic shape check: something@something.something
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) continue;
    if (seen[e]) continue;
    seen[e] = true;
    out.push(e);
  }
  return out.join(',');
}

// ── Logo helpers ─────────────────────────────────────────────

/**
 * Fetches the logo from Drive as a Blob for inline email embedding.
 *
 * Inline embedding is used rather than a hotlinked <img src="https://...">
 * because Gmail and Outlook routinely block images served from Drive.
 * The script runs as an account that can already read the file, so this
 * works whether or not the file is shared publicly.
 *
 * Returns null when the file cannot be read; callers fall back to text.
 */
function getLogoBlob_() {
  try {
    if (!LOGO_FILE_ID) return null;
    var file = DriveApp.getFileById(LOGO_FILE_ID);
    return file.getBlob().setName('wgs-logo');
  } catch (e) {
    Logger.log('Logo fetch failed: ' + e.message);
    return null;
  }
}

/**
 * Returns the logo as a base64 data URI for the web form.
 * Cached for 6 hours so the form does not hit Drive on every load.
 * Called client-side via google.script.run.
 */
function getLogoDataUri() {
  var cache = CacheService.getScriptCache();
  var hit   = cache.get('logoDataUri');
  if (hit) return hit;

  try {
    var blob = getLogoBlob_();
    if (!blob) return '';

    var uri = 'data:' + blob.getContentType() + ';base64,'
            + Utilities.base64Encode(blob.getBytes());

    // Cache entries are capped at 100KB; skip silently if oversized.
    if (uri.length < 99000) cache.put('logoDataUri', uri, 21600);
    return uri;
  } catch (e) {
    Logger.log('Logo data URI failed: ' + e.message);
    return '';
  }
}

/** Whole days between two dates, or null. */
function daysBetween_(later, earlier) {
  if (!(later instanceof Date) || !(earlier instanceof Date)) return null;
  var ms = later.getTime() - earlier.getTime();
  if (isNaN(ms)) return null;
  return Math.round((ms / 86400000) * 10) / 10;
}


// ============================================================
//  REQUEST ID GENERATOR  —  RFQ-YYYYMM-nnnnn
// ============================================================

function generateRequestId_() {
  var now    = new Date();
  var period = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0');
  var prefix = REQ_ID_PREFIX + '-' + period + '-';

  var sheet   = requireSheet_(REQUESTS(), 'Requests');
  var lastRow = sheet.getLastRow();
  var maxSeq  = 0;

  if (lastRow >= 2) {
    var colA = sheet.getRange(2, COL.REQ_ID, lastRow - 1, 1).getValues();
    for (var i = 0; i < colA.length; i++) {
      var val = String(colA[i][0] || '');
      if (val.indexOf(prefix) === 0) {
        var seq = parseInt(val.substring(prefix.length), 10);
        if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
      }
    }
  }
  return prefix + String(maxSeq + 1).padStart(5, '0');
}


// ============================================================
//  FORM SUBMISSION HANDLER
// ============================================================

function firstEmptyRequestRow_() {
  var sheet   = requireSheet_(REQUESTS(), 'Requests');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;

  var data = sheet.getRange(2, 1, lastRow - 1, COL.LAST_COL).getValues();

  for (var i = data.length - 1; i >= 0; i--) {
    for (var c = 0; c < data[i].length; c++) {
      if (data[i][c] !== '' && data[i][c] !== null && data[i][c] !== undefined) {
        return i + 3;
      }
    }
  }
  return 2;
}

/**
 * Builds the IATA code to airline name map.
 * Starts from the built-in table, then overlays the AIRLINE sheet
 * when that sheet exists so your own entries take precedence.
 */
function airlineCodeMap_() {
  var out = {};
  for (var k in IATA_MAP) out[k] = IATA_MAP[k];

  var sheet = AIRLINE();
  if (!sheet) return out;                       // sheet is optional

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return out;

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  for (var i = 0; i < data.length; i++) {
    var code = String(data[i][0] || '').trim().toUpperCase();
    var name = String(data[i][1] || '').trim().toUpperCase();
    if (code && name) out[code] = name;
  }
  return out;
}

/**
 * Resolves an airline value to a full name.
 * Handles the messiness in the source data:
 *   - inconsistent case      ("p4", "Ba", "ms")
 *   - multi-carrier itineraries ("AF/KL", "ET, LA", "KQ/TG")
 *   - values that are already full names (passed straight through)
 * Unknown codes are returned unchanged rather than being blanked.
 */
function resolveAirline_(val, map) {
  var v = String(val || '').trim();
  if (!v) return '';

  var up = v.toUpperCase();

  // Already a full name
  if (v.length > 3 && !/[\/,]/.test(v)) return up;

  // Direct code hit
  if (map[up]) return map[up];

  // Multi-carrier: split on / or comma and resolve each part
  if (/[\/,]/.test(v)) {
    var parts = v.split(/[\/,]+/);
    var names = [];
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i].trim().toUpperCase();
      if (!p) continue;
      names.push(map[p] || p);
    }
    if (names.length) return names.join(' / ');
  }

  return up;   // unknown code, keep it visible
}

/**
 * Finds a column by header name on a sheet. Returns 0 when absent.
 *
 * Reading by position was wrong: the code took column 7 as ISSUED FROM
 * whenever a sheet had seven or more columns, but on 2025 TICKETS column 7
 * is TRAVEL CONSULTANT, so consultant names were being written into the
 * Issued From field.
 */
function headerCol_(sheet, name) {
  if (!sheet || sheet.getLastColumn() < 1) return 0;
  var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var want = String(name).trim().toUpperCase();
  for (var i = 0; i < hdr.length; i++) {
    if (String(hdr[i]).trim().toUpperCase() === want) return i + 1;
  }
  return 0;
}

/**
 * Reads one transaction source sheet into a normalised array.
 * Columns are resolved by header, so a sheet that lacks ISSUED FROM
 * simply yields an empty value instead of borrowing another column.
 */
function readTxnSheet_(sheet, sourceLabel, map) {
  var out = [];
  if (!sheet) return out;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return out;

  var cId  = headerCol_(sheet, 'TRANSACTION ID') || 1;
  var cDt  = headerCol_(sheet, 'DATE')           || 2;
  var cPax = headerCol_(sheet, 'PAX')            || 3;
  var cTkt = headerCol_(sheet, 'TICKET NUMBER')  || 4;
  var cCli = headerCol_(sheet, 'CLIENT')         || 5;
  var cAir = headerCol_(sheet, 'AIRLINE')        || 6;
  var cIss = headerCol_(sheet, 'ISSUED FROM');   // 0 when the sheet has none

  var need = Math.max(cId, cDt, cPax, cTkt, cCli, cAir, cIss);
  var d = sheet.getRange(2, 1, lastRow - 1, need).getDisplayValues();

  for (var i = 0; i < d.length; i++) {
    var txnId = String(d[i][cId - 1] || '').trim();
    if (!txnId) continue;

    var ticketNo = String(d[i][cTkt - 1] || '').trim();
    var pax      = String(d[i][cPax - 1] || '').trim();

    out.push({
      txnId:      txnId,
      date:       d[i][cDt - 1],
      pax:        pax,
      ticketNo:   ticketNo,
      client:     String(d[i][cCli - 1] || '').trim(),
      airline:    resolveAirline_(d[i][cAir - 1], map),
      issuedFrom: cIss ? String(d[i][cIss - 1] || '').trim() : '',
      source:     sourceLabel,
      key:        txnId + '||' + ticketNo + '||' + pax
    });
  }
  return out;
}


// ── Chunked cache ────────────────────────────────────────────
// The combined index is roughly 2.2 MB across ~10,700 transactions, which is
// over twenty times the 100 KB ceiling on a single CacheService entry. The
// old single put() therefore threw every time, nothing was ever cached, and
// each form load re-read both sheets in full. Splitting the payload keeps it
// cached and keeps the sheets out of the request path.
var TXN_CACHE_KEY = 'txnIndexV2';
var TXN_CACHE_TTL = 21600;          // 6 hours
var CHUNK_BYTES   = 90000;

function cacheGetChunked_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var head  = cache.get(key + '_N');
    if (!head) return null;

    var n = parseInt(head, 10), keys = [];
    for (var i = 0; i < n; i++) keys.push(key + '_' + i);

    var parts = cache.getAll(keys), out = '';
    for (var j = 0; j < n; j++) {
      var piece = parts[key + '_' + j];
      if (piece === undefined || piece === null) return null;   // a chunk expired
      out += piece;
    }
    return out;
  } catch (e) {
    Logger.log('cacheGetChunked_ failed: ' + e.message);
    return null;
  }
}

function cachePutChunked_(key, str, ttl) {
  try {
    var cache = CacheService.getScriptCache();
    var n = Math.ceil(str.length / CHUNK_BYTES);
    if (n > 250) { Logger.log('Index too large to cache (' + n + ' chunks).'); return false; }

    var payload = {};
    for (var i = 0; i < n; i++) payload[key + '_' + i] = str.substr(i * CHUNK_BYTES, CHUNK_BYTES);
    payload[key + '_N'] = String(n);
    cache.putAll(payload, ttl);
    return true;
  } catch (e) {
    Logger.log('cachePutChunked_ failed: ' + e.message);
    return false;
  }
}

/**
 * Returns every selectable transaction from BOTH the live REFUNDS feed and
 * the 2025 TICKETS archive, served from the chunked cache when warm.
 */
function allTransactions_() {
  var hit = cacheGetChunked_(TXN_CACHE_KEY);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* rebuild below */ }
  }

  var map     = airlineCodeMap_();
  var live    = readTxnSheet_(REFUNDS(),   'CURRENT', map);
  var archive = readTxnSheet_(TICKETS25(), '2025',    map);
  var all     = live.concat(archive);

  cachePutChunked_(TXN_CACHE_KEY, JSON.stringify(all), TXN_CACHE_TTL);
  return all;
}

/**
 * Finds a transaction by ID, disambiguated by ticket number and passenger.
 * Searches the live REFUNDS feed first, then the 2025 archive.
 *
 * Disambiguation matters because the archive contains both group bookings
 * (one PNR, many passengers) and category rows such as VISA or HOTEL where
 * the ID repeats across dozens of unrelated customers.
 */
function lookupRefundData_(txnId, ticketNo, pax) {
  var all = allTransactions_();
  var wantId     = String(txnId || '').trim().toUpperCase();
  var wantTicket = String(ticketNo || '').trim();
  var wantPax    = String(pax || '').trim().toUpperCase();

  var idMatches = [];
  for (var i = 0; i < all.length; i++) {
    if (String(all[i].txnId).trim().toUpperCase() === wantId) {
      idMatches.push(all[i]);
    }
  }
  if (!idMatches.length) return null;
  if (idMatches.length === 1) return idMatches[0];

  // Narrow by ticket number when one was supplied
  var pool = idMatches;
  if (wantTicket) {
    var byTicket = pool.filter(function (r) {
      return String(r.ticketNo).trim() === wantTicket;
    });
    if (byTicket.length) pool = byTicket;
  }
  if (pool.length === 1) return pool[0];

  // Still ambiguous: match on passenger name
  if (wantPax) {
    for (var j = 0; j < pool.length; j++) {
      if (String(pool[j].pax).trim().toUpperCase() === wantPax) return pool[j];
    }
  }
  return pool[0];
}

/**
 * Main form submission processor.
 * params: { txnId, requestedBy, requestedByEmail, customerEmail,
 *           refundType, airline, issuedFrom, amountAdvice, comment }
 */
function processRefundRequest_(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var txnId = String(params.txnId || '').trim();
    if (!txnId) return { success: false, message: 'Transaction ID is required.' };

    var ticketNo = String(params.ticketNo || '').trim();
    var paxHint  = String(params.paxName  || '').trim();
    var refund   = lookupRefundData_(txnId, ticketNo, paxHint);
    if (!refund) {
      return {
        success: false,
        message: 'Transaction ID "' + txnId + '" was not found. Please check and try again.'
      };
    }

    // ── Duplicate guard ───────────────────────────────────────
    // Keyed on TXN ID *and* ticket number: a 2025 PNR can cover several
    // passengers, and each of those may need its own refund request.
    var sheet   = requireSheet_(REQUESTS(), 'Requests');
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var block = sheet.getRange(2, 1, lastRow - 1, COL.LAST_COL).getValues();
      var newTicket = String(refund.ticketNo || '').trim();

      for (var i = 0; i < block.length; i++) {
        var rowTxn    = String(block[i][COL.TXN_ID - 1] || '').trim();
        var rowTicket = String(block[i][COL.TICKET_NUMBER - 1] || '').trim();
        if (!rowTxn) continue;

        var rowPax = String(block[i][COL.PAX_NAME - 1] || '').trim();
        var newPax = String(refund.pax || '').trim();

        if (rowTxn.toUpperCase() === txnId.toUpperCase() &&
            rowTicket === newTicket &&
            rowPax.toUpperCase() === newPax.toUpperCase()) {
          return {
            success: false,
            message: 'A refund request for TXN "' + txnId + '"'
                   + (newTicket ? ' / ticket ' + newTicket : '')
                   + (newPax ? ' / ' + newPax : '')
                   + ' already exists (row ' + (i + 2) + '). Duplicate blocked.'
          };
        }
      }
    }

    var requestId = generateRequestId_();
    var targetRow = firstEmptyRequestRow_();
    var now       = new Date();

    var airlineName    = String(params.airline || refund.airline || '').trim();
    var requesterEmail = String(params.requestedByEmail || '').trim();
    var customerEmails = cleanEmails_(params.customerEmail);   // multi-email support

    var rowData = new Array(COL.LAST_COL).fill('');
    rowData[COL.REQ_ID - 1]          = requestId;
    rowData[COL.REQUEST_DATE - 1]    = now;                       // stage 1
    rowData[COL.REQUESTED_BY - 1]    = String(params.requestedBy || '').trim();
    rowData[COL.TXN_ID - 1]         = txnId;
    rowData[COL.TXN_DATE - 1]        = refund.date || '';
    rowData[COL.TICKET_NUMBER - 1]   = refund.ticketNo || '';
    rowData[COL.PAX_NAME - 1]        = refund.pax      || '';
    rowData[COL.CLIENT - 1]          = refund.client   || '';
    rowData[COL.AIRLINE - 1]         = airlineName;
    rowData[COL.REFUND_TYPE - 1]     = String(params.refundType || '').trim();
    // Issued From: what the TC confirmed on the form, else the source record.
    rowData[COL.ISSUED_FROM - 1]     = String(params.issuedFrom || refund.issuedFrom || '').trim();
    rowData[COL.COMMENT - 1]         = String(params.comment || '').trim();
    rowData[COL.CUSTOMER_EMAIL - 1]  = customerEmails;
    rowData[COL.STATUS - 1]          = STATUS.PENDING;

    // Amount advised to the client (col U). Stored as a real number so the
    // dashboard can sum it; blank when the TC left the field empty.
    var advice = toNumber_(params.amountAdvice);
    rowData[COL.AMOUNT_ADVICE - 1]   = (advice === null ? '' : advice);
    // COL.DISBURSEMENT (N) intentionally left blank — the refund
    // manager selects the settlement route in the sheet.
    rowData[COL.MODIFIED_BY - 1]     = String(params.requestedBy || '').trim();
    rowData[COL.REQUESTER_EMAIL - 1] = requesterEmail;
    rowData[COL.LAST_MODIFIED - 1]   = now;

    // ── Clear stale dropdown validation on auto-populated columns ──
    // Columns D through I (TXN ID, transaction date, ticket, pax, client,
    // airline) carry legacy "reject input" dropdowns sourced from
    // REFUNDS!A:A and similar. Those lists do not contain 2025 archive
    // PNRs, and the D:E rule wrongly validates the DATE column against the
    // TXN ID list, so any write was rejected. These fields are filled from
    // the source record, never typed by hand, so validation is not needed.
    sheet.getRange(targetRow, COL.TXN_ID, 1, 6).setDataValidation(null);

    sheet.getRange(targetRow, 1, 1, COL.LAST_COL).setValues([rowData]);
    sheet.getRange(targetRow, COL.APPROVED_NGN).setNumberFormat('#,##0.00');
    sheet.getRange(targetRow, COL.PROFIT).setNumberFormat('#,##0.00');
    sheet.getRange(targetRow, COL.AMOUNT_ADVICE).setNumberFormat('#,##0.00');
    if (refund.date instanceof Date) {
      sheet.getRange(targetRow, COL.TXN_DATE).setNumberFormat('dd/mm/yyyy');
    }
    SpreadsheetApp.flush();

    sendNewRequestEmail_(requestId, refund, params, airlineName, customerEmails);

    return {
      success:   true,
      requestId: requestId,
      message:   'Refund request ' + requestId + ' submitted successfully.'
    };

  } finally {
    lock.releaseLock();
  }
}


// ============================================================
//  DATA PROVIDERS  —  called by client-side JS
// ============================================================

/**
 * Server-side transaction search.
 *
 * The form used to download every transaction and filter in the browser.
 * With the sheets now holding about 10,700 records that payload is roughly
 * 2.2 MB, which is slow to serialise and was leaving the dropdown empty.
 * The browser now sends the query and receives only the matches it needs.
 *
 * @param {string} query  free text: TXN ID, PNR, passenger, ticket or client
 * @param {number} limit  maximum rows to return (defaults to 40)
 */
function searchTransactions(query, limit) {
  var q = String(query || '').trim().toLowerCase();
  var max = limit || 40;

  var all = allTransactions_();
  var out = [];

  // No query: show the most recent entries so the box is never blank.
  if (!q) {
    for (var i = all.length - 1; i >= 0 && out.length < max; i--) out.push(all[i]);
    return { total: all.length, shown: out.length, results: out };
  }

  var terms = q.split(/\s+/).filter(String);

  for (var j = 0; j < all.length; j++) {
    var t = all[j];
    var hay = (t.txnId + ' ' + t.pax + ' ' + t.ticketNo + ' ' + t.client).toLowerCase();

    var ok = true;
    for (var k = 0; k < terms.length; k++) {
      if (hay.indexOf(terms[k]) === -1) { ok = false; break; }
    }
    if (!ok) continue;

    out.push(t);
    if (out.length >= max) break;
  }

  return { total: all.length, shown: out.length, results: out };
}

/**
 * Kept for any caller that still wants the whole list. The form no longer
 * uses this: returning 10,700 records to the browser is what broke it.
 */
function getTxnOptions() {
  return allTransactions_();
}

/** Clears the chunked transaction cache and rebuilds it. */
function refreshTxnCache() {
  var cache = CacheService.getScriptCache();
  try {
    var head = cache.get(TXN_CACHE_KEY + '_N');
    if (head) {
      var n = parseInt(head, 10), keys = [TXN_CACHE_KEY + '_N'];
      for (var i = 0; i < n; i++) keys.push(TXN_CACHE_KEY + '_' + i);
      cache.removeAll(keys);
    }
  } catch (e) { Logger.log('Cache clear: ' + e.message); }

  var all = allTransactions_();       // rebuilds and re-caches
  var live = 0, arch = 0;
  for (var j = 0; j < all.length; j++) {
    if (all[j].source === '2025') arch++; else live++;
  }

  notify_('Transaction Cache Refreshed',
    all.length + ' transactions are now searchable in the form.\n\n'
    + '  REFUNDS feed : ' + live + '\n'
    + '  2025 archive : ' + arch);
}

/**
 * Airline list for the form dropdown.
 *
 * Sources, merged and de-duplicated:
 *   1. the AIRLINE sheet, when it exists
 *   2. every airline actually present in REFUNDS and 2025 TICKETS,
 *      resolved from IATA codes to full names
 *   3. the built-in IATA table
 *
 * This is why the form no longer depends on the AIRLINE sheet existing:
 * the previous version called AIRLINE().getLastRow() unconditionally and
 * threw "Cannot read properties of null" when that tab was absent.
 */
function getAirlineOptions() {
  var seen = {};

  // 1. AIRLINE sheet (optional)
  var sheet = AIRLINE();
  if (sheet && sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      var name = String(data[i][1] || '').trim().toUpperCase();
      if (name) seen[name] = true;
    }
  }

  // 2. Airlines actually present in the transaction index. Read from the
  //    cached index rather than re-scanning both sheets, which previously
  //    meant a second full pass over ~10,700 rows on every page load.
  var all = allTransactions_();
  for (var j = 0; j < all.length; j++) {
    if (all[j].airline) seen[all[j].airline] = true;
  }

  // 3. Built-in table, so common carriers are always offered
  for (var k in IATA_MAP) seen[IATA_MAP[k]] = true;

  var out = Object.keys(seen);
  out.sort();
  return out;
}

/**
 * Distinct ISSUED FROM values, gathered from the REFUNDS feed plus any
 * already used on the Requests sheet. Falls back to a sensible default
 * so the dropdown is never empty on a fresh sheet.
 */
function getIssuedFromOptions() {
  var seen = {};

  // From the transaction index (cached), not a fresh scan of REFUNDS
  var all = allTransactions_();
  for (var i = 0; i < all.length; i++) {
    var v = String(all[i].issuedFrom || '').trim().toUpperCase();
    if (v) seen[v] = true;
  }

  // Plus anything already used on the Requests sheet
  var req = REQUESTS();
  if (req && req.getLastRow() >= 2) {
    var b = req.getRange(2, COL.ISSUED_FROM, req.getLastRow() - 1, 1).getDisplayValues();
    for (var j = 0; j < b.length; j++) {
      var w = String(b[j][0] || '').trim().toUpperCase();
      if (w) seen[w] = true;
    }
  }

  if (!Object.keys(seen).length) seen['TICKET'] = true;

  var out = Object.keys(seen);
  out.sort();
  return out;
}

function getRefundTypes() {
  return REFUND_TYPES;
}

/**
 * Reads Disbursement Types live from Data Validation!D2:D.
 * Falls back to the hard-coded list if the column is empty.
 */
function getDisbursementTypes() {
  var sheet = STAFF();
  if (!sheet) return DISBURSEMENT_TYPES_FALLBACK;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return DISBURSEMENT_TYPES_FALLBACK;

  var data = sheet.getRange(2, 4, lastRow - 1, 1).getDisplayValues();  // col D
  var out  = [];
  for (var i = 0; i < data.length; i++) {
    var v = String(data[i][0] || '').trim();
    if (v) out.push(v);
  }
  return out.length ? out : DISBURSEMENT_TYPES_FALLBACK;
}

function getStaffList() {
  var sheet = STAFF();
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data  = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  var staff = [];

  for (var i = 0; i < data.length; i++) {
    var name  = String(data[i][0] || '').trim();
    var email = String(data[i][1] || '').trim();
    if (name && email) staff.push({ name: name, email: email });
  }

  staff.sort(function(a, b) { return a.name.localeCompare(b.name); });
  return staff;
}


// ============================================================
//  EDIT LOG + PROFIT + STAGE TIMESTAMPS + STATUS TRIGGER
// ============================================================

function onEditHandler(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'Requests') return;

  var range = e.range;
  var row   = range.getRow();
  var col   = range.getColumn();
  if (row < 2) return;

  // ── Edit log: manager edits to cols J-X ───────────────────
  if (col >= COL.REFUND_TYPE && col <= COL.PROFIT) {
    var editor = Session.getActiveUser().getEmail() || 'unknown';
    sheet.getRange(row, COL.MODIFIED_BY).setValue(editor);
    // The timestamp column is appended by ensureEditLogColumn(); skip it
    // rather than throw if it has not been created yet.
    if (sheet.getMaxColumns() >= COL.LAST_MODIFIED) {
      sheet.getRange(row, COL.LAST_MODIFIED).setValue(new Date());
    }
  }

  // ── PROFIT auto-calc (col X) ──────────────────────────────
  // Profit = Approved Refund Amount minus what was actually settled.
  // With Amount Paid removed, profit is entered or derived from
  // the approved figure only; recalculated on approved-amount edits.
  if (col === COL.APPROVED_NGN) {
    sheet.getRange(row, COL.PROFIT).setNumberFormat('#,##0.00');
  }

  // ── Status change: col M ──────────────────────────────────
  if (col === COL.STATUS) {
    var newStatus = normaliseStatus_(range.getValue());
    var oldStatus = normaliseStatus_(e.oldValue);
    if (newStatus === oldStatus) return;

    var requesterEmail = String(sheet.getRange(row, COL.REQUESTER_EMAIL).getValue() || '').trim();
    var now = new Date();

    // ── Stage timestamps: stamp once, never overwrite ───────
    // Column P NOW PROCESSING, Q PROCESSING DONE, R PAYOUT COMPLETED,
    // S REFUND ISSUED DATE. SERVICE CHARGE has no dedicated column, so it
    // shares R with POSTED TO TRAVCO: whichever is reached first stamps it.
    function stampOnce_(colIndex) {
      if (!sheet.getRange(row, colIndex).getValue()) {
        sheet.getRange(row, colIndex).setValue(now);
      }
    }

    if (newStatus === STATUS.PROCESSING)     stampOnce_(COL.NOW_PROCESSING);
    if (newStatus === STATUS.PROCESS_DONE)   stampOnce_(COL.PROCESSING_DONE);
    if (newStatus === STATUS.SERVICE_CHARGE) stampOnce_(COL.SERVICE_CHARGE_AT);

    if (newStatus === STATUS.POSTED || newStatus === 'DONE') {
      stampOnce_(COL.PAYOUT_COMPLETED);
      stampOnce_(COL.ISSUED_DATE);
    }

    SpreadsheetApp.flush();
    var rowData = sheet.getRange(row, 1, 1, COL.LAST_COL).getValues()[0];

    if (newStatus === STATUS.PROCESSING)     sendProcessingEmail_(row, rowData, requesterEmail);
    if (newStatus === STATUS.PROCESS_DONE)   sendProcessingDoneEmail_(row, rowData, requesterEmail);
    if (newStatus === STATUS.SERVICE_CHARGE) sendServiceChargeEmail_(row, rowData, requesterEmail);
    if (newStatus === STATUS.POSTED || newStatus === 'DONE') {
      sendPostedEmail_(row, rowData, requesterEmail);
    }
  }
}


/**
 * One-time repair: strips the legacy "reject input" dropdown validations
 * from the auto-populated columns D through I on the Requests sheet.
 *
 * Why this is needed: those rules were sourced from REFUNDS!A:A, and the
 * range D202:E993 wrongly applied the TXN ID list to the transaction DATE
 * column as well. Any form submission carrying a 2025 archive PNR, or a
 * real date in column E, was rejected outright.
 *
 * Manager-controlled dropdowns are left untouched:
 *   J  Refund Type
 *   M  Processing Status
 *   N  Disbursement Type
 *   V  Currency Type
 *
 * Run once from: Refund Tools > Fix Column Validations
 */
function fixColumnValidations() {
    var sheet = requireSheet_(REQUESTS(), 'Requests');

  var lastRow = Math.max(sheet.getLastRow(), 1000);

  // Columns D(4) through I(9): TXN ID, transaction date, ticket,
  // pax, client, airline — all written from the source record.
  sheet.getRange(2, COL.TXN_ID, lastRow - 1, 6).setDataValidation(null);
  SpreadsheetApp.flush();

  notify_('Validations Fixed',
    'Cleared the reject-input dropdowns on columns D to I '
    + '(TXN ID, Transaction Date, Ticket Number, Pax, Client, Airline) '
    + 'for rows 2 to ' + lastRow + '.\n\n'
    + 'Refund Type, Processing Status, Disbursement Type and Currency '
    + 'dropdowns were left in place.');
}

/**
 * ONE-TIME MIGRATION to the five-stage pipeline.
 *
 * Does three things, all reported before anything is written:
 *   1. Applies the new PROCESSING STATUS dropdown across the whole column.
 *      The new list was only ever applied to cell N199, so every other row
 *      still offered the retired four-value list.
 *   2. Renames PROCESSING DONE -> PROCESS DONE.
 *   3. Rewrites OUT and PAY OUT COMPLETED -> DONE.
 *
 * DONE itself is deliberately left alone. 179 rows carry it, and rewriting
 * them to POSTED TO TRAVCO would assert those refunds reached Travco, which
 * the historical data does not record. Use migrateDoneToPosted() separately
 * if that assertion is in fact correct.
 *
 * Run from: Refund Tools > Migrate Status Pipeline
 */
function migrateStatusPipeline() {
    var sheet = requireSheet_(REQUESTS(), 'Requests');
  var lastRow = Math.max(sheet.getLastRow(), 990);

  // ── 1. Rebuild the dropdown across the entire status column ──
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_ORDER, true)
    .setAllowInvalid(true)     // legacy DONE rows must not be flagged red
    .setHelpText('Pipeline: ' + STATUS_ORDER.join(' -> '))
    .build();

  sheet.getRange(2, COL.STATUS, lastRow - 1, 1).setDataValidation(rule);

  // ── 2 & 3. Normalise the retired values ──
  var vals = sheet.getRange(2, COL.STATUS, lastRow - 1, 1).getValues();
  var out = [], renamed = 0, closed = 0, cleared = 0;

  for (var i = 0; i < vals.length; i++) {
    var raw = String(vals[i][0] || '').trim();
    var up  = raw.toUpperCase();

    // A stray header value landed in a data row at some point
    if (up === 'PROCESSING STATUS') { out.push(['']); cleared++; continue; }

    if (up === 'PROCESSING DONE')  { out.push([STATUS.PROCESS_DONE]); renamed++; continue; }
    if (LEGACY_STATUS_MAP[up])     { out.push([LEGACY_STATUS_MAP[up]]); closed++;  continue; }

    out.push([vals[i][0]]);
  }

  sheet.getRange(2, COL.STATUS, out.length, 1).setValues(out);
  SpreadsheetApp.flush();

  notify_('Status Pipeline Migrated',
    'Dropdown applied to N2:N' + lastRow + ':\n  '
    + STATUS_ORDER.join('\n  ') + '\n\n'
    + renamed + ' row(s) renamed PROCESSING DONE -> PROCESS DONE\n'
    + closed  + ' row(s) rewritten OUT / PAY OUT COMPLETED -> DONE\n'
    + cleared + ' stray header value(s) cleared\n\n'
    + 'Rows already reading DONE were left as they are. Run '
    + '"Migrate DONE to POSTED TO TRAVCO" only if every one of those '
    + 'refunds did reach Travco.');
}

/**
 * Finds non-numeric values sitting in the money columns.
 *
 * Column Z currently holds two email addresses at rows 46 and 51, left behind
 * when the edit-log column moved. Text in a numeric column breaks any
 * SUMPRODUCT that multiplies it, which silently zeroed the Approved Refund
 * figure on the dashboard until the formulas were hardened.
 *
 * Read only. Reports what it finds; clearing is a separate confirmed step.
 */
function findStrayTextInMoneyColumns() {
    var sheet = requireSheet_(REQUESTS(), 'Requests');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var cols = [
    [COL.AMOUNT_ADVICE, 'L  AMOUNT ADVICE TO THE CLIENT'],
    [COL.APPROVED_NGN,  'Z  APPROVED REFUND AMOUNT (NGN)'],
    [COL.PROFIT,        'AA PROFIT (NAIRA)']
  ];

  var lines = [], total = 0;
  cols.forEach(function (spec) {
    var vals = sheet.getRange(2, spec[0], lastRow - 1, 1).getValues();
    var hits = [];
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i][0];
      if (v === '' || v === null) continue;
      if (typeof v === 'number') continue;
      if (v instanceof Date) { hits.push((i + 2) + ' (date)'); continue; }
      hits.push((i + 2) + ' "' + String(v).substring(0, 30) + '"');
    }
    total += hits.length;
    lines.push(spec[1] + ' : ' + hits.length + ' stray value(s)');
    hits.slice(0, 8).forEach(function (h) { lines.push('      row ' + h); });
    if (hits.length > 8) lines.push('      ... and ' + (hits.length - 8) + ' more');
  });

  var msg = lines.join('\n') + '\n\n'
          + (total
              ? 'Clear these cells so the totals are trustworthy. The dashboard '
              + 'formulas now coerce text to zero rather than failing outright, '
              + 'but the underlying values are still wrong.'
              : 'All three money columns are clean.');

  Logger.log(msg);
  notify_('Stray Values in Money Columns', msg);
}

/**
 * Re-creates the LAST DATE MODIFIED column at AD and refreshes the
 * PROCESSING STATUS dropdown to the corrected five-stage list.
 *
 * The column is APPENDED, never inserted, so no existing column shifts.
 * Run from: Refund Tools > Repair Sheet Structure
 */
function repairSheetStructure() {
    var sheet = requireSheet_(REQUESTS(), 'Requests');
  var msg = [];

  // 1. Ensure the sheet is wide enough
  if (sheet.getMaxColumns() < COL.LAST_COL) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), COL.LAST_COL - sheet.getMaxColumns());
    msg.push('Widened the sheet to ' + COL.LAST_COL + ' columns.');
  }

  // 2. Header for the edit-log timestamp
  var hdr = sheet.getRange(1, COL.LAST_MODIFIED);
  if (String(hdr.getValue() || '').trim() === '') {
    hdr.setValue('LAST DATE MODIFIED');
    var ref = sheet.getRange(1, COL.MODIFIED_BY);
    hdr.setFontFamily(ref.getFontFamily())
       .setFontSize(ref.getFontSize())
       .setFontWeight('bold')
       .setFontColor(ref.getFontColor())
       .setBackground(ref.getBackground())
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setWrap(true);
    msg.push('Created LAST DATE MODIFIED at column AD.');
  } else {
    msg.push('LAST DATE MODIFIED already present.');
  }

  // 3. Rebuild the status dropdown across the whole column
  var lastRow = Math.max(sheet.getLastRow(), 1000);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_ORDER, true)
    .setAllowInvalid(true)
    .setHelpText('Pipeline: ' + STATUS_ORDER.join(' -> '))
    .build();
  sheet.getRange(2, COL.STATUS, lastRow - 1, 1).setDataValidation(rule);
  msg.push('Status dropdown rebuilt on O2:O' + lastRow + '.');

  SpreadsheetApp.flush();
  notify_('Sheet Structure Repaired', msg.join('\n'));
}

/**
 * Backfill for rows that reached SERVICE CHARGE before column AC existed.
 * Those rows had the service-charge moment written into PAYOUT COMPLETED,
 * so the stage 3 and stage 4 splits read as zero. This copies that value
 * across where the row is still sitting at SERVICE CHARGE.
 */
function backfillServiceChargeDates() {
    var sheet = requireSheet_(REQUESTS(), 'Requests');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var n = lastRow - 1;
  var status = sheet.getRange(2, COL.STATUS,            n, 1).getValues();
  var payout = sheet.getRange(2, COL.PAYOUT_COMPLETED,  n, 1).getValues();
  var scDone = sheet.getRange(2, COL.SERVICE_CHARGE_DONE, n, 1).getValues();
  var svc    = sheet.getRange(2, COL.SERVICE_CHARGE_AT, n, 1).getValues();

  var out = [], moved = 0;
  for (var i = 0; i < n; i++) {
    var st = normaliseStatus_(status[i][0]);
    if (!svc[i][0] && scDone[i][0]) {
      // Prefer the legacy SERVICE CHARGE COMPLETED value where one exists
      out.push([scDone[i][0]]);
      moved++;
    } else if (st === STATUS.SERVICE_CHARGE && !svc[i][0] && payout[i][0]) {
      out.push([payout[i][0]]);
      moved++;
    } else {
      out.push([svc[i][0]]);
    }
  }

  sheet.getRange(2, COL.SERVICE_CHARGE_AT, n, 1).setValues(out);
  SpreadsheetApp.flush();
  notify_('Service Charge Dates Backfilled',
    moved + ' row(s) updated. Rows already carrying a service charge date '
    + 'were left untouched.');
}

/**
 * OPTIONAL second migration: rewrites the legacy DONE terminal state to
 * POSTED TO TRAVCO. Only run this if every DONE row genuinely reached
 * Travco, because it cannot be distinguished afterwards.
 */
function migrateDoneToPosted() {
    var sheet = requireSheet_(REQUESTS(), 'Requests');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var vals = sheet.getRange(2, COL.STATUS, lastRow - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim().toUpperCase() === 'DONE') count++;
  }

  if (!count) { notify_('Refund Tools','No rows are set to DONE. Nothing to migrate.'); return; }

  if (!confirm_('Confirm Migration',
        count + ' row(s) currently read DONE.\n\n'
        + 'Rewriting them to POSTED TO TRAVCOM asserts that each of those refunds '
        + 'was posted to Travcom. This cannot be reversed once applied.\n\n'
        + 'Proceed?')) return;

  var out = [];
  for (var j = 0; j < vals.length; j++) {
    out.push(String(vals[j][0] || '').trim().toUpperCase() === 'DONE'
      ? [STATUS.POSTED] : [vals[j][0]]);
  }
  sheet.getRange(2, COL.STATUS, out.length, 1).setValues(out);
  SpreadsheetApp.flush();

  notify_('Migration Complete', count + ' row(s) migrated to POSTED TO TRAVCOM.');
}


// ============================================================
//  EMAIL ALERT ENGINE
// ============================================================

/**
 * Call-to-action block linking straight to the tracker.
 * Appended to the internal and requester copies only — customers must
 * never receive a link into the internal sheet.
 */
function trackerLinkBlock_() {
  return '<table cellpadding="0" cellspacing="0" style="width:100%;margin:20px 0 4px;">'
    + '<tr><td style="background:#FDF0E2;border:1px solid #F0DFC8;border-radius:6px;padding:16px 18px;">'
    +   '<div style="font-size:13px;color:#23292B;margin-bottom:12px;">'
    +     '<strong>Update the progress status</strong><br>'
    +     '<span style="color:#666;">Open the tracker to set the Processing Status, '
    +     'Disbursement Type and approved amount.</span>'
    +   '</div>'
    +   '<a href="' + TRACKER_URL + '" '
    +      'style="display:inline-block;background:#F58220;color:#ffffff;'
    +      'text-decoration:none;padding:10px 20px;border-radius:5px;'
    +      'font-size:13px;font-weight:bold;">Open Refund Tracker &rarr;</a>'
    + '</td></tr></table>';
}

function sendSplitEmail_(subject, fullPairs, custPairs, title, introPara, closingPara, customerEmail, requesterEmail) {
  // Fetch the logo once per send and embed it inline in every copy.
  var logo    = getLogoBlob_();
  var hasLogo = !!logo;
  var inline  = hasLogo ? { wgslogo: logo } : undefined;

  // Internal and requester copies carry the tracker link; the customer copy
  // does not, so no client is ever handed a route into the internal sheet.
  var internalHtml = emailWrapper_(
    title,
    introPara + detailRows_(fullPairs) + closingPara + trackerLinkBlock_(),
    hasLogo);

  function opts_(extra) {
    var o = {
      replyTo: EMAIL_REPLY_TO,
      name:    'WGS Refund Operations'
    };
    if (inline) o.inlineImages = inline;
    for (var k in extra) o[k] = extra[k];
    return o;
  }

  // 1. Internal
  try {
    GmailApp.sendEmail(EMAIL_TO_INTERNAL, subject, '',
      opts_({ htmlBody: internalHtml, cc: EMAIL_CC }));
    Logger.log('Internal email sent: ' + subject);
  } catch (e) {
    Logger.log('Internal email FAILED: ' + e.message);
    toast_('Email alert failed: ' + e.message, 'Email Error', 10);
  }

  // 2. Requester
  if (requesterEmail && requesterEmail.indexOf('@') > 0) {
    try {
      GmailApp.sendEmail(requesterEmail, subject, '',
        opts_({ htmlBody: internalHtml }));
      Logger.log('Requester email sent to: ' + requesterEmail);
    } catch (e) {
      Logger.log('Requester email FAILED: ' + e.message);
    }
  }

  // 3. Customer(s) — supports multiple addresses
  var custList = cleanEmails_(customerEmail);
  if (custList) {
    var custHtml = emailWrapper_(title, introPara + detailRows_(custPairs) + closingPara, hasLogo);
    try {
      GmailApp.sendEmail(custList, subject, '',
        opts_({ htmlBody: custHtml, cc: EMAIL_TO_INTERNAL }));
      Logger.log('Customer email sent to: ' + custList);
    } catch (e) {
      Logger.log('Customer email FAILED: ' + e.message);
    }
  }
}

function emailWrapper_(title, bodyHtml, hasLogo) {
  // When hasLogo is true the caller attached the logo as an inline image
  // under the content id "wgslogo"; otherwise fall back to the wordmark.
  var brand = hasLogo
    ? '<img src="cid:wgslogo" width="' + LOGO_WIDTH_PX + '" alt="WGS" '
      + 'style="display:block;border:0;outline:none;max-width:' + LOGO_WIDTH_PX + 'px;height:auto;">'
    : '<span style="font-family:Garamond,Georgia,serif;font-size:22px;font-weight:bold;">'
      + '<span style="color:#23292B;">WG</span>'
      + '<span style="color:#F58220;">S</span></span>';

  return '<!DOCTYPE html><html><head><meta charset="utf-8"></head>'
    + '<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,Helvetica,sans-serif;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:24px 0;">'
    + '<tr><td align="center">'
    + '<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;overflow:hidden;">'

    // Header with logo — white so the logo artwork and text stay legible
    + '<tr><td style="background:#FFFFFF;padding:18px 32px;border-bottom:1px solid #E6E6E6;">'
    +   '<table cellpadding="0" cellspacing="0"><tr>'
    +     '<td style="vertical-align:middle;">' + brand + '</td>'
    +     '<td style="vertical-align:middle;padding-left:14px;color:#5A5A5A;font-size:12px;'
    +       'border-left:1px solid #E6E6E6;margin-left:14px;">'
    +       '&nbsp;Refund Operations</td>'
    +   '</tr></table>'
    + '</td></tr>'

    + '<tr><td style="background:#FDF0E2;padding:14px 32px;">'
    + '<span style="color:#23292B;font-family:Garamond,Georgia,serif;font-size:18px;font-weight:bold;">' + title + '</span>'
    + '</td></tr>'
    + '<tr><td style="padding:24px 32px;color:#23292B;font-size:14px;line-height:1.6;">' + bodyHtml + '</td></tr>'
    + '<tr><td style="background:#23292B;padding:14px 32px;color:#808080;font-size:11px;">'
    + 'Automated notification from the WGS Operations Team. '
    + 'Enquiries: <a href="mailto:customersuccess@example.com" style="color:#F58220;">customersuccess@example.com</a>'
    + '</td></tr>'
    + '</table></td></tr></table></body></html>';
}

function detailRows_(pairs) {
  var html = '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;margin:12px 0;">';
  for (var i = 0; i < pairs.length; i++) {
    var bg = (i % 2 === 0) ? '#FAFAFA' : '#FFFFFF';
    html += '<tr style="background:' + bg + ';">'
      + '<td style="padding:8px 12px;font-weight:bold;color:#555;width:40%;border-bottom:1px solid #eee;">' + pairs[i][0] + '</td>'
      + '<td style="padding:8px 12px;color:#23292B;border-bottom:1px solid #eee;">' + (pairs[i][1] || '—') + '</td>'
      + '</tr>';
  }
  return html + '</table>';
}

function fmtNaira_(val) {
  var num = toNumber_(val);
  if (num === null) return '—';
  return '₦' + num.toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fmtDate_(val) {
  if (!val) return '—';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd MMM yyyy');
  return String(val);
}

function badge_(label, color) {
  return '<span style="background:' + color + ';color:#fff;padding:3px 10px;border-radius:4px;font-size:12px;">' + label + '</span>';
}

/** Builds a turnaround summary block for the emails. */
function turnaroundPairs_(rowData) {
  var req  = rowData[COL.REQUEST_DATE - 1];
  var proc = rowData[COL.NOW_PROCESSING - 1];
  var pdon = rowData[COL.PROCESSING_DONE - 1];
  var svc  = rowData[COL.SERVICE_CHARGE_AT - 1];
  var pout = rowData[COL.PAYOUT_COMPLETED - 1];

  var pairs = [];
  var d1 = daysBetween_(proc, req);
  var d2 = daysBetween_(pdon, proc);
  var d3 = daysBetween_(svc,  pdon);
  var d4 = daysBetween_(pout, svc);
  var dT = daysBetween_(pout, req);

  // A refund that skipped the service charge stage still needs a sensible
  // stage 3: fall back to Process done straight through to Posted.
  if (d3 === null && d4 === null) d4 = daysBetween_(pout, pdon);

  if (d1 !== null) pairs.push(['Stage 1: Request → Processing',      d1 + ' days']);
  if (d2 !== null) pairs.push(['Stage 2: Processing → Process done', d2 + ' days']);
  if (d3 !== null) pairs.push(['Stage 3: Process done → Service charge', d3 + ' days']);
  if (d4 !== null) pairs.push(['Stage 4: Service charge → Posted',   d4 + ' days']);
  if (dT !== null) pairs.push(['Total turnaround',                        dT + ' days']);

  return pairs;
}



// ── 1. NEW REQUEST SUBMITTED ─────────────────────────────────

function sendNewRequestEmail_(requestId, refund, params, airlineName, customerEmails) {
  var subject = '🔔 New Refund Request: ' + requestId + ' | ' + (refund.pax || 'N/A');
  var sb      = badge_('PENDING', '#E65100');

  var fullPairs = [
    ['Request ID',        requestId],
    ['Transaction ID',    refund.txnId    || '—'],
    ['Passenger',         refund.pax      || '—'],
    ['Ticket Number',     refund.ticketNo || '—'],
    ['Client',            refund.client   || '—'],
    ['Airline',           airlineName     || '—'],
    ['Issued From',       (params.issuedFrom || refund.issuedFrom) || '—'],
    ['Refund Type',       params.refundType || '—'],
    ['Amount Advice to Client', fmtNaira_(params.amountAdvice)],
    ['Disbursement Type', 'To be assigned by refund manager'],
    ['Reason / Comment',  params.comment  || '—'],
    ['Requested By',      params.requestedBy || '—'],
    ['Customer Email(s)', customerEmails  || '—'],
    ['Status',            sb],
    ['Request Date',      fmtDate_(new Date())]
  ];

  var custPairs = [
    ['Request ID',        requestId],
    ['Transaction ID',    refund.txnId    || '—'],
    ['Passenger',         refund.pax      || '—'],
    ['Ticket Number',     refund.ticketNo || '—'],
    ['Issued From',       (params.issuedFrom || refund.issuedFrom) || '—'],
    ['Refund Type',       params.refundType || '—'],
    ['Reason / Comment',  params.comment  || '—'],
    ['Status',            sb],
    ['Request Date',      fmtDate_(new Date())]
  ];

  var intro   = '<p>A new refund request has been submitted and is awaiting review.</p>';
  var closing = '<p style="margin-top:16px;color:#555;">Please review and update the Processing Status in the tracker to proceed.</p>';

  sendSplitEmail_(subject, fullPairs, custPairs, 'New Refund Request Submitted',
    intro, closing, customerEmails, params.requestedByEmail);
}


// ── 2. PROCESSING ────────────────────────────────────────────

function sendProcessingEmail_(row, rowData, requesterEmail) {
  var requestId     = rowData[COL.REQ_ID - 1]        || '';
  var pax           = rowData[COL.PAX_NAME - 1]       || '';
  var customerEmail = rowData[COL.CUSTOMER_EMAIL - 1] || '';
  var sb            = badge_('PROCESSING', '#1565C0');
  var subject       = '⏳ Refund Processing: ' + requestId + ' | ' + pax;

  var base = [
    ['Request ID',        requestId],
    ['Transaction ID',    rowData[COL.TXN_ID - 1]        || '—'],
    ['Passenger',         pax],
    ['Ticket Number',     rowData[COL.TICKET_NUMBER - 1] || '—'],
    ['Issued From',       rowData[COL.ISSUED_FROM - 1]   || '—'],
    ['Refund Type',       rowData[COL.REFUND_TYPE - 1]   || '—'],
    ['Disbursement Type', rowData[COL.DISBURSEMENT - 1]  || '—'],
    ['Reason / Comment',  rowData[COL.COMMENT - 1]       || '—'],
    ['Status',            sb],
    ['Processing Started', fmtDate_(rowData[COL.NOW_PROCESSING - 1])]
  ];

  var fullPairs = [
    ['Request ID',        requestId],
    ['Transaction ID',    rowData[COL.TXN_ID - 1]        || '—'],
    ['Passenger',         pax],
    ['Ticket Number',     rowData[COL.TICKET_NUMBER - 1] || '—'],
    ['Client',            rowData[COL.CLIENT - 1]        || '—'],
    ['Airline',           rowData[COL.AIRLINE - 1]       || '—'],
    ['Issued From',       rowData[COL.ISSUED_FROM - 1]   || '—'],
    ['Refund Type',       rowData[COL.REFUND_TYPE - 1]   || '—'],
    ['Amount Advice to Client', fmtNaira_(rowData[COL.AMOUNT_ADVICE - 1])],
    ['Disbursement Type', rowData[COL.DISBURSEMENT - 1]  || '—'],
    ['Reason / Comment',  rowData[COL.COMMENT - 1]       || '—'],
    ['Status',            sb],
    ['Processing Started', fmtDate_(rowData[COL.NOW_PROCESSING - 1])]
  ].concat(turnaroundPairs_(rowData));

  var intro   = '<p>The following refund request is now being processed.</p>';
  var closing = '<p style="margin-top:16px;color:#555;">You will be notified when processing is complete.</p>';

  sendSplitEmail_(subject, fullPairs, base, 'Refund Processing Started',
    intro, closing, customerEmail, requesterEmail);
}


// ── 3. PROCESSING DONE ───────────────────────────────────────

function sendProcessingDoneEmail_(row, rowData, requesterEmail) {
  var requestId     = rowData[COL.REQ_ID - 1]        || '';
  var pax           = rowData[COL.PAX_NAME - 1]       || '';
  var customerEmail = rowData[COL.CUSTOMER_EMAIL - 1] || '';
  var sb            = badge_('PROCESSING DONE', '#7B4FBF');
  var subject       = '📋 Refund Processing Complete: ' + requestId + ' | ' + pax;

  var custPairs = [
    ['Request ID',        requestId],
    ['Transaction ID',    rowData[COL.TXN_ID - 1]        || '—'],
    ['Passenger',         pax],
    ['Ticket Number',     rowData[COL.TICKET_NUMBER - 1] || '—'],
    ['Issued From',       rowData[COL.ISSUED_FROM - 1]   || '—'],
    ['Refund Type',       rowData[COL.REFUND_TYPE - 1]   || '—'],
    ['Disbursement Type', rowData[COL.DISBURSEMENT - 1]  || '—'],
    ['Reason / Comment',  rowData[COL.COMMENT - 1]       || '—'],
    ['Status',            sb],
    ['Processing Done',   fmtDate_(rowData[COL.PROCESSING_DONE - 1])]
  ];

  var fullPairs = [
    ['Request ID',        requestId],
    ['Transaction ID',    rowData[COL.TXN_ID - 1]        || '—'],
    ['Passenger',         pax],
    ['Ticket Number',     rowData[COL.TICKET_NUMBER - 1] || '—'],
    ['Client',            rowData[COL.CLIENT - 1]        || '—'],
    ['Airline',           rowData[COL.AIRLINE - 1]       || '—'],
    ['Issued From',       rowData[COL.ISSUED_FROM - 1]   || '—'],
    ['Refund Type',       rowData[COL.REFUND_TYPE - 1]   || '—'],
    ['Disbursement Type', rowData[COL.DISBURSEMENT - 1]  || '—'],
    ['Amount Advice to Client', fmtNaira_(rowData[COL.AMOUNT_ADVICE - 1])],
    ['Approved Refund',   fmtNaira_(rowData[COL.APPROVED_NGN - 1])],
    ['Reason / Comment',  rowData[COL.COMMENT - 1]       || '—'],
    ['Status',            sb],
    ['Processing Done',   fmtDate_(rowData[COL.PROCESSING_DONE - 1])],
    ['BSP RA No',         rowData[COL.BSP_RA_NO - 1]     || '—']
  ].concat(turnaroundPairs_(rowData));

  var intro   = '<p>Processing is complete for the following refund request. It is now awaiting settlement.</p>';
  var closing = '<p style="margin-top:16px;color:#555;">The refund will be settled according to the selected disbursement type.</p>';

  sendSplitEmail_(subject, fullPairs, custPairs, 'Refund Processing Complete',
    intro, closing, customerEmail, requesterEmail);
}


// ── 4. SERVICE CHARGE ────────────────────────────────────────

function sendServiceChargeEmail_(row, rowData, requesterEmail) {
  var requestId     = rowData[COL.REQ_ID - 1]        || '';
  var pax           = rowData[COL.PAX_NAME - 1]       || '';
  var customerEmail = rowData[COL.CUSTOMER_EMAIL - 1] || '';
  var sb            = badge_('SERVICE CHARGE', '#B8860B');
  var subject       = '🧾 Refund Service Charge Applied: ' + requestId + ' | ' + pax;

  var custPairs = [
    ['Request ID',        requestId],
    ['Transaction ID',    rowData[COL.TXN_ID - 1]        || '—'],
    ['Passenger',         pax],
    ['Ticket Number',     rowData[COL.TICKET_NUMBER - 1] || '—'],
    ['Issued From',       rowData[COL.ISSUED_FROM - 1]   || '—'],
    ['Refund Type',       rowData[COL.REFUND_TYPE - 1]   || '—'],
    ['Disbursement Type', rowData[COL.DISBURSEMENT - 1]  || '—'],
    ['Reason / Comment',  rowData[COL.COMMENT - 1]       || '—'],
    ['Status',            sb]
  ];

  var fullPairs = [
    ['Request ID',        requestId],
    ['Transaction ID',    rowData[COL.TXN_ID - 1]        || '—'],
    ['Passenger',         pax],
    ['Ticket Number',     rowData[COL.TICKET_NUMBER - 1] || '—'],
    ['Client',            rowData[COL.CLIENT - 1]        || '—'],
    ['Airline',           rowData[COL.AIRLINE - 1]       || '—'],
    ['Issued From',       rowData[COL.ISSUED_FROM - 1]   || '—'],
    ['Refund Type',       rowData[COL.REFUND_TYPE - 1]   || '—'],
    ['Amount Advice to Client', fmtNaira_(rowData[COL.AMOUNT_ADVICE - 1])],
    ['Approved Refund',   fmtNaira_(rowData[COL.APPROVED_NGN - 1])],
    ['Disbursement Type', rowData[COL.DISBURSEMENT - 1]  || '—'],
    ['Reason / Comment',  rowData[COL.COMMENT - 1]       || '—'],
    ['Status',            sb],
    ['BSP RA No',         rowData[COL.BSP_RA_NO - 1]     || '—']
  ].concat(turnaroundPairs_(rowData));

  var intro   = '<p>The service charge has been applied to the following refund. '
              + 'It is now ready to be posted.</p>';
  var closing = '<p style="margin-top:16px;color:#555;">The next and final step is '
              + 'posting to Travco.</p>';

  sendSplitEmail_(subject, fullPairs, custPairs, 'Refund Service Charge Applied',
    intro, closing, customerEmail, requesterEmail);
}


// ── 5. POSTED TO TRAVCO (terminal) ───────────────────────────
//  Also handles the legacy DONE status so historical rows still
//  produce a completion email if their status is ever re-touched.

function sendPostedEmail_(row, rowData, requesterEmail) {
  var requestId     = rowData[COL.REQ_ID - 1]        || '';
  var pax           = rowData[COL.PAX_NAME - 1]       || '';
  var customerEmail = rowData[COL.CUSTOMER_EMAIL - 1] || '';
  var disb          = String(rowData[COL.DISBURSEMENT - 1] || '').toUpperCase();
  var sb            = badge_('POSTED TO TRAVCOM', '#319F4E');
  var subject       = '✅ Refund Completed: ' + requestId + ' | ' + pax;

  // Tailor the closing line to the disbursement route
  var settleLine;
  if (disb.indexOf('LEDGER') >= 0) {
    settleLine = 'The refund value has been credited to the client ledger and posted to Travcom.';
  } else if (disb.indexOf('DEBT') >= 0) {
    settleLine = 'The refund value has been applied against the outstanding client debt '
               + 'and posted to Travcom.';
  } else {
    settleLine = 'The refund has been paid out to the customer account and posted to Travcom.';
  }

  var custPairs = [
    ['Request ID',        requestId],
    ['Transaction ID',    rowData[COL.TXN_ID - 1]        || '—'],
    ['Passenger',         pax],
    ['Ticket Number',     rowData[COL.TICKET_NUMBER - 1] || '—'],
    ['Issued From',       rowData[COL.ISSUED_FROM - 1]   || '—'],
    ['Refund Type',       rowData[COL.REFUND_TYPE - 1]   || '—'],
    ['Disbursement Type', rowData[COL.DISBURSEMENT - 1]  || '—'],
    ['Reason / Comment',  rowData[COL.COMMENT - 1]       || '—'],
    ['Status',            sb],
    ['Completed Date',    fmtDate_(rowData[COL.PAYOUT_COMPLETED - 1])]
  ];

  var fullPairs = [
    ['Request ID',        requestId],
    ['Transaction ID',    rowData[COL.TXN_ID - 1]         || '—'],
    ['Passenger',         pax],
    ['Ticket Number',     rowData[COL.TICKET_NUMBER - 1]  || '—'],
    ['Client',            rowData[COL.CLIENT - 1]         || '—'],
    ['Airline',           rowData[COL.AIRLINE - 1]        || '—'],
    ['Issued From',       rowData[COL.ISSUED_FROM - 1]    || '—'],
    ['Refund Type',       rowData[COL.REFUND_TYPE - 1]    || '—'],
    ['Amount Advice to Client', fmtNaira_(rowData[COL.AMOUNT_ADVICE - 1])],
    ['Disbursement Type', rowData[COL.DISBURSEMENT - 1]   || '—'],
    ['Approved Refund',   fmtNaira_(rowData[COL.APPROVED_NGN - 1])],
    ['Profit',            fmtNaira_(rowData[COL.PROFIT - 1])],
    ['Reason / Comment',  rowData[COL.COMMENT - 1]       || '—'],
    ['Status',            sb],
    ['Completed Date',    fmtDate_(rowData[COL.PAYOUT_COMPLETED - 1])],
    ['BSP RA No',         rowData[COL.BSP_RA_NO - 1]      || '—'],
    ['Manager Remark',    rowData[COL.MANAGER_REMARK - 1] || '—']
  ].concat(turnaroundPairs_(rowData));

  var intro   = '<p>' + settleLine + '</p>';
  var closing = '<p style="margin-top:16px;color:#555;">This request is fully resolved. '
              + 'No further action is required.</p>';

  sendSplitEmail_(subject, fullPairs, custPairs, 'Refund Completed',
    intro, closing, customerEmail, requesterEmail);
}


// ============================================================
//  WEB APP
//  doGet() lives in WebApp.gs, which routes between the portal
//  (…/exec) and this request form (…/exec?page=form).
// ============================================================

function submitRefundRequest(formData) {
  return processRefundRequest_({
    txnId:            formData.txnId,
    ticketNo:         formData.ticketNo,
    paxName:          formData.paxName,
    requestedBy:      formData.requestedBy,
    requestedByEmail: formData.requestedByEmail,
    customerEmail:    formData.customerEmail,
    refundType:       formData.refundType,
    airline:          formData.airline,
    issuedFrom:       formData.issuedFrom,
    amountAdvice:     formData.amountAdvice,
    comment:          formData.comment
  });
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ============================================================
//  SETUP & MENU
// ============================================================

function createEditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEditHandler') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // forSpreadsheet() needs a real Spreadsheet object, so resolve it properly.
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss_()).onEdit().create();
  Logger.log('Edit trigger installed.');
  notify_('Trigger Installed',
    'The onEdit trigger for email alerts, stage timestamps and edit logging '
    + 'has been installed.');
}

function onOpen() {
  var ui = ui_();
  if (!ui) return;          // no menu bar to attach to
  ui
    .createMenu('Refund Tools')
    .addItem('Open Refund Portal', 'openPortalLink')
    .addItem('Open Request Form (sidebar)', 'openFormSidebar')
    .addItem('Install Email Trigger', 'createEditTrigger')
    .addSeparator()
    .addItem('Set Up Refund Portal', 'setupPortal')
    .addItem('Lock Status Columns (portal only)', 'lockSheetToPortal')
    .addItem('Unlock Status Columns', 'unlockSheetFromPortal')
    .addSeparator()
    .addItem('Repair Sheet Structure', 'repairSheetStructure')
    .addItem('Find Stray Text in Money Columns', 'findStrayTextInMoneyColumns')
    .addItem('Migrate Status Pipeline', 'migrateStatusPipeline')
    .addItem('Migrate DONE to POSTED TO TRAVCOM', 'migrateDoneToPosted')
    .addItem('Backfill Service Charge Dates', 'backfillServiceChargeDates')
    .addItem('Refresh Transaction Cache', 'refreshTxnCache')
    .addItem('Fix Column Validations', 'fixColumnValidations')
    .addToUi();
}

/** Shows a clickable link to the portal from the spreadsheet menu. */
function openPortalLink() {
  var url = ScriptApp.getService().getUrl();
  var html = HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif;text-align:center;padding:16px;">'
    + '<a href="' + url + '" target="_blank" '
    + 'style="display:block;padding:12px;background:#F58220;color:#fff;'
    + 'text-decoration:none;border-radius:6px;font-weight:bold;" '
    + 'onclick="setTimeout(function(){google.script.host.close()},400)">'
    + 'Open Refund Portal</a></div>')
    .setWidth(320).setHeight(110);
  ui_().showModalDialog(html, 'Refund Portal');
}

function openFormSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Submit Refund Request')
    .setWidth(460);
  var ui = ui_();
  if (!ui) {
    Logger.log('No UI available. Open the form from its web app URL instead.');
    return;
  }
  ui.showSidebar(html);
}
