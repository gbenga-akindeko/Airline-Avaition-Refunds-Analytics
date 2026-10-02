// ============================================================
//  WGS REFUND PORTAL — WEB APP LAYER
//  WebApp.gs · add to the WGS Refund Management System project,
//  alongside Code.gs and Index.html.
//
//  Moves steps 6 to 10 of the refund flow (and the Finance output
//  step) off the spreadsheet and into a role-based web portal.
//  The Requests sheet stays the single source of truth: the portal
//  reads and writes the same cells the team edits today, stamps the
//  same stage timestamps and sends the same emails from Code.gs.
//
//  IMPORTANT: delete the doGet() function in Code.gs. This file
//  replaces it with a router that serves both the portal and the form.
// ============================================================

const PORTAL_TITLE = 'WGS Refund Portal';

// ── Finance output stage ─────────────────────────────────────
// Appended after AD, never inserted, so no existing column shifts.
// ensureSettlementColumns_() writes the headers the first time.
const WCOL = {
  SETTLED_AT:      31,  // AE
  SETTLED_BY:      32,  // AF
  SETTLEMENT_TYPE: 33,  // AG
  SETTLEMENT_REF:  34,  // AH
  LAST:            34
};
const WCOL_HEADERS = [
  'SETTLEMENT CONFIRMED', 'SETTLED BY', 'SETTLEMENT OUTPUT', 'SETTLEMENT REF / NOTE'
];

const SETTLEMENT_OUTPUTS = [
  'PAID TO CLIENT PERSONAL ACCOUNT',
  'POSTED TO CLIENT LEDGER',
  'APPLIED AGAINST CLIENT DEBT'
];

// Rows posted before this date are treated as closed. Without it every
// historical POSTED / DONE row would land in Finance's settlement queue.
const SETTLEMENT_TRACKING_FROM = new Date('2026-10-01T00:00:00');

// ── Roles ────────────────────────────────────────────────────
// An entry can be a person or a Google Group. If refund@example.com
// is a group, anyone in it gets the role when they sign in with their
// own account. Anyone else on a staff domain is a Travel Consultant.
const ROLE_MEMBERS = {
  ADMIN: [
    'refund@example.com',                      // Process Owner
    'dataanalyst@example.com'                  // system administrator
  ],
  REFUND_MANAGER: [
    'refund@example.com',
    'refunds@example.com',
    'customersuccess@example.com'
  ],
  OPM_HOB: [
    'hob@example.com',                         // Head of Business / MD
    'hoo@example.com',                         // Head of Operations
    'operationsmanager@example.com'
  ],
  FINANCE: [
    'financemanager@example.com',
    'treasury@example.com'
  ],
  OVERSIGHT: [                                   // see everything, change nothing
    'cfo@holdings.example.com',
    'headrac@holdings.example.com',             // Risk, Audit & Control
    'audit.officer@holdings.example.com',
    'team.lead@example.com'                // Team Lead
  ]
};

const ROLE_LABELS = {
  ADMIN:          'Process Owner',
  REFUND_MANAGER: 'Refund Manager',
  OPM_HOB:        'OPM / HOB',
  FINANCE:        'Finance',
  OVERSIGHT:      'Oversight',
  TC:             'Travel Consultant'
};

const STAFF_DOMAINS = ['example.com', 'holdings.example.com'];

// Who may move a request out of each status, matching the flow chart.
// String literals rather than STATUS.* because Apps Script evaluates
// files in order and Code.gs may not have loaded yet.
const TRANSITIONS = {
  'PENDING':        { next: 'PROCESSING',        roles: ['REFUND_MANAGER'] },
  'PROCESSING':     { next: 'PROCESS DONE',      roles: ['REFUND_MANAGER'],
                      require: ['disbursement', 'approved'] },
  'PROCESS DONE':   { next: 'SERVICE CHARGE',    roles: ['OPM_HOB'] },
  'SERVICE CHARGE': { next: 'POSTED TO TRAVCOM', roles: ['OPM_HOB', 'FINANCE'] }
};

// Fields the portal can edit, and which roles may edit each one.
const FIELDS = {
  disbursement:  { col: 16, label: 'Disbursement type',      type: 'list',   roles: ['REFUND_MANAGER', 'OPM_HOB'] },
  approved:      { col: 26, label: 'Approved refund (NGN)',  type: 'money',  roles: ['REFUND_MANAGER', 'OPM_HOB', 'FINANCE'] },
  currency:      { col: 25, label: 'Currency type',          type: 'text',   roles: ['REFUND_MANAGER'] },
  bspRa:         { col: 23, label: 'BSP RA no',              type: 'text',   roles: ['REFUND_MANAGER'] },
  profit:        { col: 27, label: 'Service charge / profit (NGN)', type: 'money', roles: ['OPM_HOB', 'FINANCE'] },
  remark:        { col: 24, label: 'Manager / OPM remark',   type: 'text',   roles: ['REFUND_MANAGER', 'OPM_HOB', 'FINANCE'] },
  customerEmail: { col: 14, label: 'Customer email(s)',      type: 'emails', roles: ['REFUND_MANAGER'] }
};

const SEE_ALL_ROLES = ['ADMIN', 'REFUND_MANAGER', 'OPM_HOB', 'FINANCE', 'OVERSIGHT'];

// ── Super admin ──────────────────────────────────────────────
// The only account(s) that can move a refund BACKWARDS, undo a Finance
// settlement, or edit a closed request. Everyone else, ADMIN included,
// can only move forward. Deliberately code-only: it cannot be granted
// from the Portal Roles sheet.
const SUPER_ADMINS = ['dataanalyst@example.com'];

const PIPE_ORDER = ['PENDING', 'PROCESSING', 'PROCESS DONE', 'SERVICE CHARGE', 'POSTED TO TRAVCOM'];

// Travel Consultants can browse every request (read only) as well as
// submitting new ones. They never see the dashboard, profit figures or
// the tracker link, and they have no status controls. Set to false to
// restrict TCs to their own submissions.
const TC_CAN_VIEW_ALL = true;


// ============================================================
//  ROUTER
// ============================================================

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || 'portal';

  if (page === 'form') {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('WGS Refund Request')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('Portal')
    .setTitle(PORTAL_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
//  IDENTITY & PERMISSIONS
// ============================================================

function currentEmail_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); }
  catch (e) { return ''; }
}

function isGroupMember_(groupEmail, userEmail) {
  try { return GroupsApp.getGroupByEmail(groupEmail).hasUser(userEmail); }
  catch (e) { return false; }   // not a group, or not visible to the script owner
}

// ── Portal Roles sheet ───────────────────────────────────────
// Lets the Process Owner grant roles from the spreadsheet, e.g. give a
// named person REFUND_MANAGER, without editing code. Merged with
// ROLE_MEMBERS above. Changes apply within 5 minutes (role cache).
const ROLES_SHEET = 'Portal Roles';
const ROLE_KEYS   = ['ADMIN', 'REFUND_MANAGER', 'OPM_HOB', 'FINANCE', 'OVERSIGHT'];

function rolesSheet_() {
  var sh = sheet_(ROLES_SHEET);
  if (!sh) {
    sh = ss_().insertSheet(ROLES_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['EMAIL', 'ROLE', 'NAME / NOTE']])
      .setFontWeight('bold').setBackground('#FDF0E2');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 280); sh.setColumnWidth(2, 170); sh.setColumnWidth(3, 260);
    sh.getRange(2, 2, 500, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(ROLE_KEYS, true).setAllowInvalid(false).build());
    sh.getRange(2, 1, 1, 3).setValues([['dataanalyst@example.com', 'ADMIN', 'Example row, edit or delete']]);
  }
  return sh;
}

function sheetRoleMembers_() {
  var out = {};
  var sh = sheet_(ROLES_SHEET);
  if (!sh || sh.getLastRow() < 2) return out;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues();
  for (var i = 0; i < v.length; i++) {
    var e = String(v[i][0] || '').trim().toLowerCase();
    var r = String(v[i][1] || '').trim().toUpperCase();
    if (!e || ROLE_KEYS.indexOf(r) === -1) continue;
    (out[r] = out[r] || []).push(e);
  }
  return out;
}

/** Roles for an email, cached for 5 minutes to keep group lookups off every call. */
function rolesFor_(email) {
  if (!email) return [];
  var cache = CacheService.getScriptCache();
  var key = 'roles_' + email;
  var hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  var fromSheet = sheetRoleMembers_();
  var roles = [];
  ROLE_KEYS.forEach(function (role) {
    var list = (ROLE_MEMBERS[role] || []).concat(fromSheet[role] || []);
    for (var i = 0; i < list.length; i++) {
      var m = list[i].toLowerCase();
      if (m === email || isGroupMember_(m, email)) { roles.push(role); break; }
    }
  });

  var domain = email.split('@')[1] || '';
  if (STAFF_DOMAINS.indexOf(domain) !== -1) roles.push('TC');   // everyone may submit

  cache.put(key, JSON.stringify(roles), 300);
  return roles;
}

function hasAny_(roles, wanted) {
  if (roles.indexOf('ADMIN') !== -1) return true;
  for (var i = 0; i < wanted.length; i++) if (roles.indexOf(wanted[i]) !== -1) return true;
  return false;
}

/** Resolves the caller or throws. Every server entry point starts here. */
function auth_() {
  var email = currentEmail_();
  if (!email) {
    throw new Error('Could not identify your Google account. Sign in with your '
                  + 'WGS account and reload the page.');
  }
  var roles = rolesFor_(email);
  if (!roles.length) throw new Error('Your account (' + email + ') does not have access to the refund portal.');
  var seeAll = hasAny_(roles, SEE_ALL_ROLES);
  return { email: email, roles: roles, seeAll: seeAll,
           isSuper: SUPER_ADMINS.indexOf(email) !== -1,
           viewAll: seeAll || (TC_CAN_VIEW_ALL && roles.indexOf('TC') !== -1) };
}

function staffNameFor_(email) {
  var list = getStaffList();
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].email).trim().toLowerCase() === email) return list[i].name;
  }
  return '';
}


// ============================================================
//  SHEET HELPERS
// ============================================================

function ensureSettlementColumns_() {
  var sheet = requireSheet_(REQUESTS(), 'Requests');
  if (sheet.getMaxColumns() < WCOL.LAST) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), WCOL.LAST - sheet.getMaxColumns());
  }
  var hdr = sheet.getRange(1, WCOL.SETTLED_AT, 1, WCOL_HEADERS.length);
  var cur = hdr.getValues()[0];
  if (cur.join('') === '') {
    hdr.setValues([WCOL_HEADERS]);
    var ref = sheet.getRange(1, COL.MODIFIED_BY);
    hdr.setFontWeight('bold').setBackground(ref.getBackground())
       .setFontColor(ref.getFontColor()).setWrap(true)
       .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  return sheet;
}

function auditSheet_() {
  var sh = sheet_('Audit Log');
  if (!sh) {
    sh = ss_().insertSheet('Audit Log');
    sh.appendRow(['TIMESTAMP', 'USER', 'REQUEST ID', 'ROW', 'ACTION', 'FROM', 'TO', 'DETAILS']);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#FDF0E2');
    sh.setFrozenRows(1);
  }
  return sh;
}

function audit_(user, reqId, row, action, from, to, details) {
  try {
    auditSheet_().appendRow([new Date(), user, String(reqId), row, action, from || '', to || '', details || '']);
  } catch (e) { Logger.log('Audit write failed: ' + e.message); }
}

function iso_(d) {
  return (d instanceof Date && !isNaN(d.getTime())) ? d.toISOString() : '';
}

function isBlankRow_(r) {
  return !r[COL.REQ_ID - 1] && !r[COL.TXN_ID - 1]
      && !r[COL.TICKET_NUMBER - 1] && !r[COL.PAX_NAME - 1];
}

function currentStatus_(r) {
  return normaliseStatus_(r[COL.STATUS - 1]) || 'PENDING';
}

/**
 * Where a row sits in the full flow, including the Finance output step.
 *   SETTLED  Finance confirmed the payout / ledger posting
 *   CLOSED   posted before settlement tracking began, or legacy DONE
 *   POSTED   posted to Travcom, awaiting Finance
 */
function stageOf_(r) {
  var st = currentStatus_(r);
  if (st !== 'POSTED TO TRAVCOM' && st !== 'DONE') return st;
  if (r[WCOL.SETTLED_AT - 1]) return 'SETTLED';
  if (st === 'DONE') return 'CLOSED';
  var posted = r[COL.PAYOUT_COMPLETED - 1];
  if (!(posted instanceof Date) || posted < SETTLEMENT_TRACKING_FROM) return 'CLOSED';
  return 'POSTED';
}

function locateRow_(sheet, reqId, hintRow, ticket) {
  var last = sheet.getLastRow();
  var id = String(reqId || '');
  if (hintRow >= 2 && hintRow <= last) {
    var a = String(sheet.getRange(hintRow, COL.REQ_ID).getValue() || '');
    if (id && a === id) return hintRow;
    if (!id && !a && String(sheet.getRange(hintRow, COL.TICKET_NUMBER).getValue()) === String(ticket || '')) {
      return hintRow;
    }
  }
  if (id) {
    var ids = sheet.getRange(2, COL.REQ_ID, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  }
  throw new Error('Request ' + (id || '(row ' + hintRow + ')') + ' could not be found. Reload the portal.');
}

function rowToObj_(r, row, who, mine) {
  var status = currentStatus_(r);
  var stage  = stageOf_(r);
  var t      = TRANSITIONS[status];
  var locked = (stage === 'SETTLED' || stage === 'CLOSED');

  var editable = [];
  if (!locked || who.isSuper) {
    Object.keys(FIELDS).forEach(function (k) {
      if (hasAny_(who.roles, FIELDS[k].roles)) editable.push(k);
    });
  }

  var reqEmail = String(r[COL.REQUESTER_EMAIL - 1] || '');
  return {
    row:            row,
    id:             String(r[COL.REQ_ID - 1] || ''),
    requestDate:    iso_(r[COL.REQUEST_DATE - 1]),
    requestedBy:    String(r[COL.REQUESTED_BY - 1] || ''),
    requesterEmail: reqEmail.indexOf('@') > 0 ? reqEmail : '',
    txnId:          String(r[COL.TXN_ID - 1] || ''),
    txnDate:        r[COL.TXN_DATE - 1] instanceof Date ? iso_(r[COL.TXN_DATE - 1]) : String(r[COL.TXN_DATE - 1] || ''),
    ticket:         String(r[COL.TICKET_NUMBER - 1] || ''),
    pax:            String(r[COL.PAX_NAME - 1] || ''),
    client:         String(r[COL.CLIENT - 1] || ''),
    airline:        String(r[COL.AIRLINE - 1] || ''),
    refundType:     String(r[COL.REFUND_TYPE - 1] || ''),
    issuedFrom:     String(r[COL.ISSUED_FROM - 1] || ''),
    amountAdvice:   toNumber_(r[COL.AMOUNT_ADVICE - 1]),
    comment:        String(r[COL.COMMENT - 1] || ''),
    customerEmail:  String(r[COL.CUSTOMER_EMAIL - 1] || ''),
    status:         status,
    stage:          stage,
    disbursement:   String(r[COL.DISBURSEMENT - 1] || ''),
    processingAt:   iso_(r[COL.NOW_PROCESSING - 1]),
    processDoneAt:  iso_(r[COL.PROCESSING_DONE - 1]),
    serviceAt:      iso_(r[COL.SERVICE_CHARGE_AT - 1]),
    postedAt:       iso_(r[COL.PAYOUT_COMPLETED - 1]),
    bspRa:          String(r[COL.BSP_RA_NO - 1] || ''),
    remark:         String(r[COL.MANAGER_REMARK - 1] || ''),
    currency:       String(r[COL.CURRENCY_TYPE - 1] || ''),
    approved:       toNumber_(r[COL.APPROVED_NGN - 1]),
    profit:         who.seeAll ? toNumber_(r[COL.PROFIT - 1]) : null,   // hidden from TCs
    modifiedBy:     String(r[COL.MODIFIED_BY - 1] || ''),
    lastModified:   iso_(r[COL.LAST_MODIFIED - 1]),
    settledAt:      iso_(r[WCOL.SETTLED_AT - 1]),
    settledBy:      String(r[WCOL.SETTLED_BY - 1] || ''),
    settlementType: String(r[WCOL.SETTLEMENT_TYPE - 1] || ''),
    settlementRef:  String(r[WCOL.SETTLEMENT_REF - 1] || ''),
    mine:           !!mine,
    nextStatus:     t ? t.next : '',
    canAdvance:     !!(t && hasAny_(who.roles, t.roles)),
    canSettle:      stage === 'POSTED' && hasAny_(who.roles, ['FINANCE']),
    revertTargets:  who.isSuper ? revertTargets_(r) : [],
    editable:       editable
  };
}


// ============================================================
//  PORTAL API  —  called from Portal.html via google.script.run
// ============================================================

function getSession() {
  var who = auth_();
  ensureSettlementColumns_();
  auditSheet_();    // create the support tabs up front, not on first write
  rolesSheet_();

  var fieldMeta = {};
  Object.keys(FIELDS).forEach(function (k) {
    fieldMeta[k] = { label: FIELDS[k].label, type: FIELDS[k].type };
  });

  return {
    email:             who.email,
    name:              staffNameFor_(who.email),
    roles:             who.roles,
    isSuper:           who.isSuper,
    roleLabels:        (who.isSuper ? ['Super Admin'] : [])
                         .concat(who.roles.map(function (r) { return ROLE_LABELS[r] || r; })),
    seeAll:            who.seeAll,
    viewAll:           who.viewAll,
    appUrl:            ScriptApp.getService().getUrl(),
    trackerUrl:        who.seeAll ? TRACKER_URL : '',
    disbursementTypes: getDisbursementTypes(),
    settlementOutputs: SETTLEMENT_OUTPUTS,
    fields:            fieldMeta,
    transitions:       TRANSITIONS
  };
}

/**
 * One-time setup, run from the editor or the Refund Tools menu.
 * Creates the Audit Log and Portal Roles tabs and the settlement columns.
 */
function setupPortal() {
  ensureSettlementColumns_();
  auditSheet_();
  rolesSheet_();
  notify_('Refund Portal', 'Settlement columns AE–AH, the Audit Log tab and the '
        + 'Portal Roles tab are in place.\n\nPortal URL:\n' + ScriptApp.getService().getUrl());
}


// ============================================================
//  SHEET LOCK — make the portal the only way to move a refund
// ============================================================

const LOCK_TAG = '[Portal]';

/**
 * Protects the columns the portal owns so only the script owner can edit
 * them in the sheet. Everyone else changes status through the portal,
 * which runs as the owner and so can still write.
 *   O:P   status, disbursement
 *   Q:AA  stage dates, BSP RA, remark, currency, approved, profit
 *   AE:AH settlement
 * Run by the spreadsheet OWNER from Refund Tools > Lock Status Columns.
 */
function lockSheetToPortal() {
  var sheet = requireSheet_(REQUESTS(), 'Requests');
  ensureSettlementColumns_();
  var me   = Session.getEffectiveUser();
  var rows = sheet.getMaxRows() - 1;

  unlockSheetFromPortal_(sheet);

  [[COL.STATUS, 2, 'Status & disbursement'],
   [COL.NOW_PROCESSING, COL.PROFIT - COL.NOW_PROCESSING + 1, 'Stage dates & amounts'],
   [WCOL.SETTLED_AT, 4, 'Settlement']].forEach(function (spec) {
    var p = sheet.getRange(2, spec[0], rows, spec[1]).protect()
                 .setDescription(LOCK_TAG + ' ' + spec[2] + ' — update via the Refund Portal');
    p.addEditor(me);
    var others = p.getEditors().filter(function (u) { return u.getEmail() !== me.getEmail(); });
    if (others.length) p.removeEditors(others);
    if (p.canDomainEdit()) p.setDomainEdit(false);
  });

  notify_('Status Columns Locked',
    'Columns O:P, Q:AA and AE:AH on Requests can now only be edited by ' + me.getEmail()
    + '.\n\nEveryone else updates refunds through the portal:\n' + ScriptApp.getService().getUrl());
}

function unlockSheetFromPortal() {
  unlockSheetFromPortal_(requireSheet_(REQUESTS(), 'Requests'));
  notify_('Status Columns Unlocked', 'The portal locks on Requests were removed.');
}

function unlockSheetFromPortal_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (String(p.getDescription() || '').indexOf(LOCK_TAG) === 0) p.remove();
  });
}

function listRequests() {
  var who   = auth_();
  var sheet = ensureSettlementColumns_();
  var last  = sheet.getLastRow();
  if (last < 2) return { rows: [], generatedAt: iso_(new Date()) };

  var vals   = sheet.getRange(2, 1, last - 1, WCOL.LAST).getValues();
  var myName = staffNameFor_(who.email).toUpperCase();
  var out    = [];

  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (isBlankRow_(r)) continue;

    var reqEmail = String(r[COL.REQUESTER_EMAIL - 1] || '').trim().toLowerCase();
    var reqName  = String(r[COL.REQUESTED_BY - 1] || '').trim().toUpperCase();
    var mine = reqEmail === who.email || (myName && reqName === myName);
    if (!who.viewAll && !mine) continue;

    out.push(rowToObj_(r, i + 2, who, mine));
  }

  out.reverse();   // newest first
  return { rows: out, generatedAt: iso_(new Date()) };
}

/** Activity history for one request, from the Audit Log sheet. */
function getRequestHistory(reqId) {
  var who = auth_();
  var sh = sheet_('Audit Log');
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  var out = [];
  for (var i = v.length - 1; i >= 0; i--) {
    if (String(v[i][2]) !== String(reqId)) continue;
    out.push({ at: iso_(v[i][0]), user: String(v[i][1]), action: String(v[i][4]),
               from: String(v[i][5]), to: String(v[i][6]), details: String(v[i][7]) });
  }
  return out;
}

/**
 * Saves field edits and, when p.advance is true, moves the request to
 * its next status.
 *   p = { reqId, row, ticket, expectedStatus, fields: {key: value}, advance }
 */
function updateRequest(p) {
  var who  = auth_();
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  var result, notify = null;
  try {
    var sheet = ensureSettlementColumns_();
    var row   = locateRow_(sheet, p.reqId, p.row, p.ticket);
    var r     = sheet.getRange(row, 1, 1, WCOL.LAST).getValues()[0];
    var cur   = currentStatus_(r);
    var stage = stageOf_(r);
    var now   = new Date();

    if (p.expectedStatus && cur !== p.expectedStatus) {
      return { ok: false, message: 'This request was moved to ' + cur
             + ' while you had it open. Refresh and try again.' };
    }
    if ((stage === 'SETTLED' || stage === 'CLOSED') && !who.isSuper) {
      return { ok: false, message: 'This request is closed. Only the super admin can change it.' };
    }

    // ── Field edits ──
    var changes = [];
    var fields  = p.fields || {};
    Object.keys(fields).forEach(function (k) {
      var f = FIELDS[k];
      if (!f) return;
      if (!hasAny_(who.roles, f.roles)) throw new Error('You cannot edit ' + f.label + '.');

      var val = fields[k];
      if (f.type === 'money')  { val = toNumber_(val); if (val === null) val = ''; }
      else if (f.type === 'emails') val = cleanEmails_(val);
      else val = String(val || '').trim();

      var old = r[f.col - 1];
      if (String(old === null ? '' : old) === String(val)) return;

      var cell = sheet.getRange(row, f.col);
      cell.setValue(val);
      if (f.type === 'money') cell.setNumberFormat('#,##0.00');
      r[f.col - 1] = val;
      changes.push(f.label + ': ' + (old === '' ? '(blank)' : old) + ' → ' + (val === '' ? '(blank)' : val));
    });

    // ── Status move ──
    var newStatus = null;
    if (p.advance) {
      var t = TRANSITIONS[cur];
      if (!t) return { ok: false, message: 'Nothing follows ' + cur + ' in the pipeline.' };
      if (!hasAny_(who.roles, t.roles)) {
        return { ok: false, message: 'Moving a request from ' + cur + ' to ' + t.next
               + ' needs ' + t.roles.map(function (x) { return ROLE_LABELS[x]; }).join(' or ') + '.' };
      }
      var missing = (t.require || []).filter(function (k) {
        var v = r[FIELDS[k].col - 1];
        return v === '' || v === null;
      });
      if (missing.length) {
        return { ok: false, message: 'Fill in ' + missing.map(function (k) { return FIELDS[k].label; }).join(' and ')
               + ' before moving to ' + t.next + '.' };
      }

      newStatus = t.next;
      sheet.getRange(row, COL.STATUS).setValue(newStatus);
      stampStage_(sheet, row, newStatus, now);
    }

    if (!changes.length && !newStatus) {
      return { ok: true, message: 'Nothing to save.', row: rowToObj_(r, row, who, false) };
    }

    sheet.getRange(row, COL.MODIFIED_BY).setValue(who.email);
    sheet.getRange(row, COL.LAST_MODIFIED).setValue(now);
    SpreadsheetApp.flush();

    var reqId = r[COL.REQ_ID - 1];
    if (changes.length) audit_(who.email, reqId, row, 'EDIT', '', '', changes.join('; '));
    if (newStatus)      audit_(who.email, reqId, row, 'STATUS', cur, newStatus, '');

    var fresh = sheet.getRange(row, 1, 1, WCOL.LAST).getValues()[0];
    if (newStatus) notify = { row: row, data: fresh, status: newStatus };

    result = {
      ok: true,
      message: newStatus ? 'Moved to ' + newStatus + '. Alerts sent.' : 'Changes saved.',
      row: rowToObj_(fresh, row, who, false)
    };
  } finally {
    lock.releaseLock();
  }

  // Emails go out after the lock is released so a slow send never blocks others.
  if (notify) notifyStage_(notify.row, notify.data, notify.status);
  return result;
}

/**
 * Finance output step: confirms the money reached the client.
 *   p = { reqId, row, ticket, output, ref }
 */
function confirmSettlement(p) {
  var who = auth_();
  if (!hasAny_(who.roles, ['FINANCE'])) throw new Error('Only Finance can confirm settlement.');
  if (SETTLEMENT_OUTPUTS.indexOf(p.output) === -1) return { ok: false, message: 'Choose how the refund was settled.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  var fresh, row;
  try {
    var sheet = ensureSettlementColumns_();
    row = locateRow_(sheet, p.reqId, p.row, p.ticket);
    var r = sheet.getRange(row, 1, 1, WCOL.LAST).getValues()[0];

    if (stageOf_(r) !== 'POSTED') {
      return { ok: false, message: 'Only requests posted to Travcom and not yet settled can be confirmed.' };
    }

    var now = new Date();
    sheet.getRange(row, WCOL.SETTLED_AT, 1, 4)
         .setValues([[now, who.email, p.output, String(p.ref || '').trim()]]);
    sheet.getRange(row, WCOL.SETTLED_AT).setNumberFormat('dd/mm/yyyy hh:mm');
    sheet.getRange(row, COL.MODIFIED_BY).setValue(who.email);
    sheet.getRange(row, COL.LAST_MODIFIED).setValue(now);
    SpreadsheetApp.flush();

    audit_(who.email, r[COL.REQ_ID - 1], row, 'SETTLED', 'POSTED TO TRAVCOM', p.output, String(p.ref || ''));
    fresh = sheet.getRange(row, 1, 1, WCOL.LAST).getValues()[0];
  } finally {
    lock.releaseLock();
  }

  sendSettledEmail_(fresh);
  return { ok: true, message: 'Settlement confirmed. Alerts sent.', row: rowToObj_(fresh, row, who, false) };
}


// ============================================================
//  REVERT  —  super admin only
// ============================================================

/** Earlier stages a row can be sent back to, plus undoing a settlement. */
function revertTargets_(r) {
  var cur = currentStatus_(r);
  var idx = PIPE_ORDER.indexOf(cur === 'DONE' ? 'POSTED TO TRAVCOM' : cur);
  var out = PIPE_ORDER.slice(0, Math.max(idx, 0));
  if (r[WCOL.SETTLED_AT - 1]) out.push('UNSETTLE');   // back to POSTED, settlement cleared
  return out;
}

/**
 * Moves a request back to an earlier stage.
 *   p = { reqId, row, ticket, expectedStatus, toStatus, reason }
 *   toStatus 'UNSETTLE' clears the Finance settlement and leaves it POSTED.
 *
 * Stage dates after the target are cleared so they are stamped afresh
 * when the request moves forward again, which keeps the turnaround
 * figures honest. Every cleared value is written to the Audit Log.
 */
function revertRequest(p) {
  var who = auth_();
  if (!who.isSuper) throw new Error('Only the super admin can revert a refund.');

  var reason = String(p.reason || '').trim();
  if (!reason) return { ok: false, message: 'Give a reason for the reversal.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  var fresh, row, cur, to;
  try {
    var sheet = ensureSettlementColumns_();
    row = locateRow_(sheet, p.reqId, p.row, p.ticket);
    var r = sheet.getRange(row, 1, 1, WCOL.LAST).getValues()[0];
    cur = currentStatus_(r);

    if (p.expectedStatus && cur !== p.expectedStatus) {
      return { ok: false, message: 'This request was moved to ' + cur
             + ' while you had it open. Refresh and try again.' };
    }
    if (revertTargets_(r).indexOf(p.toStatus) === -1) {
      return { ok: false, message: 'A request at ' + cur + ' cannot be reverted to ' + p.toStatus + '.' };
    }

    to = p.toStatus === 'UNSETTLE' ? 'POSTED TO TRAVCOM' : p.toStatus;
    var toIdx = PIPE_ORDER.indexOf(to);
    var hdr = sheet.getRange(1, 1, 1, WCOL.LAST).getDisplayValues()[0];
    var cleared = [];

    // Stage dates belonging to stages after the target
    [['PROCESSING',        [COL.NOW_PROCESSING]],
     ['PROCESS DONE',      [COL.PROCESSING_DONE]],
     ['SERVICE CHARGE',    [COL.SERVICE_CHARGE_AT]],
     ['POSTED TO TRAVCOM', [COL.PAYOUT_COMPLETED, COL.ISSUED_DATE]]
    ].forEach(function (s) {
      if (PIPE_ORDER.indexOf(s[0]) <= toIdx) return;
      s[1].forEach(function (c) {
        var v = r[c - 1];
        if (v === '' || v === null) return;
        cleared.push(hdr[c - 1] + ' = ' + (v instanceof Date ? fmtDate_(v) : v));
        sheet.getRange(row, c).clearContent();
      });
    });

    // Any settlement is undone by every reversal
    if (r[WCOL.SETTLED_AT - 1]) {
      cleared.push('Settlement = ' + r[WCOL.SETTLEMENT_TYPE - 1] + ' by ' + r[WCOL.SETTLED_BY - 1]
                 + (r[WCOL.SETTLEMENT_REF - 1] ? ' (ref ' + r[WCOL.SETTLEMENT_REF - 1] + ')' : ''));
      sheet.getRange(row, WCOL.SETTLED_AT, 1, 4).clearContent();
    }

    sheet.getRange(row, COL.STATUS).setValue(to);
    sheet.getRange(row, COL.MODIFIED_BY).setValue(who.email);
    sheet.getRange(row, COL.LAST_MODIFIED).setValue(new Date());
    SpreadsheetApp.flush();

    audit_(who.email, r[COL.REQ_ID - 1], row, 'REVERTED',
           cur + (r[WCOL.SETTLED_AT - 1] ? ' (settled)' : ''), to,
           'Reason: ' + reason + (cleared.length ? ' | Cleared: ' + cleared.join('; ') : ''));

    fresh = sheet.getRange(row, 1, 1, WCOL.LAST).getValues()[0];
  } finally {
    lock.releaseLock();
  }

  sendRevertEmail_(fresh, cur, p.toStatus === 'UNSETTLE' ? 'POSTED TO TRAVCOM (settlement undone)' : to,
                   reason, who.email);
  return { ok: true, message: 'Reverted to ' + to + '.', row: rowToObj_(fresh, row, who, false) };
}

/** Internal-only notice: reversals are never emailed to the client. */
function sendRevertEmail_(rowData, from, to, reason, by) {
  try {
    var requestId = rowData[COL.REQ_ID - 1] || '';
    var pax       = rowData[COL.PAX_NAME - 1] || '';
    var logo      = getLogoBlob_();
    var body = '<p>The following refund has been <strong>moved back</strong> by the super admin. '
             + 'Stage dates after the new status were cleared and will be stamped again as it moves forward.</p>'
             + detailRows_([
                 ['Request ID',    requestId],
                 ['Passenger',     pax],
                 ['Client',        rowData[COL.CLIENT - 1] || '—'],
                 ['From',          from],
                 ['To',            badge_(to, '#B71C1C')],
                 ['Reason',        reason],
                 ['Reverted By',   by],
                 ['Date',          fmtDate_(new Date())]
               ])
             + trackerLinkBlock_();

    var opts = {
      htmlBody: emailWrapper_('Refund Reverted', body, !!logo),
      cc:       EMAIL_CC,
      replyTo:  EMAIL_REPLY_TO,
      name:     'WGS Refund Operations'
    };
    if (logo) opts.inlineImages = { wgslogo: logo };

    var reqEmail = String(rowData[COL.REQUESTER_EMAIL - 1] || '');
    var toList = EMAIL_TO_INTERNAL + (reqEmail.indexOf('@') > 0 ? ',' + reqEmail : '');
    GmailApp.sendEmail(toList, '↩ Refund Reverted: ' + requestId + ' | ' + pax, '', opts);
  } catch (e) {
    Logger.log('Revert alert failed: ' + e.message);
  }
}


// ============================================================
//  STAGE TIMESTAMPS & ALERTS
//  Same rules as onEditHandler in Code.gs. A script write does not fire
//  the installable onEdit trigger, so the portal has to do this itself.
// ============================================================

function stampStage_(sheet, row, status, now) {
  var map = {
    'PROCESSING':        [COL.NOW_PROCESSING],
    'PROCESS DONE':      [COL.PROCESSING_DONE],
    'SERVICE CHARGE':    [COL.SERVICE_CHARGE_AT],
    'POSTED TO TRAVCOM': [COL.PAYOUT_COMPLETED, COL.ISSUED_DATE]
  };
  (map[status] || []).forEach(function (c) {
    var cell = sheet.getRange(row, c);
    if (!cell.getValue()) cell.setValue(now);
  });
}

function notifyStage_(row, rowData, status) {
  var reqEmail = String(rowData[COL.REQUESTER_EMAIL - 1] || '');
  if (reqEmail.indexOf('@') < 1) reqEmail = '';
  try {
    if (status === 'PROCESSING')        sendProcessingEmail_(row, rowData, reqEmail);
    if (status === 'PROCESS DONE')      sendProcessingDoneEmail_(row, rowData, reqEmail);
    if (status === 'SERVICE CHARGE')    sendServiceChargeEmail_(row, rowData, reqEmail);
    if (status === 'POSTED TO TRAVCOM') sendPostedEmail_(row, rowData, reqEmail);
  } catch (e) {
    Logger.log('Stage alert failed for row ' + row + ': ' + e.message);
  }
}

function sendSettledEmail_(rowData) {
  try {
    var requestId = rowData[COL.REQ_ID - 1] || '';
    var pax       = rowData[COL.PAX_NAME - 1] || '';
    var output    = String(rowData[WCOL.SETTLEMENT_TYPE - 1] || '');
    var sb        = badge_('SETTLED', '#003049');
    var subject   = '💸 Refund Settled: ' + requestId + ' | ' + pax;
    var reqEmail  = String(rowData[COL.REQUESTER_EMAIL - 1] || '');
    if (reqEmail.indexOf('@') < 1) reqEmail = '';

    var line;
    if (output.indexOf('PERSONAL ACCOUNT') >= 0) line = 'The refund has been paid to the client’s personal account.';
    else if (output.indexOf('LEDGER') >= 0)      line = 'The refund has been posted to the client ledger.';
    else                                         line = 'The refund has been applied against the outstanding client balance.';

    var custPairs = [
      ['Request ID',     requestId],
      ['Passenger',      pax],
      ['Ticket Number',  rowData[COL.TICKET_NUMBER - 1] || '—'],
      ['Refund Type',    rowData[COL.REFUND_TYPE - 1]   || '—'],
      ['Settlement',     output],
      ['Status',         sb],
      ['Settled On',     fmtDate_(rowData[WCOL.SETTLED_AT - 1])]
    ];
    var fullPairs = custPairs.concat([
      ['Client',          rowData[COL.CLIENT - 1] || '—'],
      ['Approved Refund', fmtNaira_(rowData[COL.APPROVED_NGN - 1])],
      ['Settlement Ref',  rowData[WCOL.SETTLEMENT_REF - 1] || '—'],
      ['Confirmed By',    rowData[WCOL.SETTLED_BY - 1] || '—']
    ]).concat(turnaroundPairs_(rowData));

    sendSplitEmail_(subject, fullPairs, custPairs, 'Refund Settled',
      '<p>' + line + '</p>',
      '<p style="margin-top:16px;color:#555;">This refund is now closed.</p>',
      rowData[COL.CUSTOMER_EMAIL - 1] || '', reqEmail);
  } catch (e) {
    Logger.log('Settlement alert failed: ' + e.message);
  }
}
