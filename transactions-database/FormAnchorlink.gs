// ============================================================
//  REFUND FORM LAUNCHER
//  FormAnchorlink.gs
//  Paste into: WGS - TRANSACTIONS DATABASE
//  (Extensions > Apps Script)
//
//  Set REFUND STATUS (column AP) to "START REFUND PROCESS"
//  and the Refund Request Form opens in a new browser tab,
//  pre-searched for that row's transaction ID.
// ============================================================

// The Refund Portal web app. "?page=form" opens the request form;
// the bare URL opens the portal itself.
const FORM_URL = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';

const WATCH_COL     = 42;                       // AP = REFUND STATUS
const TRIGGER_VALUE = 'START REFUND PROCESS';

// Link opened by openForm_(); set per edit so it carries the TXN ID.
var FORM_LINK = FORM_URL + '?page=form';


/**
 * Fires when REFUND STATUS is set to START REFUND PROCESS.
 * Install once via: Run > installLauncher
 */
function onEditLauncher(e) {
  if (!e || !e.range) return;
  if (e.range.getColumn() !== WATCH_COL || e.range.getRow() < 2) return;
  if (String(e.range.getValue() || '').trim().toUpperCase() !== TRIGGER_VALUE) return;

  // Pass the row's transaction ID so the form opens already searched for it
  var sh  = e.range.getSheet();
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
              .map(function (h) { return String(h).trim().toUpperCase(); });
  var c   = hdr.indexOf('TRANSACTION ID') + 1;
  var txn = c ? String(sh.getRange(e.range.getRow(), c).getDisplayValue()).trim() : '';

  FORM_LINK = FORM_URL + '?page=form' + (txn ? '&txn=' + encodeURIComponent(txn) : '');
  openForm_();
}


/** Opens the form in a new tab via a small dialog. */
function openForm_() {
  var html =
    '<style>'
    + 'body{font-family:Arial,sans-serif;text-align:center;padding:24px 20px;margin:0;}'
    + 'a{display:block;padding:14px;background:#F58220;color:#fff;text-decoration:none;'
    +   'border-radius:6px;font-weight:bold;font-size:14px;}'
    + 'a:hover{background:#D4711A;}'
    + 'p{color:#777;font-size:12px;margin:14px 0 0;}'
    + '</style>'

    + '<a href="' + FORM_LINK + '" target="_blank" '
    +   'onclick="setTimeout(function(){google.script.host.close()},400)">'
    +   'Open Refund Request Form</a>'
    + '<p>Opening in a new tab...</p>'

    + '<script>'
    + 'var w=window.open("' + FORM_LINK + '","_blank");'
    + 'if(w){setTimeout(function(){google.script.host.close()},500);}'
    + '<\/script>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(340).setHeight(160),
    'Refund Request'
  );
}


// ── Setup ────────────────────────────────────────────────────

/** Run this ONCE to activate. */
function installLauncher() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEditLauncher') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('onEditLauncher')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert('Done. Selecting "' + TRIGGER_VALUE
    + '" in column AP will now open the refund form.');
}

// ============================================================
//  MENUS
//
//  The Refund menu is NOT built here. Apps Script allows only one
//  onOpen() per project, and declaring one in this file silently
//  disabled the menus owned by other scripts.
//
//  All menus for this project are built by Menu.gs, which calls
//  openForm_() and installLauncher() below. Do not add an onOpen()
//  to this file.
// ============================================================
