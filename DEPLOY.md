# WGS Refund Portal: deployment

The portal adds a web front end to the tracker. The `Requests` sheet is still the
source of truth, and the sheet, the dashboard and the existing onEdit alerts all keep
working. Anyone can still edit the sheet directly while the team moves over.

## 1. Add the files (WGS Refund Management System project)

1. Create a script file called **WebApp** and paste in `WebApp.gs`.
2. Create an HTML file called **Portal** and paste in `Portal.html`.
3. In **Code.gs**, delete the whole `doGet(e)` function. WebApp.gs now provides
   `doGet` and routes requests (`?page=form` serves the request form, anything
   else serves the portal). If both files define `doGet`, one of them will be ignored.
4. Optional: add a menu item to `onOpen()` in Code.gs:
   `.addItem('Open Refund Portal', 'openPortalLink')`, then add this function:

   ```js
   function openPortalLink() {
     var url = ScriptApp.getService().getUrl();
     var html = HtmlService.createHtmlOutput('<a href="' + url + '" target="_blank">Open the Refund Portal</a>');
     SpreadsheetApp.getUi().showModalDialog(html.setWidth(300).setHeight(80), 'Refund Portal');
   }
   ```

## 2. Deploy

**Deploy → Manage deployments → edit the existing web app → Version: New version.**
Editing the existing deployment keeps the current `/exec` URL, so the launcher link and
any bookmarks still work.

- **Execute as:** Me (the owner). Staff don't need edit access to the sheet.
- **Who has access:** Anyone within example.com

Authorise the new scope when prompted. GroupsApp is used to check group membership.

| URL | Opens |
|---|---|
| `…/exec` | Portal (queues, pipeline, actions) |
| `…/exec?page=form` | Refund request form (steps 1–5) |
| `…/exec?page=form&txn=ABC123` | Form, pre-searched for that transaction |

## 3. Check before go-live

- **holdings.example.com accounts (CFO, Head RAC, Josephine).** If that domain is a
  *separate* Google Workspace from example.com, "Anyone within example.com" will
  block them, and `Session.getActiveUser()` returns blank for them anyway. You can fix
  this in one of two ways: add holdings.example.com as a secondary domain on the same
  Workspace, or give those three people example.com accounts. If both domains are
  already on one Workspace, nothing needs to change.
- **Shared mailboxes.** `refund@`, `refunds@` and `customersuccess@` get their roles
  whether they're real accounts or Google Groups. If they're groups, each member signs
  in as themselves and still gets the Refund Manager role, as long as the script owner
  can see the group's members.
- **`SETTLEMENT_TRACKING_FROM`** in WebApp.gs is set to 1 Oct 2026. Anything posted
  before that date shows as *Closed*, so historical rows don't flood Finance's queue.
  Change the date if you go live on a different day.
- Open the portal once as the owner. That run creates columns **AE–AH** (settlement)
  and the **Audit Log** tab.

## 4. Optional: open the form pre-filled from the Transactions Database

**Index.html:** add this at the end of the `DOMContentLoaded` handler:

```js
google.script.url.getLocation(function (loc) {
  var txn = loc.parameter && loc.parameter.txn;
  if (!txn) return;
  searchEl.value = txn;
  google.script.run.withSuccessHandler(function (res) {
    allTxns = (res && res.results) || [];
    if (allTxns.length === 1) selectTxn(0); else renderList(res);
  }).searchTransactions(txn, 40);
});
```

**FormAnchorlink.gs** (Transactions Database): replace `onEditLauncher` so it passes
the row's transaction ID:

```js
function onEditLauncher(e) {
  if (!e || !e.range) return;
  if (e.range.getColumn() !== WATCH_COL || e.range.getRow() < 2) return;
  if (String(e.range.getValue() || '').trim().toUpperCase() !== TRIGGER_VALUE) return;

  var sh  = e.range.getSheet();
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
              .map(function (h) { return String(h).trim().toUpperCase(); });
  var c   = hdr.indexOf('TRANSACTION ID') + 1;
  var txn = c ? String(sh.getRange(e.range.getRow(), c).getDisplayValue()).trim() : '';

  FORM_LINK = FORM_URL + '?page=form' + (txn ? '&txn=' + encodeURIComponent(txn) : '');
  openForm_();
}
var FORM_LINK = FORM_URL;
```

Then, inside `openForm_()`, replace both occurrences of `FORM_URL` with `FORM_LINK`.

## Who does what in the portal

| Flow step | Status move | Who |
|---|---|---|
| 1–5 | Submit (form) → PENDING | Any staff (TC) |
| 6 → 7 | PENDING → PROCESSING | Refund Manager |
| 7 → 8 | PROCESSING → PROCESS DONE (needs disbursement type + approved amount) | Refund Manager |
| 8 → 9 | PROCESS DONE → SERVICE CHARGE | OPM / HOB |
| 9 → 10 | SERVICE CHARGE → POSTED TO TRAVCOM | OPM / HOB / Finance |
| Output | Confirm settlement (personal account / client ledger / debt) | Finance |
| — | Read-only view of everything | CFO, Head RAC, Team Lead, Josephine |
| — | Can do everything, and can edit closed requests | Process Owner (refund@) |

Every move stamps the stage date, sends the existing stage email (internal + CC list,
requester, client) and writes a line to the Audit Log. Before saving, the portal checks
that nobody else has moved the request since it was loaded. If someone has, the save
is refused, which stops two people advancing the same refund twice.

To change who does what, edit `ROLE_MEMBERS`, `TRANSITIONS` and `FIELDS` at the top of
WebApp.gs.
