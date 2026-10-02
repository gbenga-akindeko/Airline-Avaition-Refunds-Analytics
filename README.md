# WGS Refund Management System

**WGS · Refund Operations**

WGS's refund tracking runs as a web portal on Google Apps Script. It takes a
refund from the travel consultant's first request, through each approval stage, to
Finance confirming the money reached the client. Every step has an owner, a timestamp,
an email alert and an audit trail. A Power BI model sits on the same data for analytics.

## The problem it solves

Refunds used to move through a shared spreadsheet. Anyone with edit access could set
any status. Stages got skipped, there was no record of who changed what, and alerts
went to a mailbox that didn't exist (`refunds@` instead of `refund@`). Finance had no
formal point at which to confirm that a refund had actually been paid, and management
had no reliable view of turnaround times.

## How it works

```
TC submits form ──▶ PENDING ──▶ PROCESSING ──▶ PROCESS DONE ──▶ SERVICE CHARGE ──▶ POSTED TO TRAVCOM ──▶ SETTLED
   (steps 1–5)      Refund Manager   Refund Manager      OPM / HOB        OPM / HOB / Finance      Finance
```

| Step | What happens | Who |
|---|---|---|
| 1–5 | Consultant picks the transaction (searches ~10,700 live and 2025 archive tickets), adds refund type, amount advised and customer emails. A reference `RFQ-YYYYMM-nnnnn` is issued. | Travel Consultant |
| 6 → 7 | Refund picked up | Refund Manager |
| 7 → 8 | Airline refund confirmed; disbursement type and approved amount required | Refund Manager |
| 8 → 9 | WGS service charge deducted | OPM / HOB |
| 9 → 10 | Posted to Travcom (ledger) | OPM / HOB / Finance |
| Output | Paid to the client's personal account, posted to the client ledger, or applied against debt | Finance |

Each move does four things:
- stamps the stage date
- emails the refund team, the management CC list, the consultant and the client (the client's copy leaves out internal figures and the tracker link)
- writes a line to the **Audit Log**
- checks that nobody else moved the refund in the meantime

## Access model

| Role | Can do |
|---|---|
| **Super Admin** (`dataanalyst@`) | Everything, including **reverting** a refund to an earlier stage, undoing a settlement, and editing closed requests. Set in code only. |
| Process Owner (`refund@`) | All forward moves |
| Refund Manager | Pending → Processing → Process done |
| OPM / HOB | Process done → Service charge → Posted |
| Finance | Service charge → Posted; confirm settlement |
| Oversight (CFO, Head RAC, Team Lead) | Read-only, plus the dashboard |
| Travel Consultant | Submit requests; read-only view of requests; no profit figures |

Roles come from the code plus the **Portal Roles** tab, so the Process Owner can grant
access without touching code. Google Groups are supported. Optionally, the status
columns can be locked so the sheet can only be changed through the portal.

## Components

| Piece | Where | Purpose |
|---|---|---|
| `Code.gs` | WGS Refund Management System (Apps Script) | Form handling, transaction search with chunked cache, email engine, sheet onEdit alerts, maintenance tools |
| `WebApp.gs` | same project | Portal API: routing, roles, pipeline rules, settlement, revert, audit, sheet lock |
| `Index.html` | same project | Refund request form (`…/exec?page=form`) |
| `Portal.html` | same project | Portal (`…/exec`): action queue with status dropdown, pipeline cards, request details, dashboard |
| `FormAnchorlink.gs` | WGS – Transactions Database | Opens the form when REFUND STATUS is set to *START REFUND PROCESS* |
| Google Sheet | WGS Refund Tracker | System of record: `Requests`, `Audit Log`, `Portal Roles`, `REFUNDS`, `2025 TICKETS`, `Data Validation` |
| **Refunds Analytics** | Power BI Desktop / Service | Analytics model over `Requests` and `Audit Log` |

## Refunds Analytics (Power BI)

The Power BI model connects with the **Google Sheets connector**, reading the same tracker
the portal writes to. There's no export step: refresh and you have the latest data.

**Tables**

| Table | Contents |
|---|---|
| `Requests` | One row per refund. Blank rows removed, money columns numeric, legacy statuses normalised (`PROCESSING DONE`, `OUT`, `DONE` …). Derived columns: `Status`, `Pipeline Status`, `Stage` (incl. SETTLED / CLOSED), `Stage 1–4 Days`, `Total TAT Days`, `Settlement Days`, `Is Open`, `Open Age Days`, `Request Day`. |
| `Audit Log` | Every portal action, including reversals |
| `Date` | Calendar from 2024, marked as the date table, related to `Requests[Request Day]` |
| `Refund Measures` | 58 measures in folders: **Volume**, **Financial**, **Turnaround** (avg / completed / fastest / slowest / outstanding / % per stage), **Funnel**, **Audit** |

**Refresh in the Power BI Service**
1. Publish from Power BI Desktop to the workspace.
2. Go to dataset **Settings → Data source credentials → Google Sheets → Edit credentials**, choose OAuth2, and sign in with a WGS account that can open the tracker.
3. Under **Scheduled refresh**, add e.g. 08:00 / 12:00 / 17:00. No gateway is needed.

**Data notes**
- 93 legacy rows have no REQUEST DATE in the sheet. They count in all-time totals but
  drop out when a Year or Month filter is applied, the same as in the sheet dashboard.
- Approved Refund stays at ₦0 until refund managers fill that column.
- Funnel measures count a stage as reached if it has a timestamp *or* a later status.
  Legacy rows with status but no dates therefore show as completed in the funnel, but
  not in the turnaround measures.

## Repository layout

```
apps-script/            Code.gs, WebApp.gs, Index.html, Portal.html  (one Apps Script project)
transactions-database/  FormAnchorlink.gs  (bound to the Transactions Database sheet)
powerbi/                Page-background templates (.pptx + PNGs) and the Power BI colour theme
```

## Configuration placeholders

Company identifiers have been replaced so the code can be shared publicly. Set these before deploying:

| Placeholder | Where | Replace with |
|---|---|---|
| `YOUR_SPREADSHEET_ID` | `apps-script/Code.gs` (`TRACKER_ID`, `TRACKER_URL`) | ID of the tracker Google Sheet |
| `YOUR_DEPLOYMENT_ID` | `transactions-database/FormAnchorlink.gs` (`FORM_URL`) | The web app deployment ID (`…/macros/s/<id>/exec`) |
| `YOUR_LOGO_FILE_ID` | `apps-script/Code.gs` (`LOGO_FILE_ID`) | Drive file ID of the logo |
| `*@example.com`, `*@holdings.example.com` | `Code.gs` email lists, `WebApp.gs` `ROLE_MEMBERS` / `SUPER_ADMINS` / `STAFF_DOMAINS` | Your real mailboxes and Workspace domains |

## Deployment

See [DEPLOY.md](DEPLOY.md).
