# Integrating Saleoo with the BD Meetings Email/Calendar System

This document explains how the internal "BD Meetings" automation works and exactly what Saleoo needs to do to plug into it. Read this fully before writing any code — this is **not** a REST API. There is no endpoint to call. It is a Google Sheet with an Apps Script automation layer attached to it, and integration means writing rows into that sheet in the right shape.

## 1. What this system actually is

BD Meetings is a Google Apps Script project (`BDMEETINGS` repo) bound to a single Google Sheet. There is no server, no public URL, and no webhook receiver. All automation is triggered by:

- **`onEdit(e)`** — fires instantly when a cell in the sheet is edited (by a human or by the Sheets API)
- **Time-based triggers** — `processQueuedMovements` (every ~1 min), `processBatchEmails` (every ~10 min), `processBatchRowMovement` (every ~5 min, backup pass)

So "integrating the email system" means: **Saleoo needs to create/update rows in this Sheet with the correct columns populated**, and the existing script will pick them up, send emails via `MailApp.sendEmail()`, and create Calendar invites via `CalendarApp.createEvent()` automatically. Saleoo does not send emails directly — it triggers our automation by writing data, and our automation sends the email.

## 2. Two integration paths — pick one with us before building

| Path | How it works | Effort |
|---|---|---|
| **A. Google Sheets API (recommended)** | Saleoo uses a Google service account with edit access to the Sheet, and calls `spreadsheets.values.append` / `update` directly. | Low — no changes to our script needed. |
| **B. Apps Script Web App** | We deploy a small `doPost(e)` endpoint in this same Apps Script project that accepts a JSON payload from Saleoo and writes the row for it. | Medium — requires us to add and deploy a new endpoint (not present today). |

If Saleoo wants a real HTTP API instead of writing to Sheets directly, **tell us** — Path B needs a new deployment on our end, listed in section 6 below. The rest of this document assumes Path A unless we agree otherwise.

## 3. The sheet structure Saleoo must write to

The spreadsheet has multiple tabs ("stages"). Only these are active and processed by the script — anything else is ignored:

```
New Meetings, Follow Ups, No-Show, Contract Sent,
Invoice Sent, Dead Leads, Temporary Inactive, Onboarded
```

New leads from Saleoo should be appended as new rows in **"New Meetings"**.

### Column map (1-indexed, fixed — do not shift these)

| Col | Field | Notes |
|---|---|---|
| B | Opener | Must exactly match a name configured on our side (see §4). Determines who gets emailed. |
| C | Move Trigger | Dropdown status value (see §5). Leave blank on initial insert. |
| E | Company Name | Used in email subject/body and duplicate checks. |
| F | Authorized/Contact Person | |
| G | Phone | Used in duplicate detection. |
| H | Email | Lead's email — used in duplicate detection, **not** the recipient of internal notifications. |
| I | Meeting Date/Time | Required only if a meeting will be scheduled — must be a real future datetime, not text. |
| K | Notes | Editing this column is what fires the "Note Update" email to the Opener. |
| O | Schedule checkbox | `TRUE`/checked + a valid date in column I = triggers a Calendar invite when `scheduleSelectedMeetings()` runs. |

Columns A, D, J, L, M, N are either spacers or used internally (e.g. M = "Last Call" timestamp, auto-set by the script) — Saleoo should leave them blank.

**Row 1 is always the header row and must never be overwritten.**

## 4. Who gets emailed — the Opener map

Emails are addressed by **Opener name**, not by lead email. The mapping of Opener → internal email address lives in Apps Script's Script Properties on our side (not in the sheet, not visible to Saleoo):

```
Ben     → ben.arthur.wiz@gmail.com
Jane    → kaity.james.wiz@gmail.com
Jimmy   → jimmy.pearson.wiz@gmail.com
Selene  → selene.myles.wiz@gmail.com
Jasmine → jasmine.green.wiz@gmail.com
```

**Action for Saleoo:** whatever field in Saleoo represents the assigned rep/owner must be mapped to one of these exact names before being written to column B. If the name doesn't match exactly (case-sensitive-ish, trimmed), the script silently skips sending the email and just logs it — no error is raised back to Saleoo. If Saleoo needs new openers added, tell us and we'll add them to the map; this can't be self-served from the sheet.

## 5. Status values that drive row movement (column C)

These are the only values the system recognizes for moving a row between tabs (configured in the sheet's own "Settings" tab):

| Value Saleoo writes to Column C | Row moves to |
|---|---|
| `Onboarded` | Onboarded |
| `Follow Up` | Follow Ups |
| `Meeting Attended` | Follow Ups |
| `No - Show / Callback` | No-Show |
| `NI` | Dead Leads |
| `DNC` | Dead Leads |
| `No Medicare` | Dead Leads |
| `Contract Sent` | Contract Sent |
| `Pending Medicare` | Temporary Inactive |

Any other value in column C is ignored by the row-mover (no error, no movement). If Saleoo introduces new lead statuses, they need a corresponding row added to the Settings tab — talk to us before relying on a new value.

## 6. What actually sends an email (and what doesn't)

- Writing a new lead row → **no email**. This just gets it into the pipeline.
- Editing column K (Notes) on an existing row → triggers a "Note Update" email to that row's Opener, with company/contact/phone/email/notes in an HTML table. This is throttled to once every 5 minutes per row.
- Setting the column O checkbox to `TRUE` with a valid future date in column I → creates a Google Calendar event and **sends a calendar invite** (via `sendInvites: true`) to the Opener and a fixed always-included guest. This only runs when someone manually runs `scheduleSelectedMeetings()` from the Sheet's menu — it is not automatic on edit.
- If "contract", "prices", or "ben" appears in the company/notes text, the admin (`ben.arthur.wiz@gmail.com`) is auto-CC'd on the notification email.

**Important:** if Saleoo's integration writes to column K programmatically expecting an email to fire immediately, be aware of the 5-minute throttle and the fact that the actual send happens on the next `processBatchEmails` run (every ~10 minutes), not instantly.

## 7. Duplicate protection Saleoo should know about

Before a row is moved between tabs, the script checks company name + phone + email against other active tabs. **All three must match exactly** for it to count as a duplicate — partial matches are allowed through. This only runs on row *movement*, not on initial insert, so Saleoo can safely push the same lead multiple times into "New Meetings" without it being auto-blocked at that stage; dedup only kicks in once a status change tries to move it.

## 8. What we need from Saleoo before this can be wired up

1. Confirm which path (§2) Saleoo's engineer wants — direct Sheets API writes, or a hosted endpoint we build.
2. Confirm how Saleoo will determine the "Opener" value per lead, and confirm it can be mapped 1:1 to the five names in §4 (or tell us what new names/emails to add).
3. Confirm date/time format Saleoo will send for column I — must parse cleanly into a JS `Date`.
4. If using Path A: Saleoo shares the service account email they'll authenticate with, so we can share Edit access to the specific Sheet.
5. If using Path B: give us a few days — `doPost` isn't implemented yet and needs a new Apps Script deployment + URL handoff.

## 9. Things this system explicitly does NOT do (don't assume these)

- It does not send a "lead received" confirmation email to the lead itself — only internal notifications to the Opener.
- It does not expose any way for Saleoo to read data back out (no "get lead status" call). Status changes are one-directional, Saleoo → Sheet.
- It does not validate email addresses or phone formats before sending.
- It is timezone-bound to `Africa/Cairo` for all date/trigger logic — send datetimes accordingly or be explicit about timezone in your payload.
