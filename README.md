# BD Meetings → Saleoo Migration Spec: Email & Notification Logic
[Main BD MEETINGS SPREADSHEET] (https://docs.google.com/spreadsheets/d/1uicpBruuFeno2ES4hNw-TIAwNkGEI37gw8Z-A4yMpC8/edit?usp=sharing)
## Purpose of this document

This is **not** an integration guide. The Google Sheet + Apps Script system ("BD Meetings") is being retired. Saleoo is becoming the system of record for leads, replacing the Sheet entirely. The Sheet may stick around afterward only as a read-only export/archive — it won't drive any logic anymore.

This document exists so the Saleoo developer can **rebuild the same email-notification behavior natively in Saleoo**, since that behavior currently only exists buried in Apps Script (`code.js`) and was never written down anywhere else. Nothing here needs to be "integrated" with the old system — it needs to be **reimplemented** inside Saleoo, then the old Sheet/script gets switched off.

Calendar invites are explicitly **out of scope for now** — no auto-calendar-invite feature needs to be rebuilt at this stage.

## 1. The core concept: notifications are addressed by "Opener," not by lead

Every lead row has an **Opener** — the rep who owns that lead. All notification emails go to the Opener, not to the lead's own email address (the lead's email is only used for duplicate-detection, never as a send target in this system).

Today, the Opener → email mapping is a hardcoded lookup table:

```
Ben     → ben.arthur.wiz@gmail.com
Jane    → kaity.james.wiz@gmail.com
Jimmy   → jimmy.pearson.wiz@gmail.com
Selene  → selene.myles.wiz@gmail.com
Jasmine → jasmine.green.wiz@gmail.com
```

**What needs to change:** these can no longer be personal Gmail addresses. Saleoo will send via either an `@prime-ad` domain address or through Cloudflare/Supabase's email-sending service — the developer will decide which sending mechanism to use. But the *mapping concept* must be preserved: **whatever lead is assigned to whatever Opener/rep determines which address the notification email is sent to (and likely sent from)**. This Opener-keyed lookup is the one piece of logic that has to survive the migration unchanged — only the underlying addresses and transport mechanism change.

If Saleoo already has its own concept of "assigned rep" or "owner" per lead, that field should replace "Opener" directly — same role, new name.

## 2. When an email actually fires today (the trigger logic to replicate)

This is the part that's easy to get wrong because it's spread across several functions. There are exactly two triggers for sending a notification email in the current system:

### Trigger A — Notes field updated on a lead

When the "Notes" field on a lead is edited:
1. A timestamp ("Last Call") is recorded on that lead.
2. The update is queued (not sent immediately).
3. **Throttle:** if a notification was already sent for that same lead within the last **5 minutes**, the new one is silently dropped — no email, no error, no retry. This prevents spam when someone edits notes repeatedly in a short window.
4. If not throttled, a batch job picks up the queued update (today this runs every ~10 minutes; doesn't have to be a batch in Saleoo, can be immediate) and sends an email to that lead's Opener containing: company name, contact person, phone, lead email, and the notes content.
5. **Auto-CC rule:** if the word "contract", "prices", or "ben" (case-insensitive) appears anywhere in the company name or notes text, an admin address is automatically CC'd on top of the Opener. Today that's `ben.arthur.wiz@gmail.com` — this should become whatever the equivalent admin/owner address is in the new setup.

### Trigger B — Lead status changes (column "Move Trigger" in the old sheet)

When a lead's status changes to one of a fixed set of values, the lead is reclassified, but **no email is sent purely for the status change itself in the current system** — the move is silent. Only Trigger A (notes) sends email. If Saleoo wants status-change notifications too, that's a *new* feature, not something to port over — flag it as a decision point, don't assume it's required.

The status values that exist today (for reference, in case Saleoo's pipeline stages need to map to these):

| Status value | Maps to pipeline stage |
|---|---|
| Onboarded | Onboarded |
| Follow Up | Follow Ups |
| Meeting Attended | Follow Ups |
| No - Show / Callback | No-Show |
| NI | Dead Leads |
| DNC | Dead Leads |
| No Medicare | Dead Leads |
| Contract Sent | Contract Sent |
| Pending Medicare | Temporary Inactive |

## 3. Duplicate detection logic to preserve

Before any lead is reclassified/moved, the old system checks for duplicates using a composite key: **company name + phone + email must ALL match exactly** against another existing lead. Partial matches (e.g. same company, different phone) are intentionally NOT flagged as duplicates. If Saleoo already has its own duplicate detection, just confirm it uses the same all-three-fields-must-match rule rather than fuzzy/partial matching — that was a deliberate choice in the original system, not an oversight.

## 4. What's explicitly NOT being carried over right now

- **No calendar invites.** The old system could create a Google Calendar event with an invite when a checkbox was ticked. This is being dropped for now — do not build this unless asked again later.
- **No lead-facing emails.** The lead never receives anything automatically in the current system — only internal Opener/admin notifications. Don't add lead-facing emails unless that's a separate, explicit ask.
- **No read-back into the old sheet.** This isn't a two-way sync. Saleoo becomes the source of truth; the Sheet (if kept at all) is just a historical export, not something Saleoo needs to write to or read from going forward.

## 5. Open items for the Saleoo developer to decide/figure out

These were intentionally left open and don't need to be locked down before starting:

- Whether sending goes through classic SMTP on `@prime-ad`, or through Cloudflare/Supabase's email-sending service instead.
- Whether each Opener gets a distinct sending identity or there's a shared sender with a "from name" set per Opener — not yet decided on our end either.
- Whether the 5-minute throttle and ~10-minute batching are reproduced as-is, or just replaced with sane equivalents (e.g. a per-lead debounce) — the old timing was a workaround for Apps Script's trigger limitations, not a deliberate business requirement, so it doesn't need to be copied exactly.

## 6. Quick checklist for "have we hit feature parity on email logic"

- [ ] Each lead has an Opener/owner field that resolves to a real sending/receiving address (no longer Gmail)
- [ ] Editing a lead's notes triggers a notification to that lead's Opener
- [ ] A reasonable throttle/debounce exists so rapid edits don't spam multiple emails
- [ ] Notification includes company, contact, phone, lead email, and notes content
- [ ] Admin gets auto-CC'd when "contract," "prices," or "ben" appears in company/notes (or updated equivalent keyword/address)
- [ ] Duplicate check uses exact match on all three of company + phone + email before blocking a status move
- [ ] No calendar invite logic is being built right now
- [ ] No emails are sent to the lead themselves
