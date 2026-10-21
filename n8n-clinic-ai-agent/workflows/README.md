# 📥 Workflow JSON — export, SANITIZE, then commit

> ⚠️ **Nothing in this folder yet — on purpose.** Exporting from n8n can include credential IDs,
> webhook paths, Sheet/Calendar IDs and other identifiers. Never upload a raw export to a public repo.
> Follow the 5-minute checklist below, then drop the cleaned files here as:
> `1-clinic-ai-agent.json` and `2-booking-subworkflow.json`.

## Step 1 — Export from n8n

Open each workflow → top-right `⋯` menu → **Download**. (Or: Workflows list → `⋯` → Download.)
Do this for **both**: the *Clinic AI Agent* workflow and the *Booking sub-workflow*.

## Step 2 — Sanitize (open the JSON in any text editor)

Find-and-replace / delete **all** of these before committing:

- [ ] `credentials` blocks → keep the credential **name/type** if you like, delete `id` values
- [ ] Webhook `path` / `webhookId` values → replace with `REPLACE-ME`
- [ ] Google Sheet document IDs & Calendar IDs (long strings in node parameters) → `YOUR_SHEET_ID` / `YOUR_CALENDAR_ID`
- [ ] Any phone numbers, page IDs, access tokens, API keys pasted into Code/HTTP nodes → `REPLACE-ME`
- [ ] Hard-coded clinic name/address if the client didn't agree to be named → "Demo Clinic"
- [ ] Search the file for: `key`, `token`, `secret`, `Bearer`, `eyJ` — investigate every hit

Quick terminal check before committing (should print nothing):

```bash
grep -iE "bearer|api[_-]?key|secret|eyJhbGci" workflows/*.json
```

## Step 3 — Sheet structure a new user must recreate

Tabs in one Google Sheet (adjust to your exact column names):

| Tab | Columns (from the live build) |
|---|---|
| `ClinicInfo` | clinic name, opening hours per weekday, slot length, timezone |
| `Treatments` | treatment, duration (min), price |
| `Bookings` | booking id, patient name, channel, sender id, treatment, start, end, calendar event id, status (Confirmed / Cancelled / Rescheduled), created at |

## Step 4 — Import order (for reviewers / future you)

1. Import **`2-booking-subworkflow.json` first** and activate it — the agent's tools reference it.
2. Import `1-clinic-ai-agent.json`; open each of the 6 tool nodes and re-select the sub-workflow.
3. Re-connect credentials: OpenRouter · Google Calendar OAuth · Google Sheets OAuth · Meta (WhatsApp/Messenger).
4. Re-select your Sheet + Calendar in every Google node.
5. Copy both **Production** webhook URLs into the Meta app dashboard (with your verify token).
6. Activate both workflows → message the page/number → check n8n **Executions** for a green run.

## Node inventory (reconstructed from the production screenshot)

**Workflow 1 – Clinic AI Agent:** Messenger Webhook · IF → Messenger Verification · Messenger Ack · Parse Messenger Message ·
WhatsApp Webhook · WhatsApp Verification · WhatsApp Ack · Parse WhatsApp Message · Parse Message ·
Clinic AI Agent (OpenRouter Chat Model, Conversation Memory, 6 workflow tools) ·
Extract Reply · Which Channel? · Send Messenger Reply (HTTP → Graph API) · Send WhatsApp Reply

**Workflow 2 – Booking sub-workflow:** When Executed by Another Workflow · Route by Action ·
[info] Read Clinic Info → Read Treatments → Build Clinic Info ·
[availability] Read Clinic Info (Avail) → Read Day Bookings → Check Calendar FreeBusy → Compute Availability ·
[book] Read Clinic Info (Book) → Check Calendar Slot → Validate Book Slot → Slot Free? → Create Calendar Event → Prepare Booking Row → Append Booking Row → Booking Confirmed / Booking Slot Taken ·
[cancel] Find Bookings To Cancel → Pick Booking To Cancel → Cancel Booking Found? → Delete Calendar Event → Mark Booking Cancelled → Cancel Confirmed / Cancel Nothing Found ·
[my bookings] Find My Bookings → List My Bookings ·
[reschedule] Find Bookings To Move → Pick Booking To Move → Move Booking Found? → Read Clinic Info (Move) → Check New Slot → Validate New Slot → New Slot Free? → Update Calendar Event → Update Booking Row → Reschedule Confirmed / Reschedule Slot Taken / Reschedule Nothing Found
