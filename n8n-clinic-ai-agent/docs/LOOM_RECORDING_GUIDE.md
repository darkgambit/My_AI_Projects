# 🎬 Loom Recording Guide — film the REAL system (strongest proof)

A 90-second live video beats any screenshot. Recruiters want to see: **real message in → real calendar event created → real reply back.**
Total prep + filming: ~15 minutes.

## ⚠️ Pre-flight safety checklist (do this FIRST)

- [ ] Use a **test patient** (your own phone number, name "Test Patient") — never show real patient names/numbers on camera.
- [ ] **Never open any node's Credentials panel** on camera. If you must open a node, check the API-key/token field is not visible.
- [ ] Close tabs showing your n8n URL's settings/API page, hosting panel, email.
- [ ] Delete old test executions that contain personal data (n8n → Executions → select → Delete), or just don't open them.
- [ ] Browser zoom 90–100%, hide bookmarks bar (`Ctrl+Shift+B`).
- [ ] Set your Google Calendar + Bookings Sheet to show only test rows in frame (scroll/zoom).

## Setup (2 minutes)

1. Install **Loom** (Chrome extension, free) → sign in.
2. Open in tabs, left→right: **① WhatsApp Web** (or Messenger) · **② n8n – both workflows + Executions list** · **③ Google Calendar** · **④ Bookings Google Sheet**.
3. Alternative/backup: press record on the included [`../demo/index.html`](../demo/index.html) (choose *Slow (for recording)* speed) — good B-roll while you narrate, but the live tabs above are the real proof. Show both if you like: demo first (explains), live second (proves).

## Shot list + word-for-word script (~90 seconds)

> Speak slowly. Pause 1 second between shots — Loom lets you trim later.

**[0:00 – WhatsApp tab]**
> "Hi, I'm [YOUR NAME]. I built an AI appointment agent for a clinic using n8n. Patients just chat on WhatsApp or Messenger — watch. I'll ask for a cleaning tomorrow morning."

*(Type and send: `Hi, do you have a free slot tomorrow morning for a cleaning?` — let the reply arrive on camera.)*

**[0:20 – reply arrives]**
> "The agent checked the clinic's real Google Calendar — that's why it offered me only the actually-free times. I'll take 11 o'clock."

*(Send: `11:00 please, name: Test Patient` → wait for ✅ confirmation.)*

**[0:35 – n8n Executions tab]**
> "Behind the scenes: two workflows. The first is the Clinic AI Agent — webhooks for WhatsApp and Messenger, an LLM agent with conversation memory, and six booking tools. The second workflow is the booking engine those tools call."

*(Click into the latest successful execution; let viewers see the green path light up. Do NOT open credential panels.)*

**[0:55 – Google Calendar tab]**
> "Here's the event it just created…"

**[1:05 – Google Sheets tab]**
> "…and the matching row in the bookings log. Calendar and spreadsheet are always updated together, and the slot is re-validated at the moment of booking — so double-booking is impossible."

*(Bonus if time: send `I need to move it to 14:00` and show the event move — this is the moment that impresses people.)*

**[1:20 – back to n8n canvas, zoomed out]**
> "It also handles cancellations, 'show my bookings', and taken slots — offering the nearest alternative instead of failing. Full architecture and the workflow files are in the GitHub repo linked below. Thanks for watching!"

## Publish settings

- **Title:** `Clinic AI Agent — n8n WhatsApp/Messenger Appointment Automation (live demo)`
- **Description (paste):**
  ```
  Live demo of my n8n automation: patients book/reschedule/cancel clinic appointments
  over WhatsApp & Messenger. AI agent (OpenRouter LLM + memory) → Google Calendar
  FreeBusy validation → event + Google Sheets booking log, with double-booking protection.
  Full case study + workflows: [PASTE GITHUB LINK]
  ```
- Set visibility to **"Anyone with the link"**, allow comments off.
- Trim dead air at start/end; pick a thumbnail frame showing the ✅ booking confirmation.

## Then update the repo

1. Paste the Loom URL into `README.md` (the line marked `[ADD YOUR LOOM LINK HERE]`) and into `APPLICATION_COPY_PASTE.md`.
2. In job applications, give **both links**: Loom first (proof it's live), GitHub second (proof you can document/ship). Example format is in `APPLICATION_COPY_PASTE.md`.
