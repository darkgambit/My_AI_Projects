# 🎬 Video Proof WITHOUT Loom — 5 free ways

You do **not** need Loom. Pick ONE row — top rows need zero installs:

| # | Method | Install? | Best for |
|---|---|---|---|
| **0** | **Use the ready-made MP4** in this repo | ❌ nothing to film | Applying TODAY — upload & link it |
| **1** | **Windows Snipping Tool recorder** | ❌ built into Windows 11 | Filming your LIVE n8n (recommended proof) |
| **2** | **Recorder built into the demo page** | ❌ Chrome/Edge only | Narrated video of the demo, download instantly |
| **3** | **Clipchamp** | ❌ pre-installed on Windows | Screen + your face in a bubble, like Loom |
| **4** | **OBS Studio** | ✅ free download | Best quality/control if you'll record often |

---

## Option 0 — No filming at all (2 minutes)

This repo already contains a finished video tour of your workflows:

- 🎞️ `assets/workflow-demo.mp4` — 21 seconds, 1 MB, captions burned in (also plays as the GIF in the README)

Upload that file to YouTube as **Unlisted** (steps at the bottom) and paste the link in your application.
It's honest B-roll of your architecture — but a live recording (Option 1) is stronger proof, so do that when you have 15 minutes.

## Option 1 — Windows Snipping Tool (film the REAL system) ⭐

1. Press **`Win + Shift + R`** *(or open Start → "Snipping Tool" → click the 🎥 video-camera icon → **+ New**)*.
2. Drag a rectangle around your **browser window**.
3. In the small toolbar, check the **microphone** icon is ON if you'll narrate → click **Start**.
4. Do the 90-second walkthrough below (switch between WhatsApp Web → n8n Executions → Google Calendar → Sheets — all inside the recorded browser).
5. Click **Stop** → the video opens in Snipping Tool → **Save** (💾) as `clinic-ai-agent-live-demo.mp4`.
6. Upload to YouTube Unlisted (bottom of this page) → copy the link.

> Don't have Win+Shift+R? Use **Xbox Game Bar**: `Win + G` → Capture widget → ● record (or `Win + Alt + R`). It records one app window (your browser) — tab-switching still works. File lands in `Videos\Captures`.

## Option 2 — The demo page records itself (Chrome/Edge)

1. Open [`../demo/index.html`](../demo/index.html), pick speed **Slow (for recording)**.
2. Click **🔴 Record my screen — no Loom needed**, keep "include my microphone" ticked.
3. In the picker choose **This tab** / your browser window → **Share**.
4. Click a scenario (Book → Reschedule are the most impressive), narrate what each node does as it lights up.
5. **⏹ Stop recording** → preview appears → **⬇️ Download video (.webm)**.
6. Upload the `.webm` to YouTube exactly as-is (YouTube accepts WebM) → Unlisted link.

## Option 3 — Clipchamp (Loom-style face bubble, pre-installed)

Start → **Clipchamp** → *Create a new video* → **Record & create** → *Screen and camera* → record your browser → **Export → 1080p** (free). Upload the MP4 to YouTube.

## Option 4 — OBS Studio (free, pro)

[obsproject.com](https://obsproject.com) → install → **+ Source → Window Capture** (your browser) + your mic under Audio Mixer → **Start Recording** → **Stop** → file in `Videos`. Settings → Output → format `mp4`. Best if you'll make portfolio videos regularly.

---

## ⚠️ Before ANY live recording — safety checklist

- [ ] Test patient only (your own phone, "Test Patient") — no real patient names/numbers on screen
- [ ] **Never open a node's Credentials panel** or Settings → API on camera
- [ ] Close tabs with hosting panel / email / API keys; hide bookmarks bar (`Ctrl+Shift+B`)
- [ ] Zoom the Sheet/Calendar so only test rows are in frame

## 🎙️ The 90-second walkthrough (same for every option)

**[WhatsApp Web]** *"I'm [NAME]. I built an AI appointment agent for a clinic in n8n. A patient just chats on WhatsApp — watch."* → send: `Hi, do you have a free slot tomorrow morning for a cleaning?` → reply arrives → *"It checked the real Google Calendar, so it only offered actually-free times."* → send: `11:00 please, name: Test Patient` → ✅ arrives.

**[n8n Executions]** Open the green execution. *"Two workflows: the AI agent with memory and six tools, and the booking engine those tools call."* (Don't open credentials.)

**[Google Calendar]** *"Here's the event it created…"* **[Google Sheets]** *"…and the matching booking row. Both are always updated together, and the slot is re-validated at booking time, so double-booking is impossible. It also reschedules and cancels the same way. Full case study is on GitHub, linked below."*

## ☁️ Where to host the video (pick one, all free)

**A) YouTube — Unlisted ⭐ (best: plays instantly, link never expires)**
youtube.com → **+ Create → Upload video** → choose file → Title: `Clinic AI Agent — n8n WhatsApp/Messenger Appointment Automation (live demo)` → Audience: *Not made for kids* → Visibility: **Unlisted** → copy link. (Unlisted = only people with the link can watch; it never appears in search.)

**B) Google Drive** — upload → right-click → Share → *Anyone with the link* → copy link. (Works, but first-time viewers sometimes wait for processing.)

**C) Straight into the job form / LinkedIn / WhatsApp** — the 1 MB `workflow-demo.mp4` is small enough to attach or send directly.

## ✅ Finish the job

1. Paste the video link into `README.md` where it says `[ADD YOUR VIDEO LINK HERE]`.
2. Use the ready-made text in [`APPLICATION_COPY_PASTE.md`](APPLICATION_COPY_PASTE.md) — video link first (proof it's live), GitHub link second (proof you can ship & document).
