# ✍️ Copy-Paste Answers for Applications

## ✅ READY TO PASTE NOW (real links, no editing needed)

**If the form asks for ONE link — paste this (plays online directly in the browser):**
```
https://raw.githubusercontent.com/darkgambit/My_AI_Projects/clinic-ai-agent-demo-v2/n8n-clinic-ai-agent/assets/clinic-ai-agent-workflow-demo.mp4
```

**If you can add text/links — paste this:**
```
Watch the video online (66s, Full HD, captioned): https://raw.githubusercontent.com/darkgambit/My_AI_Projects/clinic-ai-agent-demo-v2/n8n-clinic-ai-agent/assets/clinic-ai-agent-workflow-demo.mp4
Video file / release: https://github.com/darkgambit/My_AI_Projects/releases/tag/clinic-ai-agent-demo-v2
Full case study + workflows: https://github.com/darkgambit/My_AI_Projects/tree/arena/315aeb7f-my-ai-projects/n8n-clinic-ai-agent
```

*(Optional prettier watch page: repo **Settings → Pages → Branch: `arena/315aeb7f-my-ai-projects`, folder `/ (root)` → Save** — then `https://darkgambit.github.io/My_AI_Projects/n8n-clinic-ai-agent/watch.html` goes live. After PR #1 is merged, the case-study link becomes `https://github.com/darkgambit/My_AI_Projects/tree/main/n8n-clinic-ai-agent` — the video/release links never change.)*

---

Replace every `[BRACKET]` below before sending. Keep the numbers honest — delete any line you can't back up.

Replace these once, everywhere:
- `[GITHUB LINK]` → `https://github.com/darkgambit/My_AI_Projects/tree/main/n8n-clinic-ai-agent` *(after merging the PR; until then use the branch/PR link)*
- `[VIDEO LINK]` → your video URL (YouTube Unlisted / Drive — see VIDEO_RECORDING_GUIDE.md)
- `[YOUR NAME]`, `[X bookings/month]`, `[clinic type/city]` → your real details

---

## 1) The form field: "Link to proof of an automation you built"

```
Live demo (90 sec): [VIDEO LINK]
Full case study + workflows + animated walkthrough: [GITHUB LINK]
```

*If the form allows only ONE link, use the GitHub link — the video is linked at the top of that README.*

## 2) Short version (2–3 sentences — cover letters, DMs)

```
I built a Clinic AI Agent in n8n: patients book, reschedule, and cancel appointments
by chatting on WhatsApp or Messenger. An LLM agent checks live Google Calendar
availability, prevents double-booking, creates the calendar event, and logs every
booking in Google Sheets. Live demo: [VIDEO LINK] · Case study & workflows: [GITHUB LINK]
```

## 3) Long version (portfolio / "describe a project" fields)

```
Clinic AI Agent — AI appointment automation (n8n, OpenRouter, Google Calendar/Sheets)

Problem: [clinic type] staff were handling all bookings manually by phone/messages
during working hours — slow replies, occasional double-bookings, no audit trail.

Solution: Two production n8n workflows.
1) A front-door AI agent: WhatsApp + Messenger webhooks (verification, instant ack,
   payload parsing) feed one LLM agent with conversation memory. It replies on
   whichever channel the patient used.
2) A booking sub-workflow the agent calls as six tools: clinic info, availability,
   book, my-bookings, cancel, reschedule. Every write re-checks Google Calendar
   FreeBusy first, then updates Calendar and a Google Sheets booking log together.
   Taken slots return the nearest alternatives instead of an error.

Result: bookings handled 24/7 in seconds, zero double-bookings, every appointment
traceable in one sheet. [ADD: ≈X bookings/month automated.]

Proof — live demo video: [VIDEO LINK]
Proof — architecture, screenshots, animated walkthrough, workflow files: [GITHUB LINK]
```

## 4) CV / LinkedIn bullet points

```
• Built and deployed a two-workflow n8n AI agent that books, reschedules, and cancels
  clinic appointments over WhatsApp & Facebook Messenger (Meta webhooks, 24/7).
• Integrated OpenRouter LLM with conversation memory + 6 custom tool workflows;
  Google Calendar FreeBusy validation prevents double-booking at write time.
• Dual-write design keeps Google Calendar (staff schedule) and Google Sheets
  (booking audit log) always in sync; graceful fallbacks for taken slots / no booking found.
```

## 5) If they ask "can you walk me through it?" (interview answer, 30 sec)

```
"A patient messages the clinic on WhatsApp. My n8n webhook acknowledges Meta
instantly, parses the message, and hands it to an AI agent with memory of the
conversation. The agent doesn't guess availability — it calls my booking
sub-workflow as a tool, which checks opening hours from Google Sheets against
live Google Calendar free/busy. When the patient picks a time, the slot is
re-validated, the calendar event is created, a row is appended to the bookings
sheet, and the confirmation goes back on WhatsApp. Same engine handles
reschedules and cancellations. I can show you the 90-second live demo."
```

## 6) "One automation you built: what was the manual process before, what did you build, and roughly how many hours does it save?"

```text
Before the automation, clinic staff had to answer appointment requests manually,
ask for the treatment and preferred time, check the clinic hours and calendar for
available slots, go back and forth with the patient, create the calendar event,
record the booking in a spreadsheet, and send a confirmation. Rescheduling or
cancelling meant finding the original booking and updating both records again.
This only worked when staff were available and created a risk of mistakes or
double-bookings.

I built a Clinic AI Agent using two n8n workflows. WhatsApp and Messenger
webhooks receive and parse patient messages, then an OpenRouter-powered AI agent
with conversation memory understands the request and calls booking tools. A
separate booking sub-workflow checks clinic information and live Google Calendar
FreeBusy availability, validates the selected slot again before writing, creates
or updates the calendar event, and updates the Google Sheets booking log at the
same time. It can book, check availability, list a patient's bookings, cancel,
and reschedule, then replies through the same channel the patient used.

A manual booking or change normally takes about 6–10 minutes across the
conversation, availability check, calendar entry, spreadsheet update, and
confirmation. At an average of 8 minutes, 10 requests per day over 26 working
days is about 35 staff hours per month. Allowing roughly 20% for exceptions or
staff oversight, the automation saves approximately 28 hours per month—about
1–1.5 working hours per day—while also responding 24/7 and preventing
double-bookings.
```

**Adjust the estimate to your actual volume:** at 8 minutes per request and an
80% automation rate, 5 requests/day saves roughly **14 hours/month**, 10/day
saves **28 hours/month**, and 20/day saves **56 hours/month**. Replace the
10-per-day assumption with your real n8n execution/booking volume before
submitting if you have it.

## Ranking: what convinces, in order

1. 🎬 **Video of the live system** (YouTube Unlisted etc.) — real message → real calendar event (do this first, ~15 min)
2. 📚 **GitHub case study** — this repo's README (architecture, node inventory, your reasoning)
3. 🎞️ **Animated GIF + interactive demo** — already in the README; makes the page memorable
4. 📥 **Sanitized workflow JSON** — lets a technical reviewer import and inspect it
5. 🖼️ Screenshot alone — weakest; never send only this
