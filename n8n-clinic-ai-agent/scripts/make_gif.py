#!/usr/bin/env python3
"""Animated GIF tour of the n8n workflow screenshot (pan/zoom + captions).
Regenerate: python3 make_gif.py  (requires Pillow)"""
from PIL import Image, ImageDraw, ImageFont
import os

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "assets", "workflow-overview.png")
OUT = os.path.join(HERE, "..", "assets", "workflow-demo.gif")
W, H = 1152, 648
BAR = 92                      # caption bar height
BG = (20, 20, 20); ACC = (234, 76, 137); GRN = (37, 211, 102); TXT = (240, 240, 240); SUB=(170,170,170)
F_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
F_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
f_title = ImageFont.truetype(F_B, 40); f_sub = ImageFont.truetype(F_R, 24)
f_cap = ImageFont.truetype(F_B, 25); f_small = ImageFont.truetype(F_R, 21)

img = Image.open(SRC).convert("RGB")
iw, ih = img.size  # 772 x 568

def crop_resize(box):
    c = img.crop(box)
    # fit into W x (H-BAR) keeping aspect, upscale
    tw, th = W, H - BAR
    s = min(tw / c.width, th / c.height)
    c = c.resize((max(1,int(c.width*s)), max(1,int(c.height*s))), Image.LANCZOS)
    canvas = Image.new("RGB", (W, H), BG)
    canvas.paste(c, ((W - c.width)//2, (th - c.height)//2))
    return canvas

def caption(frame, step, text):
    d = ImageDraw.Draw(frame)
    d.rectangle([0, H-BAR, W, H], fill=(32,32,32))
    d.rectangle([0, H-BAR, W, H-BAR+4], fill=ACC)
    d.rounded_rectangle([22, H-BAR+22, 62, H-22], radius=8, fill=ACC)
    d.text((34, H-BAR+28), str(step), font=f_cap, fill=(255,255,255))
    # wrap
    words, lines, cur = text.split(), [], ""
    for w_ in words:
        t = (cur + " " + w_).strip()
        if d.textlength(t, font=f_small) > W - 110: lines.append(cur); cur = w_
        else: cur = t
    lines.append(cur)
    y = H - BAR + 14
    for ln in lines[:3]:
        d.text((78, y), ln, font=f_small, fill=TXT); y += 26
    return frame

def title_card(lines, sub):
    f = Image.new("RGB", (W, H), BG); d = ImageDraw.Draw(f)
    d.rectangle([0,0,W,10], fill=ACC)
    y = 190
    for ln in lines:
        wlen = d.textlength(ln, font=f_title)
        d.text(((W-wlen)/2, y), ln, font=f_title, fill=TXT); y += 54
    wlen = d.textlength(sub, font=f_sub)
    d.text(((W-wlen)/2, y+22), sub, font=f_sub, fill=GRN)
    return f

FULL = (0, 0, iw, ih)
# (crop box in source coords, step label, caption)
SHOTS = [
    (FULL,             "•", "Two production n8n workflows: a Clinic AI Agent (top) + a Booking sub-workflow it calls as tools (bottom)."),
    ((55, 10, 330, 190),  "1", "Patient messages on WhatsApp or Messenger. Webhooks handle Meta verification, instant Ack, then parse each channel's payload."),
    ((250, 60, 560, 220), "2", "One normalized message reaches the Clinic AI Agent — OpenRouter chat model + conversation memory + 6 booking tools."),
    ((70, 340, 330, 540), "3", "Every tool calls ONE sub-workflow: 'Route by Action' branches to clinic info, availability, book, cancel, my-bookings, reschedule."),
    ((300, 220, 620, 360),"4", "Booking checks live availability first: Google Sheets clinic hours + Google Calendar FreeBusy — so a taken slot can never be double-booked."),
    ((430, 250, 772, 400),"5", "Slot free? Create Calendar Event, then append the booking row to Google Sheets, and return 'Booking Confirmed'."),
    ((350, 380, 772, 568),"6", "Cancel & reschedule follow the same safe path: find the patient's booking, validate the new slot, update Calendar + Sheets together."),
    ((480, 60, 772, 200), "7", "The agent's reply is extracted and routed back on the SAME channel the patient used — Messenger or WhatsApp."),
]

frames, durations = [], []
t = title_card(["Clinic AI Agent", "WhatsApp + Messenger Appointment Automation"], "Built with n8n  •  OpenRouter LLM  •  Google Calendar  •  Google Sheets")
frames.append(t); durations.append(1800)

prev = None
for box, step, cap in SHOTS:
    target = caption(crop_resize(box), step, cap)
    if prev is not None:  # smooth crossfade/pan feel
        for a in (0.33, 0.66):
            frames.append(Image.blend(prev, target, a)); durations.append(110)
    frames.append(target); durations.append(2100 if step != "•" else 1600)
    prev = target

e = title_card(["Book  •  Check Availability  •  Cancel", "Reschedule  •  My Bookings  •  Clinic Info"], "Full architecture, node inventory & setup — in the README below")
frames.append(e); durations.append(1800)

frames[0].save(OUT, save_all=True, append_images=frames[1:], duration=durations, loop=0, optimize=True)
print(f"Wrote {OUT}: {len(frames)} frames, {os.path.getsize(OUT)/1e6:.1f} MB")
