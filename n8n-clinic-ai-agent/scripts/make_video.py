#!/usr/bin/env python3
"""Finished captioned video tour of the n8n workflows.
Output: assets/clinic-ai-agent-workflow-demo.mp4 (1280x720, 30fps, H.264)
        assets/captions.en.vtt / .srt, assets/video-poster.jpg
Requires: pillow, imageio, imageio-ffmpeg, numpy"""
from PIL import Image, ImageDraw, ImageFont
import imageio.v2 as imageio, imageio_ffmpeg, subprocess
import numpy as np, os, math

HERE = os.path.dirname(os.path.abspath(__file__))
A = os.path.join(HERE, "..", "assets")
SRC = os.path.join(A, "workflow-overview.png")
TMP = os.path.join(A, "_video_noaudio.mp4")
OUT = os.path.join(A, "clinic-ai-agent-workflow-demo.mp4")
W, H, HC, FPS = 1280, 720, 600, 30
BG=(20,20,20); PINK=(234,76,137); GRN=(37,211,102); TXT=(242,242,242); SUB=(175,175,175); BARBG=(30,30,30)
FB="/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"; FR="/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
fT=ImageFont.truetype(FB,52); fS=ImageFont.truetype(FR,27); fC=ImageFont.truetype(FR,27); fB=ImageFont.truetype(FB,27); fSm=ImageFont.truetype(FB,20)

src = Image.open(SRC).convert("RGB"); IW, IH = src.size; S = 4
big = src.resize((IW*S, IH*S), Image.LANCZOS)

def clamp_box(cx, cy, w):
    h = w * HC / W
    x1 = min(max(cx - w/2, 0), IW - w); y1 = min(max(cy - h/2, 0), IH - h)
    return (x1, y1, x1 + w, y1 + h)
def lerp(a,b,t): return tuple(a[i] + (b[i]-a[i])*t for i in range(4))
def ease(t): return t*t*(3-2*t)

# dur(s), center, width, highlight(x1,y1,x2,y2 in source coords), step, caption
SHOTS = [
 (5,(386,284),772,None,"Overview","Two production n8n workflows: the Clinic AI Agent (top) and the Booking sub-workflow it calls as tools (bottom)."),
 (6,(185,105),345,(62,18,315,190),"Step 1 · Message in","A patient messages on WhatsApp or Messenger. Webhooks verify with Meta, acknowledge instantly, and parse each channel's payload."),
 (6,(405,140),365,(275,75,565,215),"Step 2 · AI agent","One normalized message reaches the Clinic AI Agent: OpenRouter LLM, conversation memory, and six booking tools."),
 (6,(200,435),365,(90,335,330,545),"Step 3 · Tool engine","Every tool calls one sub-workflow. Route by Action branches to clinic info, availability, book, cancel, my bookings, or reschedule."),
 (6,(455,290),390,(295,215,615,360),"Step 4 · Real availability","Booking checks live availability first: clinic hours from Google Sheets plus Google Calendar FreeBusy, so a taken slot is never offered."),
 (6,(565,315),425,(385,250,772,395),"Step 5 · Book safely","Slot still free? Create the Google Calendar event and append the booking row to Google Sheets. Taken? The patient gets the nearest alternatives."),
 (6,(525,465),445,(315,370,772,565),"Step 6 · Cancel & reschedule","Cancel and reschedule take the same safe path: find the patient's own booking, validate, then update Calendar and Sheets together."),
 (6,(625,120),345,(475,60,772,205),"Step 7 · Instant confirmation","The reply is extracted and routed back on the same channel the patient used: Messenger or WhatsApp. Done in seconds, 24/7."),
]

def wrap(d, text, font, maxw):
    words, lines, cur = text.split(), [], ""
    for wd in words:
        t=(cur+" "+wd).strip()
        if d.textlength(t, font=font) > maxw: lines.append(cur); cur=wd
        else: cur=t
    return lines+[cur]

def card(title_lines, sub, sub2=""):
    f=Image.new("RGB",(W,H),BG); d=ImageDraw.Draw(f)
    d.rectangle([0,0,W,12],fill=PINK); y=250
    for ln in title_lines:
        d.text(((W-d.textlength(ln,font=fT))/2,y),ln,font=fT,fill=TXT); y+=68
    d.text(((W-d.textlength(sub,font=fS))/2,y+26),sub,font=fS,fill=GRN)
    if sub2: d.text(((W-d.textlength(sub2,font=fS))/2,y+70),sub2,font=fS,fill=SUB)
    return f

def render_shot_frame(box, hi, step, cap, pulse, progress):
    x1,y1,x2,y2 = box
    crop = big.crop((int(x1*S),int(y1*S),int(x2*S),int(y2*S))).resize((W,HC), Image.LANCZOS)
    f = Image.new("RGB",(W,H),BG); f.paste(crop,(0,0)); d=ImageDraw.Draw(f,"RGBA")
    if hi:
        hx1=(hi[0]-x1)/(x2-x1)*W; hy1=(hi[1]-y1)/(y2-y1)*HC; hx2=(hi[2]-x1)/(x2-x1)*W; hy2=(hi[3]-y1)/(y2-y1)*HC
        th = 5+int(2*pulse)
        d.rounded_rectangle([hx1,hy1,hx2,hy2],radius=14,outline=PINK+(255,),width=th)
        d.rounded_rectangle([hx1,hy1,hx2,hy2],radius=14,fill=PINK+(18,))
    d.rectangle([0,HC,W,H],fill=BARBG+(255,)); d.rectangle([0,HC,W,HC+5],fill=PINK+(255,))
    d.rectangle([0,HC,int(W*progress),HC+5],fill=GRN+(255,))
    label=step.upper(); lw=d.textlength(label,font=fSm)+28
    d.rounded_rectangle([24,HC+16,24+lw,HC+48],radius=9,fill=PINK+(255,))
    d.text((38,HC+22),label,font=fSm,fill=(255,255,255,255))
    y=HC+58
    for ln in wrap(d,cap,fC,W-60)[:2]:
        d.text((24,y),ln,font=fC,fill=TXT+(255,)); y+=32
    return f

title = card(["Clinic AI Agent","WhatsApp + Messenger Appointment Automation"],"Built with n8n  •  OpenRouter LLM  •  Google Calendar  •  Google Sheets")
endc  = card(["Book  •  Reschedule  •  Cancel","Check Availability  •  My Bookings"],"Full case study, workflow files & interactive demo on GitHub:","github.com/darkgambit/My_AI_Projects")

total = 4 + sum(s[0] for s in SHOTS) + 5
writer = imageio.get_writer(TMP, fps=FPS, codec="libx264", quality=8, macro_block_size=16,
                            ffmpeg_params=["-pix_fmt","yuv420p"])
captions=[]; t=0.0; poster_saved=False
def push(img, dur, cap=None):
    global t, poster_saved
    start=t
    arr=np.asarray(img)
    for _ in range(int(dur*FPS)):
        writer.append_data(arr); t+=1/FPS
    if cap: captions.append((start,t,cap))

push(title,4,"Clinic AI Agent. WhatsApp and Messenger appointment automation, built with n8n, OpenRouter, Google Calendar and Google Sheets.")
prev_box = clamp_box(*SHOTS[0][1], SHOTS[0][2])
for dur,center,wdt,hi,step,cap in SHOTS:
    start=t; target=clamp_box(center[0],center[1],wdt); n=int(dur*FPS); trans=int(0.9*FPS)
    for i in range(n):
        p=i/n
        if i<trans: box=lerp(prev_box,target,ease(i/trans))
        else:
            zw=(target[2]-target[0])*(1-0.03*((i-trans)/max(1,n-trans)))
            box=clamp_box(center[0],center[1],zw)
        frame=render_shot_frame(box,hi,step,cap,abs(math.sin(t*2.2)),t/total)
        if not poster_saved and step=="Overview" and i==n//2:
            frame.save(os.path.join(A,"video-poster.jpg"),quality=88); poster_saved=True
        writer.append_data(np.asarray(frame)); t+=1/FPS
    captions.append((start,t,f"{step}. {cap}")); prev_box=box
push(endc,5,"Book, reschedule, cancel, check availability and my bookings. Full case study and workflow files on GitHub.")
writer.close()

def ts(s,vtt=True):
    ms=int((s%1)*1000); s=int(s); return f"{s//3600:02d}:{s//60%60:02d}:{s%60:02d}{'.' if vtt else ','}{ms:03d}"
with open(os.path.join(A,"captions.en.vtt"),"w") as fh:
    fh.write("WEBVTT\n\n"+"\n\n".join(f"{ts(a)} --> {ts(b)}\n{c}" for a,b,c in captions)+"\n")
with open(os.path.join(A,"captions.en.srt"),"w") as fh:
    fh.write("\n\n".join(f"{i+1}\n{ts(a,False)} --> {ts(b,False)}\n{c}" for i,(a,b,c) in enumerate(captions))+"\n")

ffmpeg=imageio_ffmpeg.get_ffmpeg_exe()
subprocess.run([ffmpeg,"-y","-i",TMP,"-f","lavfi","-i","anullsrc=channel_layout=stereo:sample_rate=44100",
                "-map","0:v","-map","1:a","-c:v","copy","-c:a","aac","-shortest","-movflags","+faststart",OUT],
               check=True, capture_output=True)
os.remove(TMP)
print(f"Wrote {OUT}: {total:.0f}s, {os.path.getsize(OUT)/1e6:.1f} MB + captions.en.vtt/.srt + video-poster.jpg")
