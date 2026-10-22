#!/usr/bin/env python3
"""V2: Full-HD captioned video with the workflow REDRAWN as crisp labelled
boxes (the original screenshot is 772px, so zooming it makes labels unreadable).
The real n8n canvas is shown full-frame near the end for authenticity.
Output: assets/clinic-ai-agent-workflow-demo.mp4 (1920x1080) + captions + poster."""
from PIL import Image, ImageDraw, ImageFont
import imageio.v2 as imageio, imageio_ffmpeg, subprocess
import numpy as np, os, math

HERE=os.path.dirname(os.path.abspath(__file__)); A=os.path.join(HERE,"..","assets")
SRC=os.path.join(A,"workflow-overview.png"); TMP=os.path.join(A,"_v2_noaudio.mp4"); OUT=os.path.join(A,"clinic-ai-agent-workflow-demo.mp4")
W,H,HC,FPS=1920,1080,884,30
BG=(20,20,20); PANEL=(38,38,38); LINE=(90,90,90); TXT=(242,242,242); SUB=(178,178,178)
PINK=(234,76,137); AMBER=(245,166,35); GRN=(37,211,102); BLUE=(74,158,255); TEAL=(45,212,191); VIO=(167,139,250); GREY=(150,150,150)
FB="/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"; FR="/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
fT=ImageFont.truetype(FB,74); fH=ImageFont.truetype(FB,42); fCap=ImageFont.truetype(FR,36)
fBox=ImageFont.truetype(FB,29); fBoxS=ImageFont.truetype(FB,25); fSub=ImageFont.truetype(FR,22); fChip=ImageFont.truetype(FB,23); fStep=ImageFont.truetype(FB,26)

def wrap(d,text,font,maxw):
    words,lines,cur=text.split(),[],""
    for w_ in words:
        t=(cur+" "+w_).strip()
        if d.textlength(t,font=font)>maxw and cur: lines.append(cur); cur=w_
        else: cur=t
    return lines+[cur]

def box(d,x,y,w,h,title,sub,color,state,icon=""):
    border = color if state=="on" else (GRN if state=="done" else LINE)
    width = 7 if state=="on" else 3
    if state=="on":
        d.rounded_rectangle([x-9,y-9,x+w+9,y+h+9],radius=22,outline=color,width=2)
    d.rounded_rectangle([x,y,x+w,y+h],radius=18,fill=PANEL,outline=border,width=width)
    d.ellipse([x+16,y+16,x+62,y+62],fill=color)
    if icon:
        iw=d.textlength(icon,font=fBoxS); d.text((x+39-iw/2,y+25),icon,font=fBoxS,fill=(18,18,18))
    font=fBox if w>=300 else fBoxS
    lines=wrap(d,title,font,w-36); ty=y+72
    for ln in lines[:3]:
        d.text((x+(w-d.textlength(ln,font=font))/2,ty),ln,font=font,fill=TXT); ty+=34
    if sub:
        for ln in wrap(d,sub,fSub,w-30)[:2]:
            d.text((x+(w-d.textlength(ln,font=fSub))/2,ty+2),ln,font=fSub,fill=SUB); ty+=27

def arrow(d,x1,y1,x2,y2,on=False):
    c=GRN if on else LINE; wd=5 if on else 3
    if abs(x2-x1)<4:  # vertical
        d.line([x1,y1,x2,y2],fill=c,width=wd)
        sgn=1 if y2>y1 else -1
        d.polygon([(x2,y2),(x2-8,y2-14*sgn),(x2+8,y2-14*sgn)],fill=c); return
    if abs(y2-y1)<4:
        d.line([x1,y1,x2,y2],fill=c,width=wd)
    else:
        mx=(x1+x2)//2; d.line([x1,y1,mx,y1,mx,y2,x2,y2],fill=c,width=wd,joint="curve")
    sgn=1 if x2>x1 else -1
    d.polygon([(x2,y2),(x2-14*sgn,y2-8),(x2-14*sgn,y2+8)],fill=c)

def chip(d,x,y,text,color,state):
    w=d.textlength(text,font=fChip)+36
    border=color if state in("on","done") else LINE
    d.rounded_rectangle([x,y,x+w,y+46],radius=23,fill=(30,30,30),outline=border,width=4 if state=="on" else 2)
    d.text((x+18,y+10),text,font=fChip,fill=TXT if state!="idle" else SUB)
    return w

def scene_frame(heading,boxes,links,active,cap,step,progress,chips_row=None):
    f=Image.new("RGB",(W,H),BG); d=ImageDraw.Draw(f)
    d.text((60,34),heading,font=fH,fill=TXT)
    for (a,b) in links:
        ax,ay,aw,ah=boxes[a][:4]; bx,by,bw,bh=boxes[b][:4]
        if abs(ax-bx)<5 and abs(ay-by)>5:      # stacked: bottom -> top
            y1,y2=(ay+ah,by) if by>ay else (ay,by+bh)
            arrow(d,ax+aw//2,y1,bx+bw//2,y2,on=(b<=active))
        elif bx<ax and abs(ay-by)<5:           # right-to-left same row
            arrow(d,ax,ay+ah//2,bx+bw,by+bh//2,on=(b<=active))
        else:
            arrow(d,ax+aw,ay+ah//2,bx,by+bh//2,on=(b<=active))
    for i,(x,y,w,h,t,s,c,ic) in enumerate(boxes):
        box(d,x,y,w,h,t,s,c,"on" if i==active else("done" if i<active else "idle"),ic)
    if chips_row:
        x,y,items=chips_row
        for j,(txt,c) in enumerate(items):
            x+=chip(d,x,y,txt,c,"on" if j==active%len(items) else "idle")+16
    d.rectangle([0,HC,W,H],fill=(30,30,30)); d.rectangle([0,HC,W,HC+6],fill=PINK); d.rectangle([0,HC,int(W*progress),HC+6],fill=GRN)
    lab=step.upper(); lw=d.textlength(lab,font=fStep)+40
    d.rounded_rectangle([36,HC+22,36+lw,HC+66],radius=12,fill=PINK); d.text((56,HC+30),lab,font=fStep,fill=(255,255,255))
    y=HC+82
    for ln in wrap(d,cap,fCap,W-90)[:2]:
        d.text((36,y),ln,font=fCap,fill=TXT); y+=44
    return f

def chain(n,bw,bh,y,x0=60,gap=None):
    gap=gap or int((W-120-bw*n)/max(1,n-1))
    return [(x0+i*(bw+gap),y,bw,bh) for i in range(n)]

M=60
SCENES=[]
# Overview
pos=chain(5,320,190,320)
SCENES.append(dict(dur=6,step="Overview",heading="Clinic AI Agent — how it works",
 cap="Patients chat on WhatsApp or Messenger. An AI agent books appointments against the live calendar and logs everything in Google Sheets.",
 boxes=[pos[0]+("Patient","WhatsApp · Messenger",GREY,"P"),pos[1]+("Webhooks","Verify · Ack · Parse",PINK,"W"),
        pos[2]+("Clinic AI Agent","OpenRouter LLM + Memory",VIO,"AI"),pos[3]+("Booking Engine","Route by Action (6 tools)",TEAL,"R"),
        pos[4]+("Calendar + Sheets","always updated together",BLUE,"G")],
 links=[(0,1),(1,2),(2,3),(3,4)]))
# Step 1 two rows converge
r1=chain(4,335,190,205,gap=28); r2=chain(4,335,190,525,gap=28)
conv=(1525,365,335,190)
SCENES.append(dict(dur=7,step="Step 1 · Message in",heading="Two channels, one front door",
 cap="Meta webhooks verify, acknowledge instantly (so Meta never times out), then parse each channel's payload into one format.",
 boxes=[r1[0]+("Messenger Webhook","Meta",PINK,"M"),r1[1]+("Messenger Verification","IF → verify token",PINK,"✓"),r1[2]+("Messenger Ack","200 OK instantly",PINK,"A"),r1[3]+("Parse Messenger Message","Code node",AMBER,"{}"),
        r2[0]+("WhatsApp Webhook","Meta",GRN,"W"),r2[1]+("WhatsApp Verification","verify token",GRN,"✓"),r2[2]+("WhatsApp Ack","200 OK instantly",GRN,"A"),r2[3]+("Parse WhatsApp Message","Code node",AMBER,"{}"),
        conv+("Parse Message","one clean format",AMBER,"{}")],
 links=[(0,1),(1,2),(2,3),(4,5),(5,6),(6,7),(3,8),(7,8)]))
# Step 2 agent
pos=chain(3,460,210,240,gap=60)
SCENES.append(dict(dur=7,step="Step 2 · AI agent",heading="The brain: Clinic AI Agent",
 cap="The agent remembers the conversation and chooses the right booking tool. It never guesses availability: it asks the booking engine.",
 boxes=[pos[0]+("Parse Message","{ channel, sender, text }",AMBER,"{}"),pos[1]+("Clinic AI Agent","OpenRouter Chat Model",VIO,"AI"),pos[2]+("Conversation Memory","per-patient context",VIO,"≡")],
 links=[(0,1),(1,2)],
 chips=(60,560,[("get_clinic_info",AMBER),("check_availability",BLUE),("book_appointment",GRN),("get_my_bookings",GREY),("cancel_appointment",PINK),("reschedule_appointment",TEAL)])))
# Step 3 route
pos=chain(2,520,210,230,x0=120,gap=80)
SCENES.append(dict(dur=6,step="Step 3 · Tool engine",heading="One sub-workflow behind every tool",
 cap="When Executed by Another Workflow feeds Route by Action: one safe, testable branch per action. Nothing writes data except through here.",
 boxes=[pos[0]+("When Executed by Another Workflow","tool entry point",GREY,"»"),pos[1]+("Route by Action","Switch",TEAL,"R")],
 links=[(0,1)],
 chips=(120,540,[("clinic info",AMBER),("availability",BLUE),("book",GRN),("cancel",PINK),("my bookings",GREY),("reschedule",TEAL)])))
# Step 4 availability
pos=chain(4,410,200,330,gap=36)
SCENES.append(dict(dur=7,step="Step 4 · Real availability",heading="Availability = hours + live calendar",
 cap="Clinic hours and treatments come from Google Sheets. Busy times come from Google Calendar FreeBusy. Only truly free slots are offered.",
 boxes=[pos[0]+("Read Clinic Info","Google Sheets",GRN,"S"),pos[1]+("Read Day Bookings","Google Sheets",GRN,"S"),pos[2]+("Check Calendar FreeBusy","Google Calendar",BLUE,"C"),pos[3]+("Compute Availability","Code node",AMBER,"{}")],
 links=[(0,1),(1,2),(2,3)]))
# Step 5 book 3+3
r1=chain(3,520,185,210,gap=50); r2=chain(3,520,185,520,gap=50)
SCENES.append(dict(dur=7,step="Step 5 · Book safely",heading="Booking with double-booking protection",
 cap="The chosen slot is re-validated at the moment of booking. Free: calendar event plus Sheets row. Taken: nearest alternatives, never an error or a clash.",
 boxes=[r1[0]+("Check Calendar Slot","Google Calendar",BLUE,"C"),r1[1]+("Validate Book Slot","Code node",AMBER,"{}"),r1[2]+("Slot Free?","Switch",TEAL,"?"),
        r2[2]+("Create Calendar Event","Google Calendar",BLUE,"C"),r2[1]+("Append Booking Row","Google Sheets",GRN,"S"),r2[0]+("Booking Confirmed","or Booking Slot Taken",AMBER,"✓")],
 links=[(0,1),(1,2),(2,3),(3,4),(4,5)]))
# Step 6 cancel/reschedule two chains
r1=chain(4,410,175,220,gap=36); r2=chain(4,410,175,540,gap=36)
SCENES.append(dict(dur=7,step="Step 6 · Cancel & reschedule",heading="Patients manage their own bookings",
 cap="Cancel: find the patient's booking, delete the calendar event, mark the row. Reschedule: validate the new slot first, then move both records together.",
 boxes=[r1[0]+("Find Bookings To Cancel","Google Sheets",GRN,"S"),r1[1]+("Pick Booking To Cancel","Code node",AMBER,"{}"),r1[2]+("Delete Calendar Event","Google Calendar",BLUE,"C"),r1[3]+("Cancel Confirmed","Mark Booking Cancelled",PINK,"✕"),
        r2[0]+("Find Bookings To Move","Google Sheets",GRN,"S"),r2[1]+("Check New Slot","Validate New Slot",BLUE,"C"),r2[2]+("Update Calendar Event","New Slot Free?",TEAL,"U"),r2[3]+("Reschedule Confirmed","Update Booking Row",GRN,"✓")],
 links=[(0,1),(1,2),(2,3),(4,5),(5,6),(6,7)]))
# Step 7 reply
pos=chain(4,410,200,330,gap=36)
SCENES.append(dict(dur=6,step="Step 7 · Instant confirmation",heading="Back on the same channel",
 cap="Extract Reply, then Which Channel? sends the confirmation by Messenger or WhatsApp: the exact channel the patient used. 24/7, in seconds.",
 boxes=[pos[0]+("Extract Reply","Code node",AMBER,"{}"),pos[1]+("Which Channel?","Switch",TEAL,"?"),pos[2]+("Send Messenger / WhatsApp Reply","Graph API / WhatsApp",PINK,"➤"),pos[3]+("Patient confirmed","✅ booked / moved / cancelled",GRN,"✓")],
 links=[(0,1),(1,2),(2,3)]))

def card(lines,sub,sub2=""):
    f=Image.new("RGB",(W,H),BG); d=ImageDraw.Draw(f); d.rectangle([0,0,W,14],fill=PINK); y=380
    for ln in lines: d.text(((W-d.textlength(ln,font=fT))/2,y),ln,font=fT,fill=TXT); y+=96
    d.text(((W-d.textlength(sub,font=fCap))/2,y+34),sub,font=fCap,fill=GRN)
    if sub2: d.text(((W-d.textlength(sub2,font=fCap))/2,y+96),sub2,font=fCap,fill=SUB)
    return f

real=Image.open(SRC).convert("RGB")
def real_frame(progress,step,cap):
    f=Image.new("RGB",(W,H),BG); d=ImageDraw.Draw(f)
    d.text((60,34),"The actual production workflows in n8n",font=fH,fill=TXT)
    scale=min((W-240)/real.width,(HC-150)/real.height)
    im=real.resize((int(real.width*scale),int(real.height*scale)),Image.LANCZOS)
    x=(W-im.width)//2; y=110; d.rectangle([x-6,y-6,x+im.width+6,y+im.height+6],outline=PINK,width=4); f.paste(im,(x,y))
    d.rectangle([0,HC,W,H],fill=(30,30,30)); d.rectangle([0,HC,W,HC+6],fill=PINK); d.rectangle([0,HC,int(W*progress),HC+6],fill=GRN)
    lab=step.upper(); lw=d.textlength(lab,font=fStep)+40
    d.rounded_rectangle([36,HC+22,36+lw,HC+66],radius=12,fill=PINK); d.text((56,HC+30),lab,font=fStep,fill=(255,255,255))
    yy=HC+82
    for ln in wrap(d,cap,fCap,W-90)[:2]: d.text((36,yy),ln,font=fCap,fill=TXT); yy+=44
    return f

title=card(["Clinic AI Agent","WhatsApp + Messenger Appointment Automation"],"Built with n8n  •  OpenRouter LLM  •  Google Calendar  •  Google Sheets","Workflow diagram redrawn in Full HD so every node name is readable")
endc=card(["Book  •  Reschedule  •  Cancel","Check Availability  •  My Bookings"],"Full case study, workflow files & interactive demo on GitHub:","github.com/darkgambit/My_AI_Projects")

total=3+sum(s["dur"] for s in SCENES)+6+4
writer=imageio.get_writer(TMP,fps=FPS,codec="libx264",quality=8,macro_block_size=16,ffmpeg_params=["-pix_fmt","yuv420p"])
caps=[]; t=0.0; poster=False
def push(img,dur,cap):
    global t; s=t; arr=np.asarray(img)
    for _ in range(int(dur*FPS)): writer.append_data(arr); t+=1/FPS
    caps.append((s,t,cap))

push(title,3,"Clinic AI Agent. WhatsApp and Messenger appointment automation built with n8n.")
for sc in SCENES:
    start=t; n=int(sc["dur"]*FPS); nb=len(sc["boxes"])
    for i in range(n):
        active=min(nb-1,int((i/n)*nb*1.12))
        fr=scene_frame(sc["heading"],sc["boxes"],sc["links"],active,sc["cap"],sc["step"],t/total,chips_row=sc.get("chips"))
        if not poster and sc["step"]=="Overview" and i==n//2:
            fr.save(os.path.join(A,"video-poster.jpg"),quality=90); poster=True
        writer.append_data(np.asarray(fr)); t+=1/FPS
    caps.append((start,t,f'{sc["step"]}. {sc["cap"]}'))
cap_real="This is the real n8n canvas: Clinic AI Agent on top, booking sub-workflow below. The same node names you just saw, drawn large."
start=t
for i in range(6*FPS):
    writer.append_data(np.asarray(real_frame(t/total,"Real canvas",cap_real))); t+=1/FPS
caps.append((start,t,"Real canvas. "+cap_real))
push(endc,4,"Book, reschedule, cancel, check availability and my bookings. Full case study on GitHub.")
writer.close()

def ts(s,v=True):
    ms=int((s%1)*1000); s=int(s); return f"{s//3600:02d}:{s//60%60:02d}:{s%60:02d}{'.' if v else ','}{ms:03d}"
open(os.path.join(A,"captions.en.vtt"),"w").write("WEBVTT\n\n"+"\n\n".join(f"{ts(a)} --> {ts(b)}\n{c}" for a,b,c in caps)+"\n")
open(os.path.join(A,"captions.en.srt"),"w").write("\n\n".join(f"{i+1}\n{ts(a,False)} --> {ts(b,False)}\n{c}" for i,(a,b,c) in enumerate(caps))+"\n")
ff=imageio_ffmpeg.get_ffmpeg_exe()
subprocess.run([ff,"-y","-i",TMP,"-f","lavfi","-i","anullsrc=channel_layout=stereo:sample_rate=44100","-map","0:v","-map","1:a","-c:v","copy","-c:a","aac","-shortest","-movflags","+faststart",OUT],check=True,capture_output=True)
os.remove(TMP); print(f"Wrote {OUT}: {total:.0f}s {os.path.getsize(OUT)/1e6:.1f} MB")
