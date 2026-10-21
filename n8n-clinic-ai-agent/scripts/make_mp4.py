#!/usr/bin/env python3
"""Convert assets/workflow-demo.gif -> assets/workflow-demo.mp4 (upload-ready for
YouTube / LinkedIn / Google Drive — no Loom or screen recording needed).
Regenerate: python3 make_mp4.py   (requires: pillow, imageio, imageio-ffmpeg)"""
from PIL import Image
import imageio.v2 as imageio
import numpy as np, os

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "assets", "workflow-demo.gif")
OUT = os.path.join(HERE, "..", "assets", "workflow-demo.mp4")
FPS = 10

gif = Image.open(SRC)
frames = []
for i in range(gif.n_frames):
    gif.seek(i)
    dur_ms = gif.info.get("duration", 100)
    reps = max(1, round(dur_ms / 1000 * FPS))   # expand to constant FPS
    frame = np.asarray(gif.convert("RGB"))
    frames.extend([frame] * reps)

# H.264, yuv420p -> plays everywhere (YouTube, LinkedIn, WhatsApp, browsers)
imageio.mimwrite(OUT, frames, fps=FPS, codec="libx264", quality=7,
                 macro_block_size=16, ffmpeg_params=["-pix_fmt", "yuv420p", "-movflags", "+faststart"])
print(f"Wrote {OUT}: {len(frames)} frames @ {FPS}fps = {len(frames)/FPS:.0f}s, {os.path.getsize(OUT)/1e6:.1f} MB")
