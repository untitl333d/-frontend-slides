"""Episode 5 footage prep: trim every shot to the range the edit uses and bring it to 1080x1920
with one shared grade, so the composition only places finished vertical shots.

modes
  full   cover-crop the source to 9:16 (focus_x picks the horizontal centre, 0..1)
  panel  horizontal source as a sharp framed panel over a blurred, darkened copy of itself
  rot    top-down horizontal source turned 90 degrees to fill the vertical frame
"""
import os, subprocess

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A = os.path.join(ROOT, "..", "assets")
OUT = os.path.join(ROOT, "media")

GRADE = ("eq=contrast=1.12:brightness=-0.035:saturation=0.72,"
         "colorbalance=rs=-0.06:bs=0.07:rm=-0.03:bm=0.04:rh=-0.02:bh=0.02")
GRADE_SAFE = "eq=contrast=1.05:saturation=0.85,colorbalance=bs=0.03"

# name, source, start, duration, mode, options
SHOTS = [
    ("c01_walk", "ep05/broll/g1.mp4", 3.0, 2.0, "panel", {}),
    ("c02_bag", "ep05/broll/g1.mp4", 15.0, 2.2, "panel", {}),
    ("c03_cash_car", "ep05/broll/g6_dark.mp4", 1.0, 1.2, "full", {}),
    ("c04_monitors", "ep05/broll/g2.mp4", 3.0, 1.2, "panel", {}),
    ("c05_phone", "ep05/voice/IMG_1625.MOV", 5.0, 2.7, "full", {"focus_x": 0.5}),
    ("c06_uv_tease", "ep05/broll/pexels_6266763.mp4", 2.0, 0.9, "full", {}),
    ("c07_drone", "ep05/broll/b1.mp4", 6.0, 1.5, "full", {}),
    ("c08_mouse", "ep05/broll/b2.mp4", 0.0, 1.6, "panel", {"crop": "432:336:0:0", "panel_w": 980}),
    ("c09_cuffs", "ep05/broll/b3.mp4", 1.5, 1.8, "full", {"focus_x": 0.5}),
    ("c10_latch", "ep05/broll/g8.mp4", 0.9, 1.0, "full", {}),
    ("c11_silhouettes", "ep05/broll/g5.mp4", 1.5, 1.2, "full", {"focus_x": 0.5}),
    ("c12_flashlight", "ep05/broll/g7.mp4", 2.0, 1.5, "panel", {}),
    ("c13_uv", "ep05/broll/pexels_6266763.mp4", 3.0, 0.8, "full", {}),
    ("c14_detector", "ep05/broll/g9.mp4", 4.6, 0.8, "panel", {}),
    ("c15_burn", "ep05/broll/b5.mp4", 12.0, 3.8, "full", {"focus_x": 0.28}),
    ("c16_hands_phone", "ep05/broll/u1.mov", 1.0, 1.3, "full", {}),
    ("c17_teller", "ep05/broll/g3.mp4", 2.5, 2.6, "panel", {"grade": GRADE_SAFE}),
    ("c18_police", "ep05/broll/pexels_15529015.mp4", 0.0, 4.0, "full", {}),
    ("c19_bills", "ep05/video/op_clean/op5_bills_grid.mp4", 0.0, 1.72, "rot", {}),
]


def probe_size(path):
    out = subprocess.run(["ffprobe", "-v", "error", "-select_streams", "v", "-show_entries", "stream=width,height:stream_side_data=rotation",
                          "-of", "csv=p=0", path], capture_output=True, text=True).stdout.split()
    w, h = map(int, out[0].split(",")[:2])
    rot = [l for l in out if l.lstrip("-").isdigit()]
    if rot and abs(int(rot[0])) == 90:
        w, h = h, w
    return w, h


def build(name, src, ss, dur, mode, opt):
    path = os.path.join(A, src)
    grade = opt.get("grade", GRADE)
    pre = f"crop={opt['crop']}," if "crop" in opt else ""
    if mode == "full":
        w, h = probe_size(path)
        fx = opt.get("focus_x", 0.5)
        vf = (f"{pre}scale=-2:1920:flags=lanczos,crop=1080:1920:'(in_w-1080)*{fx}':0," if w / h > 1080 / 1920
              else f"{pre}scale=1080:-2:flags=lanczos,crop=1080:1920:0:'(in_h-1920)/2',")
        vf += f"{grade},fps=30,format=yuv420p"
        fc = None
    elif mode == "rot":
        vf = f"{pre}transpose=1,scale=1080:-2:flags=lanczos,crop=1080:1920,{grade},fps=30,format=yuv420p"
        fc = None
    else:
        pw = opt.get("panel_w", 1080)
        fc = (f"[0:v]{pre}fps=30,split[a][b];"
              f"[a]scale=-2:1920,crop=1080:1920,gblur=sigma=38,eq=brightness=-0.22:saturation=0.5[bg];"
              f"[b]scale={pw}:-2:flags=lanczos,pad=iw+8:ih+8:4:4:color=0xeeeeee[fg];"
              f"[bg][fg]overlay=(W-w)/2:820-h/2,{grade},format=yuv420p[v]")
    cmd = ["ffmpeg", "-hide_banner", "-loglevel", "error", "-y", "-ss", str(ss), "-t", str(dur), "-i", path, "-an"]
    cmd += (["-filter_complex", fc, "-map", "[v]"] if fc else ["-vf", vf])
    cmd += ["-c:v", "libx264", "-crf", "17", "-preset", "medium", "-g", "15", os.path.join(OUT, name + ".mp4")]
    subprocess.run(cmd, check=True)


if __name__ == "__main__":
    import sys
    os.makedirs(OUT, exist_ok=True)
    for s in SHOTS:
        if len(sys.argv) > 1 and s[0] not in sys.argv[1:]:
            continue
        build(*s)
        print("ok", s[0])
