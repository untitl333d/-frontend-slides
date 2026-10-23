"""Episode 5 soundtrack: Socket Calibration (with a bar-aligned splice from the groove straight into the
breakdown), processed voiceover and SFX -> media/mix.wav.

Music bar = 1.894 s. Track time 0 sits at video MUSIC_AT, so the drop (track 15.184) lands on video 16.0.
At video SPLICE (a downbeat, track 28.442) the edit jumps to track 75.792, the downbeat where the bass drops out.
"""
import os, subprocess

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A = os.path.join(ROOT, "..", "assets")
OUT = os.path.join(ROOT, "media", "mix.wav")
DUR = 50.0
MUSIC = os.path.join(A, "ep05/music/socket_calibration.mp3")
MUSIC_AT = 0.816
SPLICE = 29.258
OUT_TRACK = SPLICE - MUSIC_AT          # 28.442, downbeat
IN_TRACK = 75.792                      # breakdown downbeat
FADE_END = 48.4

VOICE = [  # file, video start
    ("ep05/voice/processed/IMG_1631.wav", 2.00),
    ("ep05/voice/processed/IMG_1635.wav", 8.90),
    ("ep05/voice/processed/IMG_1637.wav", 16.30),
    ("ep05/voice/processed/IMG_1642.wav", 20.75),
    ("ep05/voice/processed/IMG_1645.wav", 29.55),
    ("ep05/voice/processed/IMG_1647.wav", 35.00),
    ("ep05/voice/processed/IMG_1651.wav", 38.57),
]

L = "sfx/"            # HyperFrames library SFX (Pixabay licence)
M = "ep05/sfx/"       # Mixkit
P = "ep05/sfx/pixabay/"
# Where the audible hit sits inside each file, so the SFX list below can be written in picture time.
ONSET = {"stamp.mp3": 0.54, "handcuffs_ratchet.mp3": 0.12, "metal_door_slam.mp3": 0.08, "camera_shutter.mp3": 0.09,
         "click.mp3": 0.05, "impact-bass-1.mp3": 0.05, "notification.mp3": 0.08, "ping.mp3": 0.32}
# file, video time of the hit, gain dB, [file in, file out]
SFX = [
    (L + "glitch-1.mp3", 0.05, -14, 0, 0.5), (L + "glitch-3.mp3", 1.05, -16, 0, 0.4),
    ("ep04/video/v5.mp4", 3.88, -13, 3.0, 5.0),               # counting machine under the count-up
    (P + "stamp.mp3", 7.19, -6, 0.45, 1.3),                   # ФАЛЬШИВКА on $362 000
    (L + "ping.mp3", 7.88, -16),                              # $900
    (P + "camera_shutter.mp3", 8.90, -10),
    (L + "whoosh-short.mp3", 9.84, -16),
    (L + "whoosh.mp3", 10.80, -16),
    (L + "notification.mp3", 11.90, -14, 0, 1.2),
    (P + "shuffle_papers.mp3", 13.38, 4, 0, 1.2),
    (L + "glitch-1.mp3", 14.78, -15, 0, 0.35),
    (L + "impact-bass-1.mp3", 16.00, -13),                    # drop
    (P + "stamp.mp3", 16.947, -7, 0.45, 1.3),                 # СЕКРЕТНО, half bar after the drop
    (P + "camera_shutter.mp3", 17.26, -10),
    (L + "click.mp3", 19.79, -8), (L + "impact-bass-1.mp3", 19.79, -15),   # mousetrap
    (M + "siren_passing.mp3", 20.60, -26, 0, 5.5),
    (P + "handcuffs_ratchet.mp3", 21.682, -6),
    (P + "shuffle_papers.mp3", 22.15, 4, 0, 1.0),
    (L + "typing.mp3", 22.35, -18, 0, 0.6), (L + "typing.mp3", 23.45, -18, 0, 0.6), (L + "typing.mp3", 24.63, -18, 0, 0.9),
    (P + "stamp.mp3", 25.47, -5, 0.45, 1.3),                  # ЗАДЕРЖАН
    (P + "metal_door_slam.mp3", 26.50, -10),
    (P + "stamp.mp3", 27.364, -7, 0.45, 1.3), (P + "stamp.mp3", 28.311, -7, 0.45, 1.3),   # В РОЗЫСКЕ x2
    (L + "glitch-3.mp3", 29.20, -14, 0, 0.35),
    (M + "clock_tick.mp3", 29.30, -24, 0, 4.9),
    (P + "camera_shutter.mp3", 29.55, -11), (P + "camera_shutter.mp3", 30.55, -11), (P + "camera_shutter.mp3", 31.60, -12),
    (L + "whoosh-short.mp3", 32.70, -15),
    (M + "beep_fail.mp3", 33.80, -14, 0, 0.9),
    (L + "ping.mp3", 33.73, -15),
    (L + "whoosh.mp3", 36.78, -16),
    (P + "camera_shutter.mp3", 38.57, -11),
    (L + "whoosh-short.mp3", 39.80, -16),
    (L + "ping.mp3", 40.99, -20),
    (L + "glitch-3.mp3", 43.35, -16, 0, 0.4),
    (P + "stamp.mp3", 44.41, -4, 0.45, 1.3), (L + "impact-bass-1.mp3", 44.41, -12),     # ФАЛЬШИВКА / вот этим
    (L + "glitch-1.mp3", 45.28, -16, 0, 0.35),
    (L + "typing.mp3", 46.30, -20, 0, 0.8),
]


def run():
    inp, ch = [], []

    def add(path, *pre):
        inp.extend([*pre, "-i", path])
        return inp.count("-i") - 1

    # music: two segments joined on downbeats, then ducked under the voice
    a = add(MUSIC, "-ss", "0", "-t", f"{OUT_TRACK + 0.02:.3f}")
    b = add(MUSIC, "-ss", f"{IN_TRACK - 0.01:.3f}", "-t", f"{DUR - SPLICE + 0.1:.3f}")
    ch.append(f"[{a}:a]aresample=48000,afade=t=out:st={OUT_TRACK - 0.015:.3f}:d=0.03,adelay={int(MUSIC_AT * 1000)}|{int(MUSIC_AT * 1000)}[ma]")
    ch.append(f"[{b}:a]aresample=48000,afade=t=in:d=0.02,adelay={int((SPLICE - 0.01) * 1000)}|{int((SPLICE - 0.01) * 1000)}[mb]")
    ch.append(f"[ma][mb]amix=inputs=2:normalize=0,afade=t=out:st={FADE_END}:d={DUR - FADE_END},volume=-5dB[music]")

    vl = []
    for i, (f, st) in enumerate(VOICE):
        n = add(os.path.join(A, f))
        ch.append(f"[{n}:a]aresample=48000,aformat=channel_layouts=stereo,adelay={int(st * 1000)}|{int(st * 1000)}[v{i}]")
        vl.append(f"[v{i}]")
    ch.append(f"{''.join(vl)}amix=inputs={len(vl)}:normalize=0,apad=whole_dur={DUR},asplit=2[voice][vkey]")
    ch.append("[music][vkey]sidechaincompress=threshold=0.03:ratio=5:attack=20:release=300:makeup=1[musicd]")

    sl = []
    for i, s in enumerate(SFX):
        f, t, g = s[:3]
        if os.path.basename(f) in ONSET:
            t -= ONSET[os.path.basename(f)] - (s[3] if len(s) > 3 else 0)
        n = add(os.path.join(A, f))
        trim = f"atrim={s[3]}:{s[4]},asetpts=PTS-STARTPTS,afade=t=out:st={max(0.0, s[4] - s[3] - 0.08):.2f}:d=0.08," if len(s) > 3 else ""
        ch.append(f"[{n}:a]aresample=48000,aformat=channel_layouts=stereo,{trim}volume={g}dB,adelay={int(t * 1000)}|{int(t * 1000)}[s{i}]")
        sl.append(f"[s{i}]")
    ch.append(f"{''.join(sl)}amix=inputs={len(sl)}:normalize=0[sfx]")
    ch.append(f"[musicd][voice][sfx]amix=inputs=3:normalize=0,atrim=0:{DUR},loudnorm=I=-14:TP=-1.0:LRA=10,aresample=48000[out]")
    subprocess.run(["ffmpeg", "-hide_banner", "-loglevel", "error", "-y", *inp, "-filter_complex", ";".join(ch),
                    "-map", "[out]", "-ac", "2", "-c:a", "pcm_s16le", OUT], check=True)


if __name__ == "__main__":
    run()
