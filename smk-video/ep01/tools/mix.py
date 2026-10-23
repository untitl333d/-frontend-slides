"""Premix of the episode soundtrack: music bed + pitched voice + SFX -> media/mix.wav.

All times are video seconds. The music drop (track 31.835s) lands on video 2.0s.
"""
import subprocess, sys, os, wave
import numpy as np

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A = os.path.join(ROOT, "..", "assets")
OUT = os.path.join(ROOT, "media", "mix.wav")
DUR = 32.46
MUSIC_OFFSET = 29.835          # track time at video 0
LP_ZONES = [(0.0, 2.0), (22.06, 29.49)]   # filtered music ("from the radio" / breakdown)

# (file, video_start, file_in, file_out, gain_db)
VOICE = [
    ("prepped/voice_9875.wav", 2.00, 0.25, 3.25, 0),
    ("prepped/voice_9876.wav", 7.34, 0.70, 7.46, 0),
    ("prepped/voice_9877.wav", 15.51, 0.70, 6.71, 0),
    ("prepped/voice_9879.wav", 22.06, 0.30, 4.09, 0),
]
# (sfx, hit time, gain_db, peak offset inside the file[, max length s]) -- glitches are clipped short:
# played in full they turn into seconds of broadband hiss under the voice
SFX = [
    ("glitch-1", 0.05, -14, 0.0, 0.5), ("glitch-3", 1.05, -16, 0.0, 0.4),
    ("impact-bass-1", 2.00, -9, 0.0),
    ("whoosh", 4.97, -13, 0.12), ("whoosh", 5.71, -14, 0.12), ("whoosh-short", 6.46, -15, 0.10),
    ("glitch-1", 7.20, -16, 0.0, 0.35),
    ("impact-bass-2", 8.20, -12, None),
    ("pop", 9.00, -16, 0.0), ("pop", 10.45, -16, 0.0),
    ("whoosh", 12.52, -13, 0.12), ("error", 13.60, -21, 0.0, 0.6),
    ("whoosh-short", 15.37, -12, 0.10),
    ("whoosh-short", 15.75, -15, 0.10), ("notification", 17.03, -16, 0.0), ("typing", 18.13, -19, 0.0),
    ("ping", 20.05, -17, 0.0),
    ("glitch-3", 22.06, -15, 0.0, 0.4), ("error", 23.60, -21, 0.0, 0.6), ("impact-bass-1", 24.29, -12, 0.0),
    ("click", 25.03, -12, 0.0), ("click", 25.78, -12, 0.0), ("click", 26.52, -12, 0.0),
    ("impact-bass-2", 29.49, -9, None), ("whoosh", 29.49, -13, 0.12),
    ("typing", 29.85, -18, 0.0), ("ping", 31.00, -17, 0.0),
]

def peak_time(path):
    raw = subprocess.run(["ffmpeg", "-loglevel", "error", "-i", path, "-ac", "1", "-ar", "8000", "-f", "s16le", "-"],
                         capture_output=True, check=True).stdout
    a = np.abs(np.frombuffer(raw, np.int16).astype(float))
    win = 160
    env = np.convolve(a, np.ones(win) / win, "same")
    return float(np.argmax(env)) / 8000

def run():
    inputs, chains, labels = [], [], []
    # music: dry + lowpassed copy, crossfaded by volume envelopes
    inputs += ["-ss", f"{MUSIC_OFFSET}", "-t", f"{DUR}", "-i", os.path.join(A, "music/titanium-instrumental.mp3")]
    lp_expr = "+".join(f"between(t,{a},{b})" for a, b in LP_ZONES)
    chains.append("[0:a]aresample=48000,asplit=2[mdry][mlpin]")
    chains.append("[mlpin]lowpass=f=800,lowpass=f=800,volume=-7dB[mlp]")
    chains.append(f"[mdry]volume='1-min(1,({lp_expr}))':eval=frame[mdryg]")
    chains.append(f"[mlp]volume='min(1,({lp_expr}))':eval=frame[mlpg]")
    chains.append("[mdryg][mlpg]amix=inputs=2:normalize=0,afade=t=out:st=31.2:d=1.26,volume=-7dB[music]")
    # voice bus
    vlabels = []
    for i, (f, start, fin, fout, g) in enumerate(VOICE):
        inputs += ["-i", os.path.join(A, f)]
        n = inputs.count("-i") - 1
        chains.append(f"[{n}:a]atrim={fin}:{fout},asetpts=PTS-STARTPTS,afftdn=nr=10:nf=-40,deesser=i=0.5:m=0.5:f=0.5,"
                      f"agate=threshold=0.012:range=0.25:ratio=2:attack=5:release=150,afade=t=in:d=0.03,afade=t=out:st={fout-fin-0.05}:d=0.05,"
                      f"volume={g}dB,adelay={int(start*1000)}|{int(start*1000)}[v{i}]")
        vlabels.append(f"[v{i}]")
    chains.append(f"{''.join(vlabels)}amix=inputs={len(vlabels)}:normalize=0,apad=whole_dur={DUR},asplit=2[voice][vkey]")
    # duck music under voice
    chains.append("[music][vkey]sidechaincompress=threshold=0.02:ratio=6:attack=15:release=250:makeup=1[musicd]")
    # sfx
    slabels = []
    peaks = {}
    for i, (name, t, g, off, *maxlen) in enumerate(SFX):
        path = os.path.join(A, "sfx", f"{name}.mp3")
        if off is None:
            off = peaks.setdefault(name, peak_time(path))
        st = max(0.0, t - off)
        inputs += ["-i", path]
        n = inputs.count("-i") - 1
        clip = f"atrim=0:{maxlen[0]},afade=t=out:st={maxlen[0] * 0.5}:d={maxlen[0] * 0.5}," if maxlen else ""
        chains.append(f"[{n}:a]aresample=48000,aformat=channel_layouts=stereo,{clip}volume={g}dB,adelay={int(st*1000)}|{int(st*1000)}[s{i}]")
        slabels.append(f"[s{i}]")
    chains.append(f"{''.join(slabels)}amix=inputs={len(slabels)}:normalize=0[sfx]")
    chains.append(f"[musicd][voice][sfx]amix=inputs=3:normalize=0,atrim=0:{DUR},"
                  "loudnorm=I=-14:TP=-1.0:LRA=9,aresample=48000[out]")
    cmd = ["ffmpeg", "-hide_banner", "-loglevel", "error", "-y", *inputs,
           "-filter_complex", ";".join(chains), "-map", "[out]", "-ac", "2", "-c:a", "pcm_s16le", OUT]
    subprocess.run(cmd, check=True)
    print("peaks", peaks)

if __name__ == "__main__":
    run()
