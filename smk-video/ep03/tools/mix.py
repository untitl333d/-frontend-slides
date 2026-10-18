"""Premix of the episode soundtrack: music bed + pitched voice + SFX -> media/mix.wav.

All times are video seconds. The music drop (track 31.835s) lands on video 2.0s.
"""
import subprocess, sys, os, wave
import numpy as np

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A = os.path.join(ROOT, "..", "assets")
OUT = os.path.join(ROOT, "media", "mix.wav")
DUR = 40.0
MUSIC_OFFSET = 29.835          # track time at video 0
LP_ZONES = [(0.0, 2.0), (21.318, 27.262)]   # filtered music ("from the radio" / own-money breakdown)
MUSIC_CUT = 35.437                          # music stops dead on the final black screen

# (file, video_start, file_in, file_out, gain_db)
VOICE = [
    ("prepped/voice_9893.wav", 2.30, 0.45, 4.10, 0),     # VO: hundreds of channels
    ("prepped/voice_9897.wav", 7.201, 0.60, 6.40, 0),    # head: not screenshots, a journal
    ("prepped/voice_9898.wav", 14.80, 0.60, 5.60, 0),    # VO: stop in every signal
    ("prepped/voice_9900.wav", 21.318, 0.90, 5.65, 0),   # head: we trade the same signals
    ("prepped/voice_9905.wav", 27.262, 0.70, 6.10, 0),   # head: CTA, up to "...на 3 дня,"
    ("prepped/voice_9905.wav", 32.662, 7.40, 8.45, 0),   # head: "Посмотришь сам." ("до 10 октября" cut out)
]
# (sfx, hit time, gain_db, peak offset inside the file[, max length s]) -- glitches are clipped short:
# played in full they turn into seconds of broadband hiss under the voice
SFX = [
    ("glitch-1", 0.05, -14, 0.0, 0.5), ("glitch-3", 1.05, -16, 0.0, 0.4),
    ("impact-bass-1", 2.00, -9, 0.0),
    ("pop", 2.45, -16, 0.0), ("pop", 2.75, -16, 0.0), ("pop", 3.01, -16, 0.0), ("whoosh", 3.43, -14, 0.12),
    ("ping", 5.21, -17, 0.0),
    ("whoosh-short", 5.55, -15, 0.10), ("whoosh-short", 5.80, -15, 0.10), ("whoosh-short", 6.05, -15, 0.10),
    ("impact-bass-1", 6.46, -11, 0.0),
    ("glitch-1", 7.05, -16, 0.0, 0.35),
    ("whoosh-short", 7.40, -15, 0.10), ("pop", 7.95, -16, 0.0), ("pop", 8.15, -16, 0.0),
    ("impact-bass-2", 8.98, -12, None),
    ("click", 11.22, -13, 0.0), ("click", 12.10, -13, 0.0),
    ("glitch-3", 14.50, -15, 0.0, 0.4),
    ("whoosh", 15.10, -14, 0.12), ("typing", 15.92, -19, 0.0, 0.75), ("impact-bass-1", 16.72, -12, 0.0),
    ("pop", 17.66, -17, 0.0), ("error", 19.22, -21, 0.0, 0.6),
    ("glitch-1", 21.20, -16, 0.0, 0.35),
    ("pop", 22.64, -17, 0.0), ("impact-bass-1", 24.22, -13, 0.0),
    ("glitch-3", 27.15, -15, 0.0, 0.4), ("typing", 28.76, -17, 0.0, 0.65),
    ("impact-bass-1", 29.32, -12, 0.0), ("pop", 29.78, -16, 0.0),
    ("whoosh", 32.04, -14, 0.12), ("glitch-1", 32.662, -15, 0.0, 0.35), ("ping", 32.70, -16, 0.0),
    ("click-soft", 35.55, -14, 0.0), ("click-soft", 35.95, -14, 0.0), ("click-soft", 36.35, -14, 0.0),
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
    chains.append(f"[mdryg][mlpg]amix=inputs=2:normalize=0,afade=t=out:st={MUSIC_CUT - 0.04}:d=0.04,volume=-7dB[music]")
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
