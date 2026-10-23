"""Episode 5 voiceover: per-take EQ match to the median spectrum, pitch -2 st (series voice),
one shared chain (cleanup, presence, de-ess, compression), then each take normalised to -16 LUFS."""
import subprocess, json, glob, os
import numpy as np

V = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "assets", "ep05", "voice")
OUT = os.path.join(V, "processed")
TAKES = ["IMG_1631", "IMG_1635", "IMG_1637", "IMG_1642", "IMG_1645", "IMG_1647", "IMG_1651"]
BANDS = [(200, 100, 300), (600, 300, 1000), (1800, 1000, 3000), (5000, 3000, 8000)]


def spectrum(path):
    raw = subprocess.run(["ffmpeg", "-loglevel", "error", "-i", path, "-vn", "-ac", "1", "-ar", "16000", "-f", "s16le", "-"],
                         capture_output=True, check=True).stdout
    x = np.frombuffer(raw, np.int16).astype(float) / 32768
    n = 2048
    fr = np.lib.stride_tricks.sliding_window_view(x, n)[::1024]
    rms = np.sqrt((fr ** 2).mean(1))
    fr = fr[rms > np.percentile(rms, 60)]  # speech frames only
    S = (np.abs(np.fft.rfft(fr * np.hanning(n), axis=1)) ** 2).mean(0)
    f = np.fft.rfftfreq(n, 1 / 16000)
    return np.array([10 * np.log10(S[(f >= a) & (f < b)].mean()) for _, a, b in BANDS])


def loudnorm(src, dst, target=-16.0):
    r = subprocess.run(["ffmpeg", "-hide_banner", "-i", src, "-af", f"loudnorm=I={target}:TP=-1.5:LRA=7:print_format=json",
                        "-f", "null", "-"], capture_output=True, text=True).stderr
    m = json.loads(r[r.rindex("{"):r.rindex("}") + 1])
    af = (f"loudnorm=I={target}:TP=-1.5:LRA=7:measured_I={m['input_i']}:measured_TP={m['input_tp']}:"
          f"measured_LRA={m['input_lra']}:measured_thresh={m['input_thresh']}:offset={m['target_offset']}:linear=true")
    subprocess.run(["ffmpeg", "-loglevel", "error", "-y", "-i", src, "-af", af + ",aresample=48000", "-ac", "2", dst], check=True)


def run():
    os.makedirs(OUT, exist_ok=True)
    specs = {t: spectrum(os.path.join(V, t + ".MOV")) for t in TAKES}
    # compare spectral *shape*: remove each take's overall level first
    shapes = {t: s - s.mean() for t, s in specs.items()}
    target = np.median(np.array(list(shapes.values())), axis=0)
    for t in TAKES:
        corr = np.clip(target - shapes[t], -4, 4)
        eq = ",".join(f"equalizer=f={fc}:t=o:w=1.2:g={g:.2f}" for (fc, _, _), g in zip(BANDS, corr))
        chain = ("highpass=f=85,lowpass=f=13000,afftdn=nr=8:nf=-55," + eq + ","
                 "rubberband=pitch=0.8909:formant=preserved:pitchq=quality,"
                 "equalizer=f=320:t=q:w=1.0:g=-2,equalizer=f=3500:t=q:w=1.2:g=2.5,equalizer=f=110:t=q:w=0.8:g=1.5,"
                 "deesser=i=0.45:m=0.5:f=0.5,"
                 "acompressor=threshold=-24dB:ratio=3.5:attack=6:release=120:makeup=4:knee=6,"
                 "silenceremove=start_periods=1:start_threshold=-48dB:start_silence=0.08,"
                 "areverse,silenceremove=start_periods=1:start_threshold=-48dB:start_silence=0.15,areverse")
        tmp = os.path.join(OUT, t + "_tmp.wav")
        subprocess.run(["ffmpeg", "-loglevel", "error", "-y", "-i", os.path.join(V, t + ".MOV"), "-vn", "-ac", "1",
                        "-ar", "48000", "-af", chain, tmp], check=True)
        loudnorm(tmp, os.path.join(OUT, t + ".wav"))
        os.remove(tmp)
        print(t, "EQ corr", np.round(corr, 1))


if __name__ == "__main__":
    run()
