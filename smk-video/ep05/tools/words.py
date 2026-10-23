import sys, json, wave, glob, os, subprocess
import numpy as np
from faster_whisper import WhisperModel
m = WhisperModel("medium", device="cpu", compute_type="int8", cpu_threads=4)
out = {}
for f in sorted(glob.glob(sys.argv[1] + "/IMG_*.wav")):
    raw = subprocess.run(["ffmpeg", "-loglevel", "error", "-i", f, "-ac", "1", "-ar", "16000", "-f", "s16le", "-"], capture_output=True).stdout
    a = np.frombuffer(raw, np.int16).astype(np.float32) / 32768
    segs, _ = m.transcribe(a, language="ru", word_timestamps=True, beam_size=5)
    w = [[round(x.start, 2), round(x.end, 2), x.word.strip()] for s in segs for x in s.words]
    out[os.path.basename(f)[:-4]] = w
    print(os.path.basename(f), " ".join(f"{x[2]}[{x[0]}-{x[1]}]" for x in w), flush=True)
json.dump(out, open(sys.argv[2], "w"), ensure_ascii=False, indent=1)
