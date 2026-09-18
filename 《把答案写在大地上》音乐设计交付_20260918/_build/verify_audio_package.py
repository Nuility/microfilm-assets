import json
import wave
from pathlib import Path


root = Path(r"D:\lzy\微电影\《把答案写在大地上》音乐设计交付_20260918\08_原创音乐成品_20260918")
files = sorted(root.rglob("*.wav"))
rows = []
bad = []

for file in files:
    with wave.open(str(file), "rb") as audio:
        row = {
            "file": str(file.relative_to(root)),
            "channels": audio.getnchannels(),
            "sample_rate": audio.getframerate(),
            "bit_depth": audio.getsampwidth() * 8,
            "seconds": round(audio.getnframes() / audio.getframerate(), 3),
        }
    rows.append(row)
    expected_bits = 16 if "PREVIEW" in row["file"] else 24
    if row["channels"] != 2 or row["sample_rate"] != 48_000 or row["bit_depth"] != expected_bits:
        bad.append(row)

summary = {
    "wav_count": len(files),
    "masters": len(list((root / "01_主混音_WAV_48k24bit").glob("*.wav"))),
    "stems": len(list((root / "02_分轨_STEMS_48k24bit").rglob("*.wav"))),
    "previews": len(list((root / "03_试听_PREVIEW_16bit").glob("*.wav"))),
    "project_event_files": len(list((root / "04_工程与主题").glob("*_events.json"))),
    "total_bytes": sum(file.stat().st_size for file in root.rglob("*") if file.is_file()),
    "invalid_files": bad,
    "masters_detail": [row for row in rows if "MASTER" in row["file"]],
}

print(json.dumps(summary, ensure_ascii=False, indent=2))
