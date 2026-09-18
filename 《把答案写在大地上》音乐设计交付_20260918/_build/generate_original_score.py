from __future__ import annotations

import csv
import hashlib
import json
import math
import wave
from pathlib import Path

import numpy as np


SR = 48_000
ROOT = Path(r"D:\lzy\微电影\《把答案写在大地上》音乐设计交付_20260918")
OUT = ROOT / "08_原创音乐成品_20260918"
MASTER_DIR = OUT / "01_主混音_WAV_48k24bit"
STEMS_DIR = OUT / "02_分轨_STEMS_48k24bit"
PREVIEW_DIR = OUT / "03_试听_PREVIEW_16bit"
PROJECT_DIR = OUT / "04_工程与主题"
DOC_DIR = OUT / "05_CUE表与技术报告"

STEM_KEYS = ("melody", "strings", "rhythm", "ambience")
STEM_LABELS = {
    "melody": "MELODY_钢琴独奏",
    "strings": "STRINGS_弦乐",
    "rhythm": "RHYTHM_节奏",
    "ambience": "AMBIENCE_氛围",
}


CUES = [
    {
        "id": "M01", "title": "一个问号", "duration": 40.0, "bpm": 58,
        "key": "D minor / open fifth", "shots": "04-07",
        "timeline": "约00:28-01:08（以锁画版为准）",
        "role": "建立未知、孤独与问题意识；从环境声中渐显，在推门前留出空间。",
    },
    {
        "id": "M02", "title": "走出去", "duration": 56.0, "bpm": 72,
        "key": "D Dorian", "shots": "07-12",
        "timeline": "约01:00-01:56（以锁画版为准）",
        "role": "把迟疑转成行动，以轻木质脉冲推动场景连接，结尾让位于钟声。",
    },
    {
        "id": "M03", "title": "旧物会说话", "duration": 74.0, "bpm": 56,
        "key": "A minor → C major colour", "shots": "13-19",
        "timeline": "约01:52-03:06（以锁画版为准）",
        "role": "承载旧物、记忆与口述历史；克制、温暖但不煽情。",
    },
    {
        "id": "M04", "title": "院子里的答案", "duration": 54.0, "bpm": 68,
        "key": "G major / Lydian colour", "shots": "20-24",
        "timeline": "约03:02-03:56（以锁画版为准）",
        "role": "人物关系变暖，答案第一次具有现实形状；保持对白清晰。",
    },
    {
        "id": "M05", "title": "钥匙与答案", "duration": 64.0, "bpm": 64,
        "key": "D minor → F major", "shots": "25-30",
        "timeline": "约03:52-04:56（以锁画版为准）",
        "role": "从抉择、停顿走向全片情绪峰值；关键台词前后主动收束。",
    },
    {
        "id": "M06", "title": "带回来", "duration": 65.0, "bpm": 68,
        "key": "D major add9", "shots": "31-36",
        "timeline": "约04:55-06:00（以锁画版为准）",
        "role": "主主题明亮回归；中段为片中播放完全让位，结尾自然落地。",
    },
]


def midi_freq(note: float) -> float:
    return 440.0 * (2.0 ** ((note - 69.0) / 12.0))


def adsr(n: int, attack: float, release: float) -> np.ndarray:
    env = np.ones(n, dtype=np.float32)
    a = min(n, max(1, int(attack * SR)))
    r = min(n, max(1, int(release * SR)))
    env[:a] = np.linspace(0.0, 1.0, a, dtype=np.float32)
    env[-r:] *= np.linspace(1.0, 0.0, r, dtype=np.float32)
    return env


def soft_piano(note: float, duration: float, velocity: float, rng: np.random.Generator) -> np.ndarray:
    n = max(1, int(duration * SR))
    t = np.arange(n, dtype=np.float32) / SR
    f = midi_freq(note)
    detune = rng.uniform(-0.0012, 0.0012)
    body = (
        np.sin(2 * np.pi * f * (1 + detune) * t) * np.exp(-t / 2.8)
        + 0.36 * np.sin(2 * np.pi * 2.01 * f * t + 0.12) * np.exp(-t / 1.4)
        + 0.15 * np.sin(2 * np.pi * 3.03 * f * t + 0.31) * np.exp(-t / 0.75)
        + 0.055 * np.sin(2 * np.pi * 5.08 * f * t) * np.exp(-t / 0.34)
    )
    attack_n = min(n, int(0.035 * SR))
    attack_noise = np.zeros(n, dtype=np.float32)
    if attack_n:
        noise = rng.normal(0.0, 1.0, attack_n).astype(np.float32)
        noise = np.cumsum(noise)
        noise /= np.max(np.abs(noise)) + 1e-9
        attack_noise[:attack_n] = 0.035 * noise * np.linspace(1.0, 0.0, attack_n, dtype=np.float32)
    return ((body + attack_noise) * adsr(n, 0.012, min(0.7, duration * 0.35)) * velocity).astype(np.float32)


def string_tone(note: float, duration: float, velocity: float, rng: np.random.Generator) -> np.ndarray:
    n = max(1, int(duration * SR))
    t = np.arange(n, dtype=np.float32) / SR
    f = midi_freq(note)
    vibrato = 1.0 + 0.0028 * np.sin(2 * np.pi * (4.7 + rng.uniform(-0.25, 0.25)) * t)
    phase = 2 * np.pi * f * np.cumsum(vibrato, dtype=np.float64) / SR
    sig = np.zeros(n, dtype=np.float64)
    for h, a in ((1, 1.0), (2, 0.27), (3, 0.16), (4, 0.07), (5, 0.04)):
        sig += a * np.sin(h * phase + rng.uniform(0, 2 * np.pi))
    breath = rng.normal(0.0, 0.008, n)
    env = adsr(n, min(1.8, duration * 0.28), min(2.4, duration * 0.34))
    return ((sig + breath) * env * velocity * 0.53).astype(np.float32)


def cello_tone(note: float, duration: float, velocity: float, rng: np.random.Generator) -> np.ndarray:
    n = max(1, int(duration * SR))
    t = np.arange(n, dtype=np.float32) / SR
    f = midi_freq(note)
    vib = 0.0035 * np.sin(2 * np.pi * 5.1 * t) * np.minimum(1.0, t / 0.8)
    phase = 2 * np.pi * f * t + (f / 5.1) * vib
    sig = (
        np.sin(phase) + 0.42 * np.sin(2 * phase + 0.1)
        + 0.22 * np.sin(3 * phase + 0.4) + 0.09 * np.sin(4 * phase)
    )
    bow = rng.normal(0.0, 0.018, n).astype(np.float32)
    return ((sig + bow) * adsr(n, 0.38, min(1.5, duration * 0.3)) * velocity * 0.58).astype(np.float32)


def clarinet_tone(note: float, duration: float, velocity: float, rng: np.random.Generator) -> np.ndarray:
    n = max(1, int(duration * SR))
    t = np.arange(n, dtype=np.float32) / SR
    f = midi_freq(note)
    phase = 2 * np.pi * f * t + 0.02 * np.sin(2 * np.pi * 4.2 * t)
    sig = np.sin(phase) + 0.33 * np.sin(3 * phase) + 0.10 * np.sin(5 * phase)
    return (sig * adsr(n, 0.18, min(0.8, duration * 0.25)) * velocity * 0.5).astype(np.float32)


def glass_tone(note: float, duration: float, velocity: float, rng: np.random.Generator) -> np.ndarray:
    n = max(1, int(duration * SR))
    t = np.arange(n, dtype=np.float32) / SR
    f = midi_freq(note)
    sig = (
        np.sin(2 * np.pi * f * t) * np.exp(-t / 2.0)
        + 0.42 * np.sin(2 * np.pi * f * 2.73 * t + 0.3) * np.exp(-t / 0.85)
        + 0.18 * np.sin(2 * np.pi * f * 4.17 * t + 0.8) * np.exp(-t / 0.45)
    )
    return (sig * adsr(n, 0.003, min(0.55, duration * 0.4)) * velocity * 0.55).astype(np.float32)


def wooden_pulse(note: float, duration: float, velocity: float, rng: np.random.Generator) -> np.ndarray:
    n = max(1, int(duration * SR))
    t = np.arange(n, dtype=np.float32) / SR
    f = midi_freq(note)
    sig = np.sin(2 * np.pi * f * t) * np.exp(-t / 0.18)
    sig += 0.28 * np.sin(2 * np.pi * 2.4 * f * t) * np.exp(-t / 0.07)
    noise = rng.normal(0.0, 1.0, n).astype(np.float32) * np.exp(-t / 0.025)
    return ((sig + 0.09 * noise) * velocity * 0.55).astype(np.float32)


INSTRUMENTS = {
    "piano": soft_piano,
    "strings": string_tone,
    "cello": cello_tone,
    "clarinet": clarinet_tone,
    "glass": glass_tone,
    "wood": wooden_pulse,
}


def pan(mono: np.ndarray, position: float) -> np.ndarray:
    position = float(np.clip(position, -1.0, 1.0))
    angle = (position + 1.0) * math.pi / 4.0
    return np.column_stack((mono * math.cos(angle), mono * math.sin(angle))).astype(np.float32)


def add_note(stem: np.ndarray, start: float, duration: float, note: float, velocity: float,
             instrument: str, position: float, rng: np.random.Generator, events: list[dict]) -> None:
    if start >= len(stem) / SR:
        return
    mono = INSTRUMENTS[instrument](note, duration, velocity, rng)
    begin = max(0, int(start * SR))
    end = min(len(stem), begin + len(mono))
    stem[begin:end] += pan(mono[:end - begin], position)
    events.append({
        "start": round(start, 3), "duration": round(duration, 3), "note": note,
        "velocity": round(velocity, 3), "instrument": instrument, "pan": position,
    })


def add_chord(stem: np.ndarray, start: float, duration: float, notes: list[int], velocity: float,
              instrument: str, rng: np.random.Generator, events: list[dict]) -> None:
    spread = np.linspace(-0.64, 0.64, len(notes))
    for i, note in enumerate(notes):
        add_note(stem, start + i * 0.035, duration, note, velocity / math.sqrt(len(notes)),
                 instrument, float(spread[i]), rng, events)


def ambience_track(duration: float, root_note: int, intensity: float,
                   rng: np.random.Generator) -> np.ndarray:
    n = int(duration * SR)
    t = np.arange(n, dtype=np.float32) / SR
    control_count = max(3, int(duration * 16))
    control = rng.normal(0.0, 1.0, control_count).astype(np.float32)
    slow_noise = np.interp(np.arange(n), np.linspace(0, n - 1, control_count), control).astype(np.float32)
    slow_noise /= np.max(np.abs(slow_noise)) + 1e-9
    f = midi_freq(root_note)
    drone = 0.42 * np.sin(2 * np.pi * f * t) + 0.2 * np.sin(2 * np.pi * 1.5 * f * t + 0.8)
    air = 0.32 * slow_noise + drone
    mono = air * adsr(n, 2.4, 3.2) * intensity
    left = mono * (0.83 + 0.10 * np.sin(2 * np.pi * 0.071 * t))
    right = mono * (0.83 + 0.10 * np.sin(2 * np.pi * 0.061 * t + 1.2))
    return np.column_stack((left, right)).astype(np.float32)


def add_reverb(stereo: np.ndarray, amount: float) -> np.ndarray:
    dry = stereo.astype(np.float32, copy=True)
    wet = np.zeros_like(dry)
    taps = ((0.071, 0.34), (0.113, 0.26), (0.173, 0.19), (0.241, 0.13), (0.337, 0.08))
    for delay, gain in taps:
        d = int(delay * SR)
        wet[d:, 0] += dry[:-d, 1] * gain
        wet[d:, 1] += dry[:-d, 0] * gain
    return (dry + wet * amount).astype(np.float32)


def automation(duration: float, points: list[tuple[float, float]]) -> np.ndarray:
    n = int(duration * SR)
    xs = np.asarray([p[0] for p in points], dtype=np.float32) * SR
    ys = np.asarray([p[1] for p in points], dtype=np.float32)
    return np.interp(np.arange(n), xs, ys).astype(np.float32)[:, None]


def schedule_piano_pattern(stem: np.ndarray, start: float, end: float, bpm: float,
                           pattern: list[int], velocity: float, rng: np.random.Generator,
                           events: list[dict], beat_step: float = 1.0) -> None:
    beat = 60.0 / bpm
    pos = start
    idx = 0
    while pos < end:
        note = pattern[idx % len(pattern)]
        add_note(stem, pos, min(3.2, end - pos + 1.0), note,
                 velocity * (1.0 if idx % 4 == 0 else 0.78), "piano",
                 -0.25 + 0.16 * (idx % 3), rng, events)
        pos += beat * beat_step
        idx += 1


def make_cue(cue: dict) -> tuple[dict[str, np.ndarray], list[dict]]:
    duration = cue["duration"]
    n = int(duration * SR)
    stems = {key: np.zeros((n, 2), dtype=np.float32) for key in STEM_KEYS}
    events: list[dict] = []
    rng = np.random.default_rng(20260918 + int(cue["id"][1:]))

    if cue["id"] == "M01":
        stems["ambience"] += ambience_track(duration, 38, 0.10, rng)
        motif = [(4.2, 62), (8.8, 64), (13.5, 65), (18.6, 64)]
        for i, (start, note) in enumerate(motif):
            add_note(stems["melody"], start, 4.4, note, 0.32 + 0.03 * i, "piano", -0.18, rng, events)
        for start, note in ((23.0, 50), (29.0, 53), (34.0, 52)):
            add_note(stems["melody"], start, 5.5, note, 0.18, "cello", 0.18, rng, events)
        for start, chord in ((19.0, [50, 57, 64]), (28.0, [53, 57, 62]), (34.0, [52, 57, 64])):
            add_chord(stems["strings"], start, 7.8, chord, 0.25, "strings", rng, events)
        curve = automation(duration, [(0, 0), (3, 0.7), (22, 0.9), (34, 0.8), (40, 0)])

    elif cue["id"] == "M02":
        stems["ambience"] += ambience_track(duration, 38, 0.075, rng)
        schedule_piano_pattern(stems["melody"], 0.8, 35.0, 72, [50, 57, 64, 65, 57, 62, 64, 69],
                               0.22, rng, events, 0.5)
        schedule_piano_pattern(stems["melody"], 36.0, 48.0, 72, [50, 57, 64, 69],
                               0.16, rng, events, 1.0)
        for start, chord in ((0, [50, 57, 64]), (8, [53, 59, 64]), (16, [55, 62, 69]),
                             (24, [50, 57, 65]), (32, [53, 59, 64]), (40, [55, 62, 69])):
            add_chord(stems["strings"], start, 9.2, chord, 0.21, "strings", rng, events)
        beat = 60 / 72
        pos = 1.0
        i = 0
        while pos < 43:
            add_note(stems["rhythm"], pos, 0.42, 38 if i % 4 == 0 else 45,
                     0.22 if i % 4 == 0 else 0.13, "wood", -0.35 if i % 2 == 0 else 0.35, rng, events)
            pos += beat
            i += 1
        for start, note in ((44.0, 74), (46.5, 76), (49.0, 77), (51.5, 76)):
            add_note(stems["rhythm"], start, 1.8, note, 0.15, "glass", 0.25, rng, events)
        curve = automation(duration, [(0, 0), (0.7, 0.8), (32, 0.88), (43, 0.66), (51, 0.35), (56, 0)])

    elif cue["id"] == "M03":
        stems["ambience"] += ambience_track(duration, 33, 0.105, rng)
        for start, chord in ((0, [45, 52, 60]), (12, [48, 55, 64]), (24, [43, 50, 59]),
                             (36, [45, 52, 60]), (48, [48, 55, 64]), (60, [43, 50, 59])):
            add_chord(stems["strings"], start, 14.0, chord, 0.235, "strings", rng, events)
        notes = [(3, 57), (10, 59), (17, 60), (25, 59), (36, 52), (44, 55),
                 (52, 57), (61, 60), (67, 59)]
        for i, (start, note) in enumerate(notes):
            add_note(stems["melody"], start, 5.0, note, 0.25 if i < 4 else 0.21,
                     "piano", -0.25, rng, events)
        for start, note, dur in ((18, 45, 9), (29, 48, 8), (42, 52, 10), (56, 55, 9), (65, 52, 7)):
            add_note(stems["melody"], start, dur, note, 0.22, "cello", 0.22, rng, events)
        for start, note in ((4, 81), (40, 79), (55, 84), (68, 81)):
            add_note(stems["rhythm"], start, 3.4, note, 0.11, "glass", 0.42, rng, events)
        curve = automation(duration, [(0, 0), (2, 0.72), (32, 0.80), (49, 0.88), (64, 0.68), (74, 0)])

    elif cue["id"] == "M04":
        stems["ambience"] += ambience_track(duration, 43, 0.075, rng)
        for start, chord in ((0, [43, 50, 59]), (9, [45, 52, 61]), (18, [47, 54, 62]),
                             (27, [48, 55, 62]), (36, [45, 52, 59]), (45, [43, 50, 59])):
            add_chord(stems["strings"], start, 10.5, chord, 0.245, "strings", rng, events)
        for start, note in ((1.2, 67), (5.0, 69), (8.8, 71), (13.0, 69),
                            (22, 71), (27, 74), (33, 73), (39, 71), (46, 69), (50, 67)):
            add_note(stems["melody"], start, 3.6, note, 0.235, "piano", -0.24, rng, events)
        for start, note, dur in ((15, 62, 6), (24, 66, 7), (32, 69, 7), (41, 67, 7)):
            add_note(stems["melody"], start, dur, note, 0.18, "clarinet", 0.28, rng, events)
        curve = automation(duration, [(0, 0), (1.2, 0.72), (24, 0.82), (38, 0.95), (47, 0.65), (54, 0)])

    elif cue["id"] == "M05":
        stems["ambience"] += ambience_track(duration, 38, 0.075, rng)
        for start, chord in ((0, [50, 57, 65]), (9, [46, 53, 62]), (18, [41, 48, 57]),
                             (27, [48, 55, 64]), (42, [50, 57, 65]), (49, [53, 60, 69]),
                             (55, [41, 48, 57]), (59, [53, 60, 69])):
            add_chord(stems["strings"], start, 10.0 if start < 42 else 7.2, chord,
                      0.24 if start < 42 else 0.31, "strings", rng, events)
        for start, note in ((1, 62), (5, 64), (9, 65), (14, 64), (20, 57), (25, 60),
                            (42, 62), (46, 65), (50, 69), (54, 72), (57, 69), (60, 65)):
            add_note(stems["melody"], start, 4.0, note, 0.24 if start < 42 else 0.31,
                     "piano", -0.22, rng, events)
        for start, note, dur in ((12, 50, 8), (22, 53, 8), (44, 57, 8), (52, 60, 8), (57, 65, 6)):
            add_note(stems["melody"], start, dur, note, 0.22 if start < 42 else 0.29,
                     "cello", 0.25, rng, events)
        beat = 60 / 64
        pos = 43.0
        i = 0
        while pos < 59.2:
            add_note(stems["rhythm"], pos, 0.45, 38 if i % 4 == 0 else 45,
                     0.14 + 0.006 * i, "wood", -0.3 if i % 2 == 0 else 0.3, rng, events)
            pos += beat
            i += 1
        curve = automation(duration, [(0, 0), (1.2, 0.75), (27, 0.85), (30, 0.22),
                                      (40, 0.18), (43, 0.72), (54, 1.0), (60, 0.70), (64, 0)])

    else:  # M06
        stems["ambience"] += ambience_track(duration, 38, 0.065, rng)
        for start, chord in ((0, [50, 57, 64, 66]), (9, [54, 61, 69]), (30, [50, 57, 64, 66]),
                             (39, [55, 62, 71]), (48, [54, 61, 69]), (56, [50, 57, 64, 66])):
            add_chord(stems["strings"], start, 11.0, chord, 0.235 if start < 30 else 0.27,
                      "strings", rng, events)
        for start, note in ((0.8, 62), (4.4, 64), (8.0, 66), (12.0, 69),
                            (30.5, 62), (34.0, 64), (37.5, 66), (41.0, 69),
                            (47.0, 71), (52.0, 69), (56.0, 66), (59.2, 74)):
            add_note(stems["melody"], start, 3.8, note, 0.24 if start < 30 else 0.285,
                     "piano", -0.22, rng, events)
        for start, note, dur in ((12, 57, 7), (38, 62, 7), (46, 66, 8), (54, 69, 8)):
            add_note(stems["melody"], start, dur, note, 0.18 if start < 30 else 0.21,
                     "clarinet", 0.26, rng, events)
        beat = 60 / 68
        pos = 31.0
        i = 0
        while pos < 52:
            add_note(stems["rhythm"], pos, 0.4, 38 if i % 4 == 0 else 45,
                     0.12, "wood", -0.28 if i % 2 == 0 else 0.28, rng, events)
            pos += beat
            i += 1
        curve = automation(duration, [(0, 0), (0.8, 0.75), (16, 0.80), (19.5, 0.12),
                                      (20, 0.04), (29.5, 0.04), (30.5, 0.72),
                                      (51, 0.92), (59, 0.78), (65, 0)])

    # Cue-level shaping is applied to all deliverable stems so that the split mix
    # reconstructs the master exactly before master normalization.
    for key in STEM_KEYS:
        reverb_amount = 0.46 if key in ("melody", "strings") else (0.22 if key == "rhythm" else 0.08)
        stems[key] = add_reverb(stems[key], reverb_amount) * curve
    return stems, events


def peak_normalize(stems: dict[str, np.ndarray], target_db: float = -1.0) -> tuple[dict[str, np.ndarray], np.ndarray, float]:
    master = sum(stems.values(), np.zeros_like(next(iter(stems.values()))))
    peak = float(np.max(np.abs(master)))
    target = 10 ** (target_db / 20.0)
    gain = target / max(peak, 1e-9)
    gain = min(gain, 5.0)
    scaled = {k: (v * gain).astype(np.float32) for k, v in stems.items()}
    master = sum(scaled.values(), np.zeros_like(master)).astype(np.float32)
    return scaled, master, gain


def write_pcm24(path: Path, stereo: np.ndarray) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with wave.open(str(path), "wb") as wf:
        wf.setnchannels(2)
        wf.setsampwidth(3)
        wf.setframerate(SR)
        block = SR * 5
        for start in range(0, len(stereo), block):
            data = np.clip(stereo[start:start + block], -1.0, 1.0)
            ints = np.rint(data.reshape(-1) * 8_388_607.0).astype("<i4")
            packed = ints.view(np.uint8).reshape(-1, 4)[:, :3]
            wf.writeframesraw(packed.tobytes())


def write_pcm16(path: Path, stereo: np.ndarray) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    rng = np.random.default_rng(918)
    with wave.open(str(path), "wb") as wf:
        wf.setnchannels(2)
        wf.setsampwidth(2)
        wf.setframerate(SR)
        block = SR * 5
        for start in range(0, len(stereo), block):
            data = np.clip(stereo[start:start + block], -1.0, 1.0)
            dither = rng.uniform(-1 / 65536, 1 / 65536, data.shape).astype(np.float32)
            ints = np.rint(np.clip(data + dither, -1.0, 1.0) * 32767.0).astype("<i2")
            wf.writeframesraw(ints.tobytes())


def dbfs(value: float) -> float:
    return 20.0 * math.log10(max(value, 1e-12))


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def write_documents(reports: list[dict]) -> None:
    overview = """# 《把答案写在大地上》原创音乐成品 v1.0

本目录包含 6 首为本片定制的原创器乐配乐，共约 5 分 53 秒。所有声音均由原创音频合成与编曲生成，未使用第三方歌曲、商业采样包或来源不明的录音素材。

## 文件夹

- `01_主混音_WAV_48k24bit`：剪辑与混音直接调用的立体声主混音。
- `02_分轨_STEMS_48k24bit`：每首四组立体声分轨；同一起点、同一长度，可无缝重建主混音。
- `03_试听_PREVIEW_16bit`：完整长度的 16-bit 试听副本。
- `04_工程与主题`：每首事件清单、速度、调性及主题设计资料。
- `05_CUE表与技术报告`：镜头调用表、技术检测、使用说明。

## 统一音乐语言

四音“问号动机”以 **D–E–F–E** 建立未解感，随剧情经历小调、调式与大调变形，片尾转为 **D–E–F♯–A**，表达“答案不是终点，而是带回现实的行动”。全片以毡音钢琴、柔弦、木质脉冲、独奏大提琴/单簧管与空气氛围为核心音色。

## 六首成品

| Cue | 名称 | 时长 | BPM | 调性色彩 | 对应镜头 | 主要功能 |
|---|---|---:|---:|---|---|---|
"""
    for cue in CUES:
        overview += f"| {cue['id']} | {cue['title']} | {cue['duration']:.0f}s | {cue['bpm']} | {cue['key']} | {cue['shots']} | {cue['role']} |\n"
    overview += """

## 使用优先级

1. 正片优先调用 `MASTER_v01.wav`。
2. 对白密集处使用 STEMS 降低 `MELODY` 与 `STRINGS`，保留 `AMBIENCE` 维持连续性。
3. M05 的 30–40 秒、M06 的 20–30 秒已做主动留白；不要用其他音乐填满。
4. 最终响度以整片混录为准；单曲已控制峰值不高于 -1 dBFS，但未替代整片对白/音效总线的最终响度校准。
"""
    (OUT / "00_成品总览.md").write_text(overview, encoding="utf-8-sig")

    usage = """# 音乐调用与混音说明

## 交付规格

- 主混音与分轨：48,000 Hz、24-bit PCM、立体声 WAV。
- 试听版：48,000 Hz、16-bit PCM、立体声 WAV（带微量抖动）。
- 所有分轨均从 00:00:00 起始，长度与对应 MASTER 完全一致。
- MASTER 峰值统一归一至约 -1 dBFS；整片最终响度应在对白、音效齐备后再校准。

## 分轨说明

- `MELODY_钢琴独奏`：钢琴与大提琴/单簧管/玻璃音等叙事性声部。
- `STRINGS_弦乐`：和声、长音与情绪托底。
- `RHYTHM_节奏`：木质脉冲与极少量标点性玻璃音。
- `AMBIENCE_氛围`：低频空气、调性底色与空间连续性。

## 剪辑规则

- 避免硬切旋律中段；若镜头时长变化，优先在和弦交界或留白区剪切。
- 对白出现时，先自动化降低 MELODY 2–5 dB，再视需要降低 STRINGS 1–3 dB。
- 关键对白前 6–12 帧可快速收弱，台词结束后用 12–24 帧回升。
- 片中钟声、旧物细节声、钥匙声与片中播放声优先于音乐。
- M06 20–30 秒为设计性近静默，供镜头33片中播放，不应补满。

## 版权与署名

本批音乐为本项目定制的原创程序化器乐配乐，不含第三方歌曲和外部采样。建议片尾署名：

> 原创音乐设计与制作：项目原创配乐（AI/程序化音频辅助）

如项目发行平台要求披露生成方式，可保留括号内容；具体权属与发行条款仍以制作方所在地法律、使用的制作流程及平台规则为准。
"""
    (DOC_DIR / "README_音乐调用与混音说明.md").write_text(usage, encoding="utf-8-sig")

    theme = """# 四音主题与变形说明

- 原始“问号动机”：D4–E4–F4–E4（未解决的小二度回落）。
- M01：稀疏陈述，保留不确定感。
- M02：置于 D Dorian 的行进型分解织体中，产生行动感。
- M03：移位至 A minor/C major 色彩，成为记忆线索。
- M04：G major/Lydian 色彩，关系逐渐打开。
- M05：D minor 中加宽音程，高潮转向 F major。
- M06：变为 D4–E4–F♯4–A4，最后落向 D major add9，完成主题释义。

事件 JSON 中的 `note` 使用 MIDI 音高编号；每条事件包含起始秒数、时长、力度、音色和声像，可作为重新配器或 MIDI 重建依据。
"""
    (PROJECT_DIR / "四音问号主题_音高与变形说明.md").write_text(theme, encoding="utf-8-sig")

    with (DOC_DIR / "镜头音乐调用表.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Cue", "音乐名称", "对应镜头", "建议时间轴", "时长秒", "BPM", "调性", "主混音文件", "作用"])
        for cue in CUES:
            filename = f"{cue['id']}_{cue['title']}_MASTER_v01.wav"
            writer.writerow([cue["id"], cue["title"], cue["shots"], cue["timeline"], cue["duration"],
                             cue["bpm"], cue["key"], filename, cue["role"]])

    (DOC_DIR / "音频技术报告.json").write_text(
        json.dumps({"generated": "2026-09-18", "sample_rate": SR, "reports": reports},
                   ensure_ascii=False, indent=2), encoding="utf-8-sig")


def main() -> None:
    for directory in (MASTER_DIR, STEMS_DIR, PREVIEW_DIR, PROJECT_DIR, DOC_DIR):
        directory.mkdir(parents=True, exist_ok=True)

    reports = []
    for cue in CUES:
        print(f"Rendering {cue['id']} {cue['title']}...", flush=True)
        stems, events = make_cue(cue)
        stems, master, gain = peak_normalize(stems)
        base = f"{cue['id']}_{cue['title']}"
        master_path = MASTER_DIR / f"{base}_MASTER_v01.wav"
        preview_path = PREVIEW_DIR / f"{base}_PREVIEW_16bit.wav"
        write_pcm24(master_path, master)
        cue_stem_dir = STEMS_DIR / base
        for key, audio in stems.items():
            write_pcm24(cue_stem_dir / f"{base}_{STEM_LABELS[key]}_v01.wav", audio)
        write_pcm16(preview_path, master)
        event_doc = {
            "cue": cue, "sample_rate": SR, "render_gain": gain,
            "event_count": len(events), "events": sorted(events, key=lambda x: x["start"]),
        }
        (PROJECT_DIR / f"{base}_events.json").write_text(
            json.dumps(event_doc, ensure_ascii=False, indent=2), encoding="utf-8-sig")
        peak = float(np.max(np.abs(master)))
        rms = float(np.sqrt(np.mean(master.astype(np.float64) ** 2)))
        reports.append({
            "cue": cue["id"], "title": cue["title"], "file": str(master_path.relative_to(OUT)),
            "duration_seconds": len(master) / SR, "channels": 2, "bit_depth": 24,
            "sample_rate": SR, "peak_dbfs": round(dbfs(peak), 3),
            "rms_dbfs": round(dbfs(rms), 3), "sha256": sha256(master_path),
            "stems": len(stems), "event_count": len(events),
        })
        del stems, master

    write_documents(reports)
    print(json.dumps(reports, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
