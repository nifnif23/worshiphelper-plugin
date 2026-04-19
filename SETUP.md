# WorshipHelper — Speech Recognition Setup

The speech recognition runs on **Vosk** — a free, fully offline engine.
No cloud, no cost, and it works without an internet connection.

---

## Model recommendation

**Use the small model (`vosk-model-small-en-us-0.15`, ~50 MB).**

The small model is fast to load, small enough to ship inside the MSI, and
combined with the two-phase grammar (SCAN → FOCUS-per-book) and the
PhoneticCorrector rationaliser it handles typical Bible references very
well. The large model (`vosk-model-en-us-0.22`, ~1.8 GB) gives slightly
better accuracy on unusual words but is rarely worth the extra size.

---

## If you build via GitHub Actions (normal)

The `main.yml` workflow already points to the small model:

```yaml
VOSK_MODEL_URL: https://alphacephei.com/vosk/models/vosk-model-small-en-us-0.15.zip
VOSK_MODEL_CACHE_KEY: vosk-model-small-en-us-0.15
```

Run the workflow as normal.

---

## If you build locally

You need to download the model once and put it in the right place before building.

**1. Download the small model** from https://alphacephei.com/vosk/models:

| Model | Size | Notes |
|---|---|---|
| `vosk-model-small-en-us-0.15` | ~50 MB | **Recommended — bundled by the workflow** |
| `vosk-model-en-us-0.22` | ~1.8 GB | Optional — better accuracy, much bigger MSI |

**2. Extract it** to:
```
WorshipHelperVSTO\data\vosk-model\
```
The folder should contain the model files directly (not nested inside another folder):
```
WorshipHelperVSTO\data\vosk-model\
    am\
    conf\
    graph\
    ...
```

**3. Install NuGet packages** — open Package Manager Console and run:
```
Install-Package Vosk -Version 0.3.38
Install-Package NAudio -Version 2.2.1
```

**4. Build normally.**

---

## Troubleshooting

**"Vosk model not found" on startup**
The model isn't where the add-in expects it. It looks first next to the DLL
(`data\vosk-model\`), then in `%APPDATA%\WorshipHelper\vosk-model\`.
Check the WorshipHelper log for the exact path it tried.

**No recognition / nothing happens**
- Check Windows Sound settings — correct mic set as default input?
- Speak clearly and include book + chapter + verse: *"John three sixteen"*
- Open the Speech Monitor (Monitor button in the ribbon) to see:
  * The live mic-level bar (green activity = the mic is picking up audio)
  * The last heard utterance and its confidence
  * The current recognition phase (SCAN or FOCUS: <book>)

**Recognition works but references aren't detected**
Vosk heard something but the Bible reference detector filtered it out.
Use the Speech Monitor to see the raw recognised text. Common causes:
- Book name too mangled (try speaking more slowly/clearly)
- No chapter/verse after the book name
- Combined confidence below threshold (check the log)

---

## What's in this build

**Small model + rationaliser pipeline**
- Phase 1 (SCAN): broad grammar of book names + numbers + a few structure
  words (`verse`, `chapter`, `to`, etc.) + common mishearings.
- Phase 2 (FOCUS): once a book name fires, the grammar collapses to just
  that book's variants + numbers — so chapter/verse come out cleanly.
- PhoneticCorrector rewrites known Vosk mishearings before detection
  (e.g. `tea` → `three`, `heaven` → `seven`, `of us` → `verse`).

**Speech Monitor (diagnostic panel)**
- Live mic-level bar — confirms the add-in is actually receiving audio.
- Per-result confidence bar — red / yellow / green at 50 % / 75 %.
- Phase badge in the top-right — `SCAN` or `FOCUS: <book>`.
- Copy button — copies the last detected reference to the clipboard.
- Log auto-trims at 500 lines so it never grows without bound.
