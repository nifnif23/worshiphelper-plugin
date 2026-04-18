# WorshipHelper — Speech Recognition Setup

The speech recognition has been upgraded from Windows' built-in System.Speech
to **Vosk** — a free, fully offline engine with significantly better accuracy,
especially for unusual words (Habakkuk, Thessalonians, Zechariah, etc.) and
non-US accents.

---

## Model recommendation

**Use the large model (`vosk-model-en-us-0.22`, ~1.8 GB).**

It is meaningfully more accurate for accented speech and reduces phonetic
misfires (e.g. "eight" being heard as "amos"). On modern hardware (8 GB+ RAM,
SSD) the cold load time is ~2–3 seconds — negligible for a service context.

The small model (`vosk-model-small-en-us-0.15`, ~50 MB) is available as a
fallback for low-spec machines, but is not recommended for regular use.

---

## If you build via GitHub Actions (normal)

Edit `.github/workflows/main.yml` and make sure the model URL points to the
large model:

```yaml
$modelUrl = "https://alphacephei.com/vosk/models/vosk-model-en-us-0.22.zip"
```

Then run the workflow as normal. The MSI will be larger (~1.8 GB extra) but
accuracy will be significantly better.

---

## If you build locally

You need to download the model once and put it in the right place before building.

**1. Download the large model** from https://alphacephei.com/vosk/models:

| Model | Size | Recommendation |
|---|---|---|
| `vosk-model-en-us-0.22` | ~1.8 GB | **Recommended** |
| `vosk-model-small-en-us-0.15` | ~50 MB | Low-spec fallback only |

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
- Check the debug panel (Monitor button) to see what Vosk is actually hearing

**Recognition works but references aren't detected**
Vosk heard something but the Bible reference detector filtered it out.
Use the debug panel to see the raw recognised text. Common causes:
- Book name too mangled (try speaking more slowly/clearly)
- No chapter/verse after the book name
- Combined confidence below threshold (check the log)

---

## What changed in this update

**`SpeechListener.cs`** — large model recommended; pre-warm call added so the
first real utterance doesn't stutter on model JIT initialisation.

**`PhoneticCorrector.cs`** — new Pass 5: "amos" in non-leading position is
corrected to "eight" (Vosk was consistently mapping /eɪt/ → "amos").

**`SpeechListener.cs` grammar** — all short abbreviations removed (gen, exod,
jn, rom, etc.). Fewer vocabulary targets means fewer false Vosk landings.

**`TestRibbonItem.cs`** — speech-to-UI marshalling switched from
`Control.BeginInvoke` to `SynchronizationContext.Post`, fixing the bug where
references weren't inserted when the Monitor panel was closed. Template
selection for speech insertion is now independent — use the new "Set Template"
ribbon button in the Speech group.
