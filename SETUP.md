# WorshipHelper — Speech Recognition Setup

The speech recognition has been upgraded from Windows' built-in System.Speech
to **Vosk** — a free, fully offline engine with significantly better accuracy,
especially for unusual words (Habakkuk, Thessalonians, Zechariah, etc.) and
non-US accents.

---

## If you build via GitHub Actions (normal)

**Nothing extra to do.** The workflow automatically downloads the Vosk model
during the build and bundles it inside the MSI. Just run the workflow and
install the resulting MSI as usual.

The bundled model is `vosk-model-small-en-us-0.15` (~50 MB). It's fast and
good enough for structured phrases like Bible references. If you want the
larger, more accurate model, see "Upgrading the model" below.

---

## If you build locally

You need to download the model once and put it in the right place before building.

**1. Download a model** from https://alphacephei.com/vosk/models

| Model | Size | Notes |
|---|---|---|
| `vosk-model-small-en-us-0.15` | ~50 MB | Start here |
| `vosk-model-en-us-0.22` | ~1.8 GB | Best accuracy for Sunday use |

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
Install-Package Vosk -Version 0.3.45
Install-Package NAudio -Version 2.2.1
```

**4. Build normally.**

---

## Upgrading to the large model

The large model (`vosk-model-en-us-0.22`, ~1.8 GB) is noticeably more accurate
but makes the MSI much larger and slows the build by several minutes.

To switch, edit `.github/workflows/main.yml` and change the two lines:
```yaml
$modelUrl  = "https://alphacephei.com/vosk/models/vosk-model-small-en-us-0.15.zip"
```
to:
```yaml
$modelUrl  = "https://alphacephei.com/vosk/models/vosk-model-en-us-0.22.zip"
```

---

## Troubleshooting

**"Vosk model not found" on startup**
The model isn't where the add-in expects it. It looks first next to the DLL
(`data\vosk-model\`), then in `%APPDATA%\WorshipHelper\vosk-model\`.
Check the WorshipHelper log for the exact path it tried.

**No recognition / nothing happens**
- Check Windows Sound settings — correct mic set as default input?
- Speak clearly and include book + chapter + verse: *"John three sixteen"*
- Check the debug panel to see what Vosk is actually hearing

**Recognition works but references aren't detected**
Vosk heard something but the Bible reference detector filtered it out.
Use the debug panel to see the raw recognised text. Common causes:
- Book name too mangled (try speaking more slowly/clearly)
- No chapter/verse after the book name
- Combined confidence below threshold (check the log)

---

## What changed in this update

**`SpeechListener.cs`** — rewritten to use Vosk + NAudio instead of System.Speech.
Same public API, same events throughout the rest of the project.

**`BibleReferenceDetector.cs`** — three fixes:
- Removed `"of"` from the preamble skip list (was breaking "Song of Solomon")
- Fuzzy matching with edit distance scaled by word length, so long obscure
  names like Habakkuk, Zechariah, Thessalonians get enough slack to match
  even when the engine mishears them
- Small confidence penalty for fuzzy matches (still accepted, but ranked lower
  than exact matches if both fire)

**`SpokenNumberConverter.cs`** — fixed a bug where "twenty and six" would fail
to split into chapter 20, verse 6

**`.github/workflows/main.yml`** — added a step to download and bundle the
Vosk model automatically during the build
