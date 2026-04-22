# WorshipHelper v4 -- Faster-Whisper Upgrade

## Files in this package

```
python_server/
  requirements.txt              Python dependencies
  stt_engine.py                 Faster-Whisper wrapper
  server.py                     WebSocket STT server
  build_verse_embeddings.py     One-shot DB builder

WorshipHelperVSTO/
  SpeechListener.cs             Rewritten thin facade (same public API)
  Audio/
    MicrophoneCapture.cs        NAudio mic wrapper
    Chunker.cs                  PCM aggregator + silence gate
  Networking/
    PythonClient.cs             WebSocket client w/ auto-reconnect
  Detection/
    PatternMatcher.cs           Facade over BibleReferenceDetector
    VerseDatabase.cs            SQLite verse embedding loader
    SemanticSearch.cs           Cosine similarity search
  Feedback/
    FeedbackStore.cs            SQLite detection log
    CorrectionEngine.cs         Runtime threshold adjuster
  AddIn/
    Pipeline.cs                 Replaces ThisAddIn_SpeechIntegration.cs
  WorshipHelperVSTO.csproj      Updated (Vosk removed, new files added)
  packages.config               Updated (Vosk removed)
```

## Files to DELETE from existing project

- `ThisAddIn_SpeechIntegration.cs` (replaced by `AddIn/Pipeline.cs`)
- `data/vosk-model/` folder
- Remove Vosk entries from packages.config and CI HintPath injection

## First-time Python setup (on target machine)

```powershell
cd python_server
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

# One-shot: embed every verse (3-5 min GPU, 15 min CPU)
python build_verse_embeddings.py `
    --bible "..\WorshipHelperVSTO\data\Bibles\NASB.xmm" `
    --out   "..\WorshipHelperVSTO\data\verses.sqlite"

# Run the server (keep running while PowerPoint is open)
python server.py --model large-v3 --device auto
```

## Build the VSTO add-in

1. Copy new files into existing project at the paths shown above
2. Delete `ThisAddIn_SpeechIntegration.cs`
3. Replace `WorshipHelperVSTO.csproj` and `packages.config`
4. Add `data/verses.sqlite` to installer payload
5. Build as normal

## Autostart recommendation

Ship a `worshiphelper-stt.bat` + Windows scheduled task so the Python server
launches at login. The C# client auto-connects and auto-reconnects.
