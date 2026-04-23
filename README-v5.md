# WorshipHelper v5 — "PewBeam-grade" speech hardening

## The short version

Your v4 log showed this:

```
21:07:58  DEBUG seg texts: ['John 3 verse seven.']            <- you said this
21:08:06  DEBUG seg texts: ['John 3.','John 4.','John 5.',    <- all hallucinated
                            'John 6.','John 7.','John 8.','John 9.']
21:08:13  DEBUG seg texts: ['John 16, John 3.']
```

That is the textbook **"Whisper hallucinates on silence"** failure mode. It is
not a mystery and it is fixable. v5 rebuilds the stack so you can put it on
auto-scripture and relax.

---

## Why v4 hallucinated (root causes)

1. **Biased priming**: `stt_engine.py` told Whisper — on every single chunk —
   *"A Christian minister announces Bible references such as John 3:16, …"*.
   So whenever you stopped talking, Whisper filled the silence with the
   examples from that prompt.
2. **`condition_on_previous_text=True`**: Whisper carried its last output into
   the next window and just incremented the chapter number (John 3 → John 4 → …).
3. **Fixed 1.5-second windows**: the server transcribed every 1.5s of audio
   whether or not you were actually speaking.
4. **Sampling temperature of 0.2** on low-signal audio → synthesised content.
5. **`no_speech_threshold=0.8`** was way too permissive (Whisper's own default
   is 0.6; for short-utterance work 0.5 is saner).
6. **No compression-ratio filter / no blacklist / no dedup** — Whisper's own
   recommended hallucination safeguards were all off.
7. **Client-side silence gate at `0.003` RMS** — that passes room tone.

---

## What v5 does differently

### Python server (rewrite)

* **Utterance-based transcription**. Audio streams in; the server uses
  Silero VAD to carve out complete utterances (speech bracketed by ≥700 ms
  of silence) and transcribes each one exactly once.
* **No biasing prompt**. `initial_prompt=None`; no example references.
* **`condition_on_previous_text=False`**. Each utterance stands alone.
* **`beam_size=1` on CPU** (greedy): both faster *and* more accurate on
  short utterances.
* **Whisper's own hallucination safeguards ON**:
  * `no_speech_threshold=0.50`
  * `compression_ratio_threshold=2.4`
  * `log_prob_threshold=-1.0`
  * `hallucination_silence_threshold=2.0`
* **Post-decode filters**:
  * Hallucination phrase blacklist ("Thanks for watching", "Amen.",
    "Hallelujah" in isolation, …)
  * Same-prefix-different-number dedup ("John 3" vs "John 4" within 6 s → drop)
  * Exact-match dedup within 6 s
* **Hotwords** (book names) pushed per-session so the decoder knows those are
  real English words — *without* baking example references into the prompt.

### C# client

* **Smarter silence gate** (`Chunker.cs`):
  * RMS floor at `0.010` (≈ –40 dBFS, room tone fails)
  * Voiced-sub-frame ratio ≥ 15 % required (kills steady HVAC hums)
* **New `HallucinationGuard` class** — client-side defence-in-depth:
  * Rejects if `no_speech_prob > 0.35`
  * Rejects if `avg_logprob < –0.80`
  * Rejects if `compression_ratio > 2.0`
  * **Rejects rapid-fire same-book loops** (≥ 3 distinct chapters of the same
    book in 3 s — exactly the John 3/4/5/6/7 pattern)
  * Defers chapter-only references with low confidence + no intent word
    (`turn`, `read`, `chapter`, `verse`, …) so the existing 10-second grace
    window has time to promote them.
* **Local dedup cache** in `SpeechListener` (belt-and-braces after
  server-side dedup).
* **Metrics pipe**: `avg_logprob`, `compression_ratio`, `no_speech_prob`,
  `duration_seconds` flow all the way to the detection service so every
  stage can vote.
* **Confidence thresholds tightened**:
  * `MinEngineConfidence`: 0.10 → **0.30**
  * `MinCombinedConfidence`: 0.40 → **0.55**

---

## File map

```
python_server/
  stt_engine.py                REPLACE  -- hallucination-proof engine
  server.py                    REPLACE  -- utterance-based VAD protocol
  vad.py                       ADD      -- Silero VAD aggregator
  embed_server.py              ADD      -- companion /embed HTTP server
  requirements.txt             REPLACE  -- adds torch (Silero), pins ranges
  worshiphelper-stt.ps1        ADD      -- one-command launcher
  build_verse_embeddings.py    KEEP     -- unchanged

WorshipHelperVSTO/
  Audio/Chunker.cs             REPLACE  -- tighter silence/voiced-ratio gate
  Audio/MicrophoneCapture.cs   KEEP     -- unchanged
  Networking/PythonClient.cs   REPLACE  -- richer metrics, hotwords, ping
  Detection/HallucinationGuard.cs  ADD  -- NEW: rapid-fire + trust guard
  Detection/PatternMatcher.cs  KEEP     -- unchanged
  Detection/VerseDatabase.cs   KEEP     -- unchanged
  Detection/SemanticSearch.cs  KEEP     -- unchanged
  SpeechListener.cs            REPLACE  -- surfaces metrics, pushes hotwords
  SpeechToScriptureService.cs  REPLACE  -- wires the guard + tighter defaults
  BibleReferenceDetector.cs    KEEP     -- unchanged (already solid)
  PhoneticCorrector.cs         KEEP     -- unchanged
  SpokenNumberConverter.cs     KEEP     -- unchanged
  ReferenceValidator.cs        KEEP     -- unchanged
  (all other files)            KEEP     -- unchanged
```

No new NuGet packages. csproj only needs one added `<Compile Include=>`
line (`Detection\HallucinationGuard.cs`) — see `csproj-additions.txt`.

---

## Install & run

```powershell
cd python_server

# First time only (creates .venv, installs deps, downloads Silero + Whisper)
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

# Every time after that (keep this window open while PowerPoint is up)
.\worshiphelper-stt.ps1                 # auto-detect GPU/CPU, small.en
# or
.\worshiphelper-stt.ps1 -Model small.en -Device cpu
.\worshiphelper-stt.ps1 -Model distil-small.en   # 2-3× faster, tiny accuracy hit
.\worshiphelper-stt.ps1 -Model large-v3 -Device cuda   # best, if you have a GPU
```

The PowerShell script also starts the companion `embed_server.py` on
port 8766 for semantic search. Use `-NoEmbed` to skip that.

### Model recommendations

| Hardware                | Model          | Notes                                    |
|-------------------------|----------------|------------------------------------------|
| CPU only (laptop)       | `small.en`     | default, ~0.5 s latency per utterance    |
| CPU only (fast laptop)  | `distil-small.en` | 2–3× faster, ~same accuracy            |
| CPU (desktop, >8 cores) | `medium.en`    | better on accents, ~1–1.5 s latency      |
| GPU (4 GB+)             | `large-v3`     | best accuracy, ~0.3 s latency            |
| GPU (4 GB+)             | `distil-large-v3` | 6× faster than large-v3, near-parity  |

For ministers with Nigerian, Indian, or other non-US English accents,
`medium.en` on CPU or `distil-large-v3` on GPU makes a noticeable
difference — worth the extra install cost.

---

## Verifying the fix

After deploying v5, your log for the same 15 s of "John 3 verse seven
[…silence…]" should look like this:

```
21:07:48  INFO stt_engine - ok: John 3 verse seven (dur=1.80s conf=0.91 logprob=-0.25 cr=1.12)
[… then silence for 10+ seconds …]
21:07:58  DEBUG vad - endpoint: reason=silence dur=0.00s      <-- no utterance, nothing to transcribe
```

No more `['John 3.','John 4.','John 5.',…]` storm. If you ever see one
again, the client-side `HallucinationGuard` will reject the 3rd, 4th, … and
log `"guard rejected … rapidFire: 3 distinct chapters of John in 3s"`.

---

## Tuning knobs (if your room is unusual)

All in `python_server/stt_engine.py` — constants at the top:

| Constant           | Default | Raise if…                             | Lower if…                          |
|--------------------|--------:|---------------------------------------|------------------------------------|
| `MIN_RMS`          |  0.008  | Quiet room picks up too much         | Loud room eats quiet speakers      |
| `NO_SPEECH_MAX`    |  0.50   | Hallucinations still slip through    | Real speech gets dropped           |
| `LOGPROB_MIN`      | -1.00   | Too-tentative outputs slip through   | Correct low-confidence speech lost |
| `COMPRESSION_MAX`  |  2.40   | Repetitive real speech sneaks by     | Real speech with "Amen amen" lost  |
| `DEDUP_WINDOW_S`   |  6.0    | Same-ref repeats desired faster      | More aggressive dedup              |

And in `python_server/vad.py`:

| Constant               | Default | Raise if…                          | Lower if…                        |
|------------------------|--------:|------------------------------------|----------------------------------|
| `VAD_END_SILENCE_MS`   |   700   | Cuts references mid-sentence       | Feels sluggish                   |
| `VAD_SPEECH_MIN_MS`    |   160   | Stray HVAC sounds start utterances | Short words like "Luke" missed   |
| `PRE_SPEECH_PAD_MS`    |   250   | First syllable clipping            | Pre-roll noise leaks in          |

And in `WorshipHelperVSTO/Audio/Chunker.cs`:

| Property            | Default | Raise if…                        | Lower if…                   |
|---------------------|--------:|----------------------------------|-----------------------------|
| `MinRms`            |  0.010  | Quiet speakers cut off           | Still picking up HVAC       |
| `MinVoicedRatio`    |  0.15   | Short utterances get dropped     | Steady hum still escapes    |

---

## What to test after deploying

1. **Say one reference, then stop for 15 s.** → exactly one transcript.
2. **"Let's turn to Psalm twenty-three … verse four."** (long pause) → one
   final insertion of Psalm 23:4. The preliminary toast ("heard Psalm 23")
   should upgrade in place when "verse four" arrives.
3. **Play instrumental worship music at moderate volume near the mic.** →
   zero transcripts. The `/health` endpoint's `utterances_dropped` counter
   should increase.
4. **Project-fan hum + quiet room.** → zero transcripts.
5. **Nigerian/African accent test**: "Zechariah nine verse one through
   three." → Zechariah 9:1-3. The PhoneticCorrector already covers this.
6. **Fast speaker**: "Read with me John three sixteen." → John 3:16 within
   ~1 s of the period of silence.

---

## Autostart

Recommended: a Windows Scheduled Task that runs `worshiphelper-stt.ps1` at
login, hidden window. The C# client auto-reconnects, so you can close and
reopen PowerPoint freely without restarting the server.

```powershell
# Create the task (run once in an admin shell)
$action  = New-ScheduledTaskAction -Execute "powershell.exe" `
             -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$PWD\worshiphelper-stt.ps1`""
$trigger = New-ScheduledTaskTrigger -AtLogOn
$set     = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName "WorshipHelper STT" -Action $action `
    -Trigger $trigger -Settings $set -RunLevel Limited
```
