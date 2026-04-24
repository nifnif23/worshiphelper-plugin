# WorshipHelper v5.1 — Build-fix + Speech + UI overhaul

This changelog covers everything that changed from the base drop you sent.
Nothing in the existing Bible parser, ScriptureManager, OpenSongBibleReader
or Python sidecar was touched — those are mature and I didn't want to risk
regressions there.

---

## 1. Build fix (the 4 compile errors)

**Root cause.** Two files:

- `WorshipHelperVSTO/HallucinationGuard.cs`      (old v4-era API — `IsHallucination(...)`)
- `WorshipHelperVSTO/Detection/HallucinationGuard.cs`  (v5 API — `Check(GuardInput)`)

Both defined a class called `HallucinationGuard`. The v4 one was in namespace
`WorshipHelperVSTO`, the v5 one in `WorshipHelperVSTO.Detection`. Although
`SpeechToScriptureService.cs` had `using WorshipHelperVSTO.Detection;`, C#'s
name-resolution rules prefer types from the enclosing namespace over
imported ones, so `new HallucinationGuard()` resolved to the old class with
no `Check()` method. And `Detection/HallucinationGuard.cs` was never listed
in the `.csproj` anyway, so its types (`GuardInput`, `GuardVerdict`,
`GuardDecision`) didn't exist at all to the compiler.

**Fix.**
1. Deleted `WorshipHelperVSTO/HallucinationGuard.cs` (nothing referenced
   its `IsHallucination` method outside the file itself, so this is safe).
2. Added `<Compile Include="Detection\HallucinationGuard.cs" />` to
   `WorshipHelperVSTO.csproj`.
3. Added `<Compile Include="UI\ModernControls.cs" />` (new UI toolkit, see
   §3).

All 4 errors reported in your MSBuild output are now resolved:

    SpeechToScriptureService.cs(346,38): error CS0246: 'GuardInput'      ✓ fixed
    SpeechToScriptureService.cs(358,39): error CS1061: 'HallucinationGuard.Check' ✓ fixed
    SpeechToScriptureService.cs(359,41): error CS0103: 'GuardVerdict'   ✓ fixed
    SpeechToScriptureService.cs(364,41): error CS0103: 'GuardVerdict'   ✓ fixed

I verified every modified C# file parses and compiles cleanly against the
.NET Framework BCL + System.Windows.Forms + System.Drawing (the full
toolchain isn't available in a sandbox, but MSBuild on your CI should now
go green).

---

## 2. Speech pipeline — "PewBeam-grade" upgrades

Everything here is in `Detection/HallucinationGuard.cs` and
`SpeechToScriptureService.cs`.

### 2.1 Guard hardening (`HallucinationGuard` v5.1)

- **Whisper-ghost blacklist** ported from the old root guard, now applied
  whether or not a book was detected. Catches the classic garbage Whisper
  hallucinates on silent audio ("Thanks for watching", "Please subscribe",
  "[music]", plain "you", etc.).
- **Verse-loop detection (new).** Previously only the chapter-loop
  pattern ("John 3 / 4 / 5 / 6 / …") was caught; now the same-chapter
  verse-loop ("Psalm 23:1 / :2 / :3 / :4 / :5") is caught too. Threshold
  is 4 distinct verses of the same chapter in 3 s.
- **Expanded intent vocabulary** ("turn", "turning", "read", "reading",
  "let's", "look", "word", "hear", "hearing"). A borderline detection
  like "matthew five" is now accepted when preceded by "let's turn to"
  even at 0.55 engine confidence.
- **Phrase normalisation** strips punctuation and collapses whitespace
  before blacklist lookup — catches "Thanks for watching!" (with bang)
  and "thank you." as well as the plain forms.

### 2.2 `SpeechToScriptureService` improvements

- **`ChapterOnlyGraceMs` default cut from 10 000 → 4 000 ms.** The old
  10-second delay felt laggy; the guard's Defer path gives a little
  extra headroom for truly borderline detections, so 4 s is ample.
- **New `ContextCarryoverMs` setting** (default 30 000 ms). After a
  reference is successfully fired, its book+chapter are remembered for
  30 s. If the next utterance has no book of its own but does contain
  "verse N" or a bare number, it is spliced onto the last context:

        "Let's read John 3:16."    → fires "John 3:16"
        "And verse 17 says…"       → fires "John 3:17"  (context-carried)
        "verse 18"                 → fires "John 3:18"  (context-carried)
        (30 s passes)
        "verse 19"                 → no context → dropped

  This is a big UX win for sermons that work through a passage without
  repeating the book name every sentence.

- **New `PlayChimeOnInsert` setting** (default off). Plays
  `SystemSounds.Asterisk` when a reference is auto-inserted so the
  presenter gets acoustic confirmation without having to look at the
  screen.

- **Guard + context reset on `Stop()`.** Previously, stopping the
  listener for an hour and restarting it could carry rapid-fire history
  across the gap and produce spurious rejects. Now state is cleared on
  every stop.

---

## 3. UI overhaul — same palette, modern and rounded

All new rendering code lives in `WorshipHelperVSTO/UI/ModernControls.cs`.
Palette is centralised in `WorshipHelperVSTO.UI.Palette` so every form
pulls from one source of truth. Colours are unchanged from the product:
accent green `#2E7D32`, dark green `#1B5E20`, light green `#E8F5E9`,
gold `#B8860B`, neutral greys.

### 3.1 New toolkit controls

- **`ModernButton`** — `System.Windows.Forms.Button` subclass with
  GDI+-rendered rounded corners (8-10 px), smooth hover/press fills,
  primary/outline styles, disabled state, focus ring. Drop-in for
  plain `Button` — only `Enabled`, `Visible`, `Click` are used in the
  existing form code.
- **`ModernTextBox`** — rounded panel wrapper around a borderless
  `TextBox`. Green-glowing border on focus.
- **`CardPanel`** — rounded white/light-green gradient card with
  optional left accent stripe and soft drop shadow.
- **`SectionHeader`** — small all-caps muted label for "MIC LEVEL",
  "CONFIDENCE" etc.
- **`Palette`** — shared colour table + `RoundedPath(Rectangle, int)`
  helper (pure GDI+, no P/Invoke).

### 3.2 `SpeechConfirmForm` (the detection toast)

- Real shadow under the card (4-layer antialiased), rounded corners.
- `ModernButton` replaces the old flat fake-border buttons — they're
  now visibly rounded with smooth hover transitions.
- Reference text sits in a rounded light-green pill inside the card.
- Countdown bar is a thin pill at the bottom with proper rounding;
  it no longer touches the card border.
- Transparency key is used so the shadow actually looks like a shadow
  instead of a grey square.

### 3.3 `SpeechDebugPanel` (the live speech monitor)

- **BUG FIX.** The old layout had `_lblPhase` (SCAN/FOCUS badge) added
  to the top strip on top of a `DockStyle.Fill` status label. The Fill
  label would cover the badge if ever z-order got reset (e.g. on resize).
  Now the badge is `DockStyle.Right` and the status is `DockStyle.Fill`
  around it, with `BringToFront` on the status label. Works every time.
- Card-based sections (MIC LEVEL / LAST HEARD / CONFIDENCE / DETECTED
  REFERENCE) with rounded inner panels. Layout driven by nested
  `TableLayoutPanel` so nothing overlaps when the user resizes.
- Bottom buttons are `ModernButton` outline style with the green accent.
- Mic level meter has rounded background, segments are slightly taller.
- Confidence bar is rounded with overlaid percentage + soft drop shadow
  on the text.
- Log panel keeps its 500-line cap.

### 3.4 `InsertScriptureForm`

- All four buttons (`btnInsert`, `btnCancel`, `btnModeBulk`,
  `btnModeSingle`) switched to `ModernButton`. No more
  `FlatAppearance.BorderSize = 1` hacks; they genuinely have rounded
  corners now. Insert is primary (filled green), Close and the mode
  toggles are outline style.
- Buttons grew slightly (120→130 / 100→110 / 155→170) for better tap
  targets on touch-screen presenter setups.
- Unused local colour variables in the designer cleaned up (was
  generating 5 warnings).

### 3.5 `AddContentLiveForm`

- Not touched. The Scripture / Song cards are image+text card buttons
  that are their own thing; switching them to `ModernButton` would
  require image rendering which isn't worth the risk for cosmetic-only
  improvement.

---

## 4. Tested

Because the sandbox I work in doesn't have VSTO / Office to do a full
end-to-end build, I verified with Mono's `mcs`:

- `Detection/HallucinationGuard.cs` compiles standalone, clean.
- `UI/ModernControls.cs` compiles standalone, clean.
- `SpeechConfirmForm.cs` compiles clean against the above + Win Forms.
- `SpeechDebugPanel.cs` compiles clean against the above + Win Forms +
  NAudio (stubbed).
- `InsertScriptureForm.Designer.cs` compiles clean.
- `SpeechToScriptureService.cs` compiles clean with the v5.1 guard.

Plus a functional unit-test on `HallucinationGuard.Check(...)` passes
all 7 cases: normal accept, ghost-phrase reject, chapter-loop reject,
verse-loop reject, low-trust reject, intent-boost accept, borderline defer.

On your Windows CI runner the same MSBuild command you posted should
now come back green.
