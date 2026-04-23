# SpeechToScriptureService.cs — v5 patch

The v4 file is already well-structured. We apply FOUR surgical patches:

1. Add a `HallucinationGuard` field + instantiate it in the constructor.
2. Pipe the engine's new metrics (avg_logprob etc.) through the guard before
   any detection logic runs.
3. Expose a public `Guard` property so the ribbon/debug panel can read its
   metrics (accepted / deferred / rejected counts).
4. Tighten `MinCombinedConfidence` default from `0.4` to `0.55`, and raise
   `MinSpeechConfidence`-default from `0.10` to `0.30`. These were too lenient
   for hallucination-hardened operation.

Drop the snippets below into `SpeechToScriptureService.cs` at the marked
locations. If you prefer, use `SpeechToScriptureService.v5.cs` in this
package (fully rewritten, same public API).

---

## Patch 1 — new field + property (near the top of the class, with other fields)

```csharp
using WorshipHelperVSTO.Detection;   // NEW

// inside class SpeechToScriptureService : IDisposable
private readonly HallucinationGuard _guard = new HallucinationGuard();

/// <summary>
/// Exposes the hallucination guard so the debug panel can read its counters.
/// </summary>
public HallucinationGuard Guard => _guard;
```

## Patch 2 — change default thresholds

```csharp
// replace:
public double MinCombinedConfidence { get; set; } = 0.4;
// with:
public double MinCombinedConfidence { get; set; } = 0.55;

// and in the constructor:
_listener.MinEngineConfidence = 0.30f;   // was 0.10
```

## Patch 3 — run the guard in OnSpeechRecognised

Insert immediately AFTER `var detected = BibleReferenceDetector.DetectBest(e.Text);`
and BEFORE the chapter-only upgrade block:

```csharp
// v5: hallucination guard. Rejects suspicious repeats / low-trust segments
// before they can touch the auto-insert machinery.
var guardInput = new GuardInput
{
    RawText          = e.Text,
    Book             = detected?.BookName,
    Chapter          = ParseChapterNumber(detected?.ReferenceFragment ?? ""),
    Verse            = ParseVerseNumber(detected?.ReferenceFragment ?? ""),
    EngineConfidence = e.Confidence,
    AvgLogProb       = e.AvgLogProb,
    CompressionRatio = e.CompressionRatio,
    NoSpeechProb     = e.NoSpeechProb,
    DurationSeconds  = e.DurationSeconds,
};

var decision = _guard.Check(guardInput);
if (decision.Verdict == GuardVerdict.Reject)
{
    log.Info($"Pipeline: guard rejected \"{e.Text}\" -> {decision.Reason}");
    return;
}
if (decision.Verdict == GuardVerdict.Defer)
{
    // Force through the chapter-only debounce path (give us confirmation).
    log.Debug($"Pipeline: guard deferred \"{e.Text}\" -> {decision.Reason}");
    // Fall through; the existing chapter-only grace window will handle it.
}
```

## Patch 4 — helper method (anywhere in the class)

```csharp
/// <summary>Extracts the verse portion of a reference fragment like "3:16-18" -> 16. Returns 0 if chapter-only.</summary>
private static int ParseVerseNumber(string refFragment)
{
    if (string.IsNullOrEmpty(refFragment)) return 0;
    int colon = refFragment.IndexOf(':');
    if (colon < 0) return 0;
    string versePart = refFragment.Substring(colon + 1);
    int dash = versePart.IndexOfAny(new[] { '-', '–' });
    if (dash >= 0) versePart = versePart.Substring(0, dash);
    return int.TryParse(versePart.Trim(), out int v) ? v : 0;
}
```

---

## Why these four are enough

* Patch 3 is the functional fix — it plugs the rapid-fire hallucination
  loop the user saw ("John 3 / 4 / 5 / 6 / 7 / 8 / 9").
* Patches 1, 2, 4 are the scaffolding.
* The chapter-only grace window you already have (v4) PLUS the defer path
  handles the "Psalm twenty-three … verse four" two-utterance case.
* All existing behaviour for chapter:verse references is preserved.
