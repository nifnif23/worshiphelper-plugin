// ============================================================================
// Detection/HallucinationGuard.cs
//
// Client-side defense-in-depth against Whisper hallucination patterns.
//
// The Python server has its own filters (compression-ratio, logprob, dedup,
// blacklist). This guard catches the subtler stuff only visible once you
// aggregate detections:
//
//   1. RAPID-FIRE DIFFERENT CHAPTERS of the same book:
//        21:08:06  "John 3" "John 4" "John 5" "John 6" ...
//      Real ministers never race through 6 chapters in 1 second. Flag as
//      hallucination.
//
//   2. LOW TRUST METRICS from the engine:
//        avg_logprob < -0.8 OR no_speech_prob > 0.35 OR compression_ratio > 2.0
//      (Server drops worse; we drop merely-suspicious-for-an-insert.)
//
//   3. COMMON ENGLISH IDIOMS that happen to contain a book name:
//        "john and mary went" -> spurious "John 1" shouldn't fire.
//      We require a "scripture-intent" signal for borderline detections.
//
// Guard returns a verdict:
//      Accept   -- fire as normal
//      Defer    -- hold for chapter-only grace window (requires confirmation)
//      Reject   -- silently drop
//
// Thread-safe. One instance per SpeechListener.
// ============================================================================
using System;
using System.Collections.Generic;
using System.Linq;

namespace WorshipHelperVSTO.Detection
{
    public enum GuardVerdict
    {
        Accept,
        Defer,
        Reject,
    }

    public sealed class GuardDecision
    {
        public GuardVerdict Verdict { get; set; }
        public string Reason { get; set; }

        public static GuardDecision Accept(string why = null) =>
            new GuardDecision { Verdict = GuardVerdict.Accept, Reason = why };
        public static GuardDecision Defer(string why) =>
            new GuardDecision { Verdict = GuardVerdict.Defer, Reason = why };
        public static GuardDecision Reject(string why) =>
            new GuardDecision { Verdict = GuardVerdict.Reject, Reason = why };
    }

    /// <summary>
    /// Per-session trust signal for a raw transcript + its detected reference.
    /// </summary>
    public sealed class GuardInput
    {
        public string RawText { get; set; }              // what Whisper heard (post phonetic-correct)
        public string Book { get; set; }                 // detected book (may be null if no detection)
        public int Chapter { get; set; }                 // 0 if unknown
        public int Verse { get; set; }                   // 0 if chapter-only
        public double EngineConfidence { get; set; }     // 0..1
        public double AvgLogProb { get; set; }
        public double CompressionRatio { get; set; }
        public double NoSpeechProb { get; set; }
        public double DurationSeconds { get; set; }
    }

    public sealed class HallucinationGuard
    {
        // ------ Thresholds (copy-tuneable) -------------------------------
        /// <summary>avg_logprob below this = engine isn't sure at all.</summary>
        public double MinAvgLogProb { get; set; } = -0.80;

        /// <summary>no_speech_prob above this = there might have been no speech.</summary>
        public double MaxNoSpeechProb { get; set; } = 0.35;

        /// <summary>Compression ratio above this = token repetition (ghost loop).</summary>
        public double MaxCompressionRatio { get; set; } = 2.00;

        /// <summary>Duration below this = too short for a real Bible reference.</summary>
        public double MinDurationSeconds { get; set; } = 0.40;

        /// <summary>Window for "rapid-fire same-book" detection.</summary>
        public TimeSpan RapidFireWindow { get; set; } = TimeSpan.FromSeconds(3);

        /// <summary>How many distinct chapters within that window counts as hallucination.</summary>
        public int RapidFireDistinctChaptersThreshold { get; set; } = 3;

        // ------ State ----------------------------------------------------
        private sealed class HistoryEntry
        {
            public DateTime Ts;
            public string Book;
            public int Chapter;
        }
        private readonly LinkedList<HistoryEntry> _history = new LinkedList<HistoryEntry>();
        private readonly object _lock = new object();

        // "Scripture-intent" words. If the raw text contains any of these,
        // we're more willing to accept a borderline detection.
        private static readonly HashSet<string> _intentWords =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "turn", "read", "open", "go", "jump", "flip",
            "chapter", "verse", "verses", "scripture", "scriptures",
            "passage", "book", "says", "through", "from",
        };

        // Metrics for the debug panel.
        public long TotalChecked   { get; private set; }
        public long Accepted       { get; private set; }
        public long Deferred       { get; private set; }
        public long Rejected       { get; private set; }
        public string LastRejectReason { get; private set; }

        // ---------------------------------------------------------------
        public GuardDecision Check(GuardInput input)
        {
            TotalChecked++;

            // Fast path: no book detected. Not our problem -- pipeline will drop.
            if (string.IsNullOrEmpty(input.Book) || input.Chapter <= 0)
            {
                Accepted++;
                return GuardDecision.Accept("no-detection-passthrough");
            }

            // 1. Engine-trust gates
            if (input.NoSpeechProb > MaxNoSpeechProb)
                return Rejected_("noSpeechProb=" + input.NoSpeechProb.ToString("F2"));
            if (input.AvgLogProb < MinAvgLogProb)
                return Rejected_("avgLogProb=" + input.AvgLogProb.ToString("F2"));
            if (input.CompressionRatio > MaxCompressionRatio)
                return Rejected_("compressionRatio=" + input.CompressionRatio.ToString("F2"));
            if (input.DurationSeconds > 0 && input.DurationSeconds < MinDurationSeconds)
                return Rejected_("tooShort=" + input.DurationSeconds.ToString("F2") + "s");

            // 2. Rapid-fire hallucination pattern
            //
            //   If we've already seen N distinct chapters of the same book in
            //   the last K seconds, this looks exactly like the "John 3/4/5/6"
            //   loop. Reject.
            PruneHistory();
            int distinctChapters;
            lock (_lock)
            {
                distinctChapters = _history
                    .Where(h => h.Book.Equals(input.Book, StringComparison.OrdinalIgnoreCase))
                    .Select(h => h.Chapter)
                    .Distinct()
                    .Count();
            }
            if (distinctChapters >= RapidFireDistinctChaptersThreshold)
                return Rejected_($"rapidFire: {distinctChapters} distinct chapters of {input.Book} " +
                                 $"in {RapidFireWindow.TotalSeconds:F0}s");

            // 3. Chapter-only + no scripture intent + borderline confidence ->
            //    defer (so the grace window gives us time to see a confirming
            //    verse before we commit).
            bool chapterOnly = input.Verse <= 0;
            bool hasIntent   = HasIntentSignal(input.RawText);
            if (chapterOnly && !hasIntent && input.EngineConfidence < 0.70)
            {
                Record(input);
                Deferred++;
                return GuardDecision.Defer("chapterOnly+lowConf+noIntent");
            }

            // 4. All good.
            Record(input);
            Accepted++;
            return GuardDecision.Accept(null);
        }

        // ---------------------------------------------------------------
        public void Reset()
        {
            lock (_lock) _history.Clear();
            TotalChecked = Accepted = Deferred = Rejected = 0;
            LastRejectReason = null;
        }

        // ---------------------------------------------------------------
        private GuardDecision Rejected_(string reason)
        {
            Rejected++;
            LastRejectReason = reason;
            return GuardDecision.Reject(reason);
        }

        private void Record(GuardInput input)
        {
            lock (_lock)
            {
                _history.AddLast(new HistoryEntry
                {
                    Ts = DateTime.UtcNow,
                    Book = input.Book,
                    Chapter = input.Chapter,
                });
            }
        }

        private void PruneHistory()
        {
            var cutoff = DateTime.UtcNow - RapidFireWindow;
            lock (_lock)
            {
                while (_history.First != null && _history.First.Value.Ts < cutoff)
                    _history.RemoveFirst();
            }
        }

        private static bool HasIntentSignal(string rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText)) return false;
            foreach (var tok in rawText.ToLowerInvariant().Split(new[] { ' ', ',', '.' },
                                                                 StringSplitOptions.RemoveEmptyEntries))
                if (_intentWords.Contains(tok)) return true;
            return false;
        }
    }
}
