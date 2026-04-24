// ============================================================================
// Detection/HallucinationGuard.cs   —   v5.1  (hardened)
//
// Client-side defense-in-depth against Whisper hallucination patterns.
//
// The Python server already filters aggressively on its side (compression
// ratio, logprob, dedup, blacklist). This guard catches the subtler stuff
// that is only visible once you aggregate a few detections on the client:
//
//   1. RAPID-FIRE DIFFERENT CHAPTERS of the same book:
//        21:08:06  "John 3" "John 4" "John 5" "John 6" ...
//      Real ministers never race through six chapters in one second.
//
//   1b. RAPID-FIRE DIFFERENT VERSES of the same chapter (NEW in v5.1):
//        21:08:06  "Psalm 23:1" "Psalm 23:2" "Psalm 23:3" "Psalm 23:4" ...
//      The verse-loop version of (1). Same handling: reject.
//
//   2. LOW TRUST METRICS from the engine:
//        avg_logprob < -0.80 OR no_speech_prob > 0.35 OR compression > 2.00
//      (Server drops worse; we drop merely-suspicious for an insert.)
//
//   3. KNOWN WHISPER GHOST PHRASES (NEW in v5.1):
//        "thanks for watching", "please subscribe", "bye bye", etc.
//      Subsumes the old root-level HallucinationGuard's blacklist.
//
//   4. COMMON ENGLISH IDIOMS that happen to contain a book name:
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
using System.Text.RegularExpressions;

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

        /// <summary>How many distinct verses of the same chapter within the window
        /// counts as the verse-loop hallucination (v5.1).</summary>
        public int RapidFireDistinctVersesThreshold { get; set; } = 4;

        // ------ Whisper ghost-phrase blacklist (v5.1, ported from old guard) ------
        //
        // These are phrases Whisper hallucinates when it has trained on YouTube /
        // subtitled podcast data and the audio stream is actually silent or music.
        // If the WHOLE transcript matches one of these, drop it — even if a
        // book name happened to slip through the phonetic corrector.
        private static readonly HashSet<string> _ghostPhrases =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "thanks for watching", "thank you for watching",
            "thanks for watching!", "thank you.", "thank you",
            "please subscribe", "subscribe to my channel",
            "like and subscribe", "please like and subscribe",
            "don't forget to subscribe", "hit that subscribe button",
            "see you next time", "see you in the next video",
            "bye", "bye.", "bye bye", "goodbye",
            "you", "okay", "ok", "yeah",
            "hmm", "hmm.", "um", "uh", "oh",
            ".", "...", ". . .", "!", "?",
            "music", "[music]", "(music)", "♪", "♪♪",
            "silence", "[silence]", "[silent]",
            "transcribed by", "transcription by",
            "applause", "[applause]", "(applause)",
            "laughter", "[laughter]", "(laughter)",
            "amen", "god bless", "god bless you",
            "hallelujah", "praise the lord",
        };

        // ------ State ----------------------------------------------------
        private sealed class HistoryEntry
        {
            public DateTime Ts;
            public string Book;
            public int Chapter;
            public int Verse;
        }
        private readonly LinkedList<HistoryEntry> _history = new LinkedList<HistoryEntry>();
        private readonly object _lock = new object();

        // "Scripture-intent" words. If the raw text contains any of these,
        // we're more willing to accept a borderline detection.
        private static readonly HashSet<string> _intentWords =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "turn", "turning", "read", "reading", "open", "go", "jump", "flip",
            "chapter", "verse", "verses", "scripture", "scriptures",
            "passage", "book", "says", "through", "from", "let's", "lets",
            "hear", "hearing", "look", "word",
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
            input = input ?? new GuardInput();

            // 0. Whisper-ghost blacklist — apply even when no book detected,
            //    because a ghost phrase like "Thanks for watching!" can still
            //    leak downstream if we're not vigilant.
            if (!string.IsNullOrWhiteSpace(input.RawText))
            {
                string normalised = NormaliseForBlacklist(input.RawText);
                if (_ghostPhrases.Contains(normalised))
                    return Rejected_($"ghostPhrase=\"{normalised}\"");
                // Short echoes of ghost phrases e.g. "thank" alone:
                if (normalised.Length <= 3 && !ContainsDigit(normalised))
                    return Rejected_("tooShort");
            }

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

            // 2. Rapid-fire hallucination patterns
            PruneHistory();
            int distinctChapters;
            int distinctVerses;
            lock (_lock)
            {
                var sameBook = _history
                    .Where(h => h.Book != null
                             && h.Book.Equals(input.Book, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                distinctChapters = sameBook.Select(h => h.Chapter).Distinct().Count();
                distinctVerses   = sameBook
                    .Where(h => h.Chapter == input.Chapter && h.Verse > 0)
                    .Select(h => h.Verse)
                    .Distinct()
                    .Count();
            }

            //   2a. If we've already seen N distinct chapters of the same book in
            //       the last K seconds, this looks exactly like the "John 3/4/5/6"
            //       loop. Reject.
            if (distinctChapters >= RapidFireDistinctChaptersThreshold)
                return Rejected_($"rapidFireChapters: {distinctChapters} chapters of {input.Book} " +
                                 $"in {RapidFireWindow.TotalSeconds:F0}s");

            //   2b. Same-chapter verse loop ("Psalm 23:1 / 23:2 / 23:3 / 23:4"):
            //       also a hallucination pattern that Whisper produces on
            //       slightly-different-audio loops. Only reject when the
            //       incoming item would be the 4th+ distinct verse.
            if (input.Verse > 0 && distinctVerses >= RapidFireDistinctVersesThreshold)
                return Rejected_($"rapidFireVerses: {distinctVerses} verses of {input.Book} " +
                                 $"{input.Chapter} in {RapidFireWindow.TotalSeconds:F0}s");

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
                    Ts      = DateTime.UtcNow,
                    Book    = input.Book,
                    Chapter = input.Chapter,
                    Verse   = input.Verse,
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
            foreach (var tok in rawText.ToLowerInvariant().Split(new[] { ' ', ',', '.', '?', '!' },
                                                                 StringSplitOptions.RemoveEmptyEntries))
                if (_intentWords.Contains(tok)) return true;
            return false;
        }

        private static string NormaliseForBlacklist(string text)
        {
            string t = text.ToLowerInvariant().Trim();
            t = Regex.Replace(t, @"[^a-z0-9\s]", " ");
            t = Regex.Replace(t, @"\s+", " ").Trim();
            return t;
        }

        private static bool ContainsDigit(string s)
        {
            for (int i = 0; i < s.Length; i++) if (char.IsDigit(s[i])) return true;
            return false;
        }
    }
}
