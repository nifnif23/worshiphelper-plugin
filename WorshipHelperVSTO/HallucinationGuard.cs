// ============================================================================
// HallucinationGuard.cs
//
// Client-side hallucination filter applied to transcripts received from the
// Python sidecar (PythonClient.TranscriptReceived) before they reach
// BibleReferenceDetector.
//
// The Python sidecar (stt_engine.py) already applies these guards server-side,
// but the C# layer adds a second pass because:
//   1. Network jitter can cause a stale transcript to arrive after a session
//      reset, bypassing the server's dedup state.
//   2. The server's dedup window is per-engine; the guard here is per-UI-session
//      so multiple active sessions don't cross-contaminate.
//   3. Confidence thresholds may differ between the server's calibration and
//      the end-user's microphone environment.
//
// Usage:
//   private readonly HallucinationGuard _guard = new HallucinationGuard();
//
//   // in TranscriptReceived handler:
//   if (_guard.IsHallucination(e.Text, e.Confidence)) return;
//   // ... proceed to BibleReferenceDetector
//
// Thread-safe: all mutable state is guarded by _lock.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using log4net;

namespace WorshipHelperVSTO
{
    /// <summary>
    /// Filters transcripts that are almost certainly Whisper hallucinations
    /// rather than real speech, before they reach the reference detector.
    /// </summary>
    public sealed class HallucinationGuard
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(HallucinationGuard));

        // ---------------------------------------------------------------
        // Configuration
        // ---------------------------------------------------------------

        /// <summary>
        /// Transcripts whose confidence is below this threshold are dropped.
        /// Matches the Python sidecar's NO_SPEECH_MAX → confidence mapping.
        /// </summary>
        public float MinConfidence { get; set; } = 0.15f;

        /// <summary>
        /// If the same normalised text arrives again within this window, it
        /// is suppressed as a duplicate.  Mirrors Python's DEDUP_WINDOW_S.
        /// </summary>
        public TimeSpan DedupWindow { get; set; } = TimeSpan.FromSeconds(6);

        // ---------------------------------------------------------------
        // Common Whisper hallucination phrases (mirrors stt_engine.py)
        // ---------------------------------------------------------------
        private static readonly HashSet<string> HallucinationPhrases = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "thanks for watching", "thank you for watching",
            "thanks for watching!", "thank you.", "thank you",
            "please subscribe", "subscribe to my channel",
            "you", "bye", "bye.", "bye bye",
            ".", "...", ". . .",
            "music", "[music]", "(music)",
            "silence", "[silence]",
            "transcribed by", "transcription by",
            "please like and subscribe",
            "see you next time", "see you in the next video",
            "don't forget to subscribe",
            "hmm", "hmm.", "um", "uh",
            "yeah", "okay", "ok",
            "amen",
            "god bless", "god bless you",
            "hallelujah", "praise the lord",
        };

        // ---------------------------------------------------------------
        // Dedup state
        // ---------------------------------------------------------------
        private readonly object _lock = new object();
        private string _lastNormalisedText = string.Empty;
        private DateTime _lastEmitTime = DateTime.MinValue;

        // ---------------------------------------------------------------
        // Public API
        // ---------------------------------------------------------------

        /// <summary>
        /// Returns <c>true</c> if the transcript should be suppressed.
        /// Logs the reason at DEBUG level so production logs stay clean.
        /// </summary>
        /// <param name="text">Raw transcript text from the sidecar.</param>
        /// <param name="confidence">0..1 confidence from TranscriptEventArgs.</param>
        public bool IsHallucination(string text, float confidence)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                Log.Debug("HallucinationGuard: drop — empty text");
                return true;
            }

            // Gate 1: confidence floor
            if (confidence < MinConfidence)
            {
                Log.DebugFormat("HallucinationGuard: drop — confidence {0:F2} < {1:F2} for: {2}",
                    confidence, MinConfidence, text);
                return true;
            }

            string normalised = Normalise(text);

            // Gate 2: too short after normalisation (single char / punctuation only)
            if (normalised.Length <= 2)
            {
                Log.DebugFormat("HallucinationGuard: drop — too short after normalise: {0}", text);
                return true;
            }

            // Gate 3: known-bad phrase blacklist
            if (HallucinationPhrases.Contains(normalised) || HallucinationPhrases.Contains(text.Trim()))
            {
                Log.DebugFormat("HallucinationGuard: drop — blacklist match: {0}", text);
                return true;
            }

            // Gate 4: near-duplicate suppression
            lock (_lock)
            {
                if (IsNearDuplicate(normalised))
                {
                    Log.DebugFormat("HallucinationGuard: drop — dedup within {0}s: {1}",
                        DedupWindow.TotalSeconds, text);
                    return true;
                }

                _lastNormalisedText = normalised;
                _lastEmitTime = DateTime.UtcNow;
            }

            return false;
        }

        /// <summary>
        /// Clears dedup state.  Call on session disconnect or explicit reset.
        /// </summary>
        public void Reset()
        {
            lock (_lock)
            {
                _lastNormalisedText = string.Empty;
                _lastEmitTime = DateTime.MinValue;
            }
        }

        // ---------------------------------------------------------------
        // Helpers
        // ---------------------------------------------------------------

        /// <summary>
        /// Aggressive normalisation: lowercase, strip punctuation, collapse whitespace.
        /// Mirrors stt_engine.py's <c>_normalise()</c>.
        /// </summary>
        private static string Normalise(string text)
        {
            string t = text.ToLowerInvariant().Trim();
            t = Regex.Replace(t, @"[^a-z0-9\s:]", " ");
            t = Regex.Replace(t, @"\s+", " ").Trim();
            return t;
        }

        /// <summary>
        /// Returns true if <paramref name="normalised"/> matches the last emitted
        /// text exactly, or differs only in a trailing integer (the classic
        /// "John 3 / John 4 / John 5" hallucination-loop signature).
        /// Must be called inside _lock.
        /// </summary>
        private bool IsNearDuplicate(string normalised)
        {
            if (string.IsNullOrEmpty(_lastNormalisedText))
                return false;

            if ((DateTime.UtcNow - _lastEmitTime) > DedupWindow)
                return false;

            if (string.Equals(normalised, _lastNormalisedText, StringComparison.Ordinal))
                return true;

            // "john 3" vs "john 4" -- same prefix, trailing counter incremented
            string[] tokA = normalised.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            string[] tokB = _lastNormalisedText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (tokA.Length == 2 && tokB.Length == 2
                && string.Equals(tokA[0], tokB[0], StringComparison.Ordinal)
                && int.TryParse(tokA[1], out _)
                && int.TryParse(tokB[1], out _))
            {
                return true;
            }

            return false;
        }
    }
}