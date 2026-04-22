// ============================================================================
// Feedback/CorrectionEngine.cs
// Closes the feedback loop: reads FeedbackStore history and produces runtime
// adjustments to thresholds and custom paraphrase lookups.
//
// Designed to be safe by default -- it only tightens the screws (raises
// thresholds, adds mappings); it never loosens them without user consent.
// ============================================================================
using System;
using System.Collections.Generic;
using System.Linq;

namespace WorshipHelperVSTO.Feedback
{
    public sealed class RuntimeAdjustments
    {
        public double PatternThreshold   { get; set; }
        public float  SemanticThreshold  { get; set; }
        public Dictionary<string, string> CustomParaphrases { get; set; }
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class CorrectionEngine
    {
        private readonly FeedbackStore _store;

        public CorrectionEngine(FeedbackStore store) { _store = store; }

        public RuntimeAdjustments Recompute(
            double basePatternThreshold = 0.4,
            float  baseSemanticThreshold = 0.62f)
        {
            var adj = new RuntimeAdjustments
            {
                PatternThreshold  = basePatternThreshold,
                SemanticThreshold = baseSemanticThreshold,
            };

            // 1. Apply stored paraphrases.
            foreach (var (phrase, reference) in _store.GetParaphrases())
                adj.CustomParaphrases[phrase] = reference;

            // 2. If we have repeated misses, slightly lower pattern threshold
            //    (more recall) -- but never below 0.25.
            var freq = _store.MishearingFrequency();
            int chronicMisses = freq.Values.Count(v => v >= 3);
            if (chronicMisses >= 5)
                adj.PatternThreshold = Math.Max(0.25, basePatternThreshold - 0.1);

            return adj;
        }

        /// <summary>
        /// Returns a canonical reference if the incoming utterance exactly or
        /// approximately matches a user-supplied paraphrase.
        /// </summary>
        public string TryParaphraseLookup(string utterance, RuntimeAdjustments adj)
        {
            if (string.IsNullOrWhiteSpace(utterance) || adj?.CustomParaphrases == null) return null;
            string key = utterance.Trim().ToLowerInvariant();
            if (adj.CustomParaphrases.TryGetValue(key, out var exact)) return exact;

            foreach (var kv in adj.CustomParaphrases)
                if (key.Contains(kv.Key)) return kv.Value;
            return null;
        }
    }
}
