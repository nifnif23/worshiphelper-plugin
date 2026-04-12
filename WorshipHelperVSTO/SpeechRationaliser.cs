// ============================================================================
// SpeechRationaliser.cs
// Pre-processes raw Vosk output into the cleanest possible text for
// BibleReferenceDetector to parse.
//
// With a constrained grammar, Vosk output should already be close to correct,
// but this layer handles:
//   1. [unk] tokens  — strip them (non-Bible speech produces these)
//   2. Spoken punctuation — "colon" → recognised by SpokenNumberConverter
//                           "dash" / "hyphen" → recognised as range separator
//   3. Digit strings — already handled by WordsToNumber, no change needed
//
// Example transformations:
//   "zechariah nine colon eight dash ten" → "zechariah nine colon eight dash ten"
//   "[unk] zechariah [unk] nine eight to ten [unk]" → "zechariah nine eight to ten"
//   "john three [unk] sixteen" → "john three sixteen"
// ============================================================================

using System.Text.RegularExpressions;

namespace WorshipHelperVSTO
{
    public static class SpeechRationaliser
    {
        /// <summary>
        /// Cleans raw Vosk output before it enters the Bible reference detection
        /// pipeline.  Safe to call with any string; always returns a string
        /// (never null).
        /// </summary>
        public static string Rationalise(string rawVoskText)
        {
            if (string.IsNullOrWhiteSpace(rawVoskText))
                return string.Empty;

            string text = rawVoskText;

            // 1. Strip [unk] tokens produced by the grammar for out-of-vocabulary
            //    phonemes.  Without stripping, they can break book-name matching.
            text = Regex.Replace(text, @"\[unk\]", " ", RegexOptions.IgnoreCase);

            // 2. Collapse multiple spaces left behind by stripping
            text = Regex.Replace(text, @"\s{2,}", " ").Trim();

            return text;
        }
    }
}
