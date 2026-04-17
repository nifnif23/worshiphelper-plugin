// ============================================================================
// SpokenNumberConverter.cs
// Converts spoken English number words into digit strings for Bible references.
//
// Handles: 0–199, cardinals, ordinals (for book prefixes), and colloquial
// forms like "one oh five" → 105.
//
// Drop into:  WorshipHelperVSTO/SpokenNumberConverter.cs
// Namespace:  WorshipHelperVSTO
// ============================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace WorshipHelperVSTO
{
    /// <summary>
    /// Reusable component that converts spoken English numbers into digit strings.
    /// Designed specifically for the range needed in Bible references (0–199).
    /// </summary>
    public static class SpokenNumberConverter
    {
        // -----------------------------------------------------------------------
        // Lookup tables
        // -----------------------------------------------------------------------

        private static readonly Dictionary<string, int> Cardinals = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            {"zero",0},{"oh",0},{"o",0},
            {"one",1},{"two",2},{"three",3},{"four",4},{"five",5},
            {"six",6},{"seven",7},{"eight",8},{"nine",9},{"ten",10},
            {"eleven",11},{"twelve",12},{"thirteen",13},{"fourteen",14},
            {"fifteen",15},{"sixteen",16},{"seventeen",17},{"eighteen",18},
            {"nineteen",19},{"twenty",20},{"thirty",30},{"forty",40},
            {"fifty",50},{"sixty",60},{"seventy",70},{"eighty",80},{"ninety",90},
            {"hundred",100},
        };

        /// <summary>
        /// Ordinal → digit value.  Used primarily for book-name prefixes
        /// ("first" → 1, "second" → 2, "third" → 3).
        /// </summary>
        public static readonly Dictionary<string, int> Ordinals = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            {"first",1},{"1st",1},
            {"second",2},{"2nd",2},
            {"third",3},{"3rd",3},
            {"fourth",4},{"4th",4},
            {"fifth",5},{"5th",5},
            {"sixth",6},{"6th",6},
            {"seventh",7},{"7th",7},
            {"eighth",8},{"8th",8},
            {"ninth",9},{"9th",9},
            {"tenth",10},{"10th",10},
            {"eleventh",11},{"11th",11},
            {"twelfth",12},{"12th",12},
        };

        /// <summary>
        /// Filler words that can appear between number words in speech and
        /// should be ignored during conversion.
        /// </summary>
        private static readonly HashSet<string> Fillers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "and", "a",
            // spoken punctuation — already replaced before reaching WordsToNumber,
            // but listed here as a safety net in case they appear in isolation
            "colon", "dash", "hyphen",
        };

        // -----------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------

        /// <summary>
        /// Converts a spoken ordinal or cardinal to its numeric prefix digit.
        /// Used for book-name prefixes: "first" → "1", "second" → "2", "3rd" → "3".
        /// Returns null if the word is not recognised as an ordinal/cardinal 1–3.
        /// </summary>
        public static string OrdinalToDigit(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return null;
            word = word.Trim().ToLowerInvariant();

            if (Ordinals.TryGetValue(word, out int ordVal) && ordVal >= 1 && ordVal <= 3)
                return ordVal.ToString();

            if (Cardinals.TryGetValue(word, out int cardVal) && cardVal >= 1 && cardVal <= 3)
                return cardVal.ToString();

            return null;
        }

        /// <summary>
        /// Returns true if the token is a word that represents a number
        /// (cardinal, ordinal, or filler like "and").
        /// </summary>
        public static bool IsNumberWord(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return false;
            word = word.Trim().ToLowerInvariant();
            return Cardinals.ContainsKey(word) || Ordinals.ContainsKey(word) || Fillers.Contains(word);
        }

        /// <summary>
        /// Converts a sequence of spoken number words into an integer.
        /// Examples:
        ///   "three"                    → 3
        ///   "sixteen"                  → 16
        ///   "twenty one"               → 21
        ///   "one hundred"              → 100
        ///   "one hundred and three"    → 103
        ///   "one hundred twenty one"   → 121
        ///   "one oh five"              → 105   (colloquial "1-0-5")
        ///
        /// Returns null if the input cannot be interpreted as a number.
        /// </summary>
        public static int? WordsToNumber(string spokenPhrase)
        {
            if (string.IsNullOrWhiteSpace(spokenPhrase)) return null;

            // If the phrase is already a pure digit string, parse directly
            string trimmed = spokenPhrase.Trim();
            if (int.TryParse(trimmed, out int directNum))
                return directNum;

            var tokens = Tokenise(trimmed);
            if (tokens.Count == 0) return null;

            // Try the "oh" pattern first: "one oh five" → digit-by-digit
            int? ohResult = TryOhPattern(tokens);
            if (ohResult.HasValue) return ohResult;

            // Standard English number parsing
            return ParseStandardNumber(tokens);
        }

        /// <summary>
        /// Takes a raw spoken fragment that should represent a Bible chapter:verse
        /// reference and returns a normalised digit string.
        /// 
        /// The input may contain filler words like "chapter", "verse", "verses",
        /// "and", "from", "to", "through" which are stripped or used as delimiters.
        ///
        /// Examples:
        ///   "three sixteen"                       → "3:16"
        ///   "chapter three verse sixteen"          → "3:16"
        ///   "twenty three"                         → "23"   (chapter only)
        ///   "three sixteen to eighteen"            → "3:16-18"
        ///   "three sixteen and seventeen"          → "3:16-17"
        ///   "one oh five"                          → "1:05"  (chapter 1, verse 05 → "1:5")
        ///   "fifty three five"                     → "53:5"
        ///   "one hundred and nineteen verse one oh five" → "119:105"
        ///
        /// Returns null if the input cannot be parsed into a valid reference fragment.
        /// </summary>
        public static string SpokenToReferenceFragment(string spoken)
        {
            if (string.IsNullOrWhiteSpace(spoken)) return null;

            string normalised = spoken.Trim().ToLowerInvariant();

            // Remove filler words that indicate structure but don't carry number value
            // We keep "and" and "to" as they serve as delimiters
            normalised = Regex.Replace(normalised, @"\b(chapter|chapters)\b", " ", RegexOptions.IgnoreCase);
            normalised = Regex.Replace(normalised, @"\b(verse|verses)\b", "VERSESEP", RegexOptions.IgnoreCase);
            normalised = Regex.Replace(normalised, @"\b(from)\b", " ", RegexOptions.IgnoreCase);
            normalised = Regex.Replace(normalised, @"\b(through)\b", "to", RegexOptions.IgnoreCase);

            // Spoken punctuation from the Vosk grammar vocabulary:
            //   "zechariah nine colon eight dash ten" → chapter 9, verse 8–10
            normalised = Regex.Replace(normalised, @"\bcolon\b", "VERSESEP", RegexOptions.IgnoreCase);
            normalised = Regex.Replace(normalised, @"\b(dash|hyphen)\b", "to", RegexOptions.IgnoreCase);

            // Collapse whitespace
            normalised = Regex.Replace(normalised, @"\s+", " ").Trim();

            // Split on "to" (range indicator) to detect verse ranges
            // "three sixteen to eighteen" → ["three sixteen", "eighteen"]
            string rangeEnd = null;
            var toParts = Regex.Split(normalised, @"\bto\b", RegexOptions.IgnoreCase);
            if (toParts.Length == 2)
            {
                normalised = toParts[0].Trim();
                string endPart = toParts[1].Trim();
                int? endNum = WordsToNumber(endPart);
                if (endNum.HasValue)
                    rangeEnd = endNum.Value.ToString();
            }

            // Also detect "X and Y" pattern for consecutive verses:
            // "three sixteen and seventeen" → "3:16-17"
            if (rangeEnd == null)
            {
                // Only match "and" that's between two number groups at the end
                var andMatch = Regex.Match(normalised, @"^(.+?)\band\b\s*(\w+)\s*$", RegexOptions.IgnoreCase);
                if (andMatch.Success)
                {
                    string beforeAnd = andMatch.Groups[1].Value.Trim();
                    string afterAnd = andMatch.Groups[2].Value.Trim();
                    int? afterNum = WordsToNumber(afterAnd);
                    if (afterNum.HasValue && !string.IsNullOrWhiteSpace(beforeAnd))
                    {
                        normalised = beforeAnd;
                        rangeEnd = afterNum.Value.ToString();
                    }
                }
            }

            // Now parse the main part.
            // Split on VERSESEP (from "verse" keyword) to find chapter vs verse boundary
            string chapterStr = null;
            string verseStr = null;

            if (normalised.Contains("VERSESEP"))
            {
                var vsParts = normalised.Split(new[] { "VERSESEP" }, StringSplitOptions.RemoveEmptyEntries);
                if (vsParts.Length >= 1)
                {
                    int? ch = WordsToNumber(vsParts[0].Trim());
                    if (ch.HasValue) chapterStr = ch.Value.ToString();
                }
                if (vsParts.Length >= 2)
                {
                    int? vs = WordsToNumber(vsParts[1].Trim());
                    if (vs.HasValue) verseStr = vs.Value.ToString();
                }
            }
            else
            {
                // No explicit "verse" keyword — heuristic split.
                // Try to interpret the spoken numbers as chapter + verse.
                // Strategy: tokenise, try splitting at each possible boundary.
                var result = SplitChapterVerse(normalised);
                chapterStr = result.chapter;
                verseStr = result.verse;
            }

            if (chapterStr == null) return null;

            // Build the reference string
            string reference;
            if (verseStr != null)
            {
                reference = $"{chapterStr}:{verseStr}";
            }
            else
            {
                reference = chapterStr;
            }

            // Append range end if present
            if (rangeEnd != null)
            {
                reference += $"-{rangeEnd}";
            }

            return reference;
        }

        // -----------------------------------------------------------------------
        // Private helpers
        // -----------------------------------------------------------------------

        private static List<string> Tokenise(string text)
        {
            return Regex.Split(text.ToLowerInvariant(), @"[\s\-]+")
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();
        }

        /// <summary>
        /// Detects the colloquial "oh" pattern: "one oh five" = 1, 0, 5 → 105.
        /// Returns null if the pattern doesn't match.
        /// </summary>
        private static int? TryOhPattern(List<string> tokens)
        {
            // The "oh" pattern requires at least 3 tokens and "oh"/"o"/"zero" in position ≥1
            bool hasOh = tokens.Any(t => t == "oh" || t == "o" || t == "zero");
            if (!hasOh || tokens.Count < 3) return null;

            // Every token must be a single-digit cardinal (0-9) or a known filler
            var digits = new List<int>();
            foreach (var token in tokens)
            {
                if (Fillers.Contains(token)) continue;
                if (Cardinals.TryGetValue(token, out int val) && val >= 0 && val <= 9)
                {
                    digits.Add(val);
                }
                else
                {
                    return null; // Not a digit-by-digit pattern
                }
            }

            if (digits.Count < 2) return null;

            int result = 0;
            foreach (int d in digits)
                result = result * 10 + d;

            return result;
        }

        /// <summary>
        /// Standard English number parsing for values 0–199.
        /// </summary>
        private static int? ParseStandardNumber(List<string> tokens)
        {
            int total = 0;
            bool foundAny = false;
            bool prevWasHundred = false;

            for (int i = 0; i < tokens.Count; i++)
            {
                string t = tokens[i];

                if (Fillers.Contains(t))
                {
                    continue;
                }

                if (!Cardinals.TryGetValue(t, out int val))
                {
                    // Also accept ordinals for standalone number parsing
                    if (Ordinals.TryGetValue(t, out int ordVal))
                    {
                        total += ordVal;
                        foundAny = true;
                        continue;
                    }
                    return null; // Unknown word
                }

                if (val == 100)
                {
                    // "one hundred" or just "hundred"
                    if (total == 0) total = 1;
                    total *= 100;
                    prevWasHundred = true;
                }
                else
                {
                    total += val;
                    prevWasHundred = false;
                }

                foundAny = true;
            }

            return foundAny ? total : (int?)null;
        }

        /// <summary>
        /// Heuristic split of spoken number tokens into chapter and verse parts.
        /// 
        /// Strategy:
        /// 1. Try each possible split point (left = chapter, right = verse).
        /// 2. Both sides must parse as valid numbers.
        /// 3. Prefer the split where chapter is smallest sensible value (i.e., leftmost split).
        ///    In Bible references, chapters are typically 1-150 and verses 1-176.
        /// 4. If no valid split produces two numbers, treat the whole thing as chapter-only.
        /// </summary>
        private static (string chapter, string verse) SplitChapterVerse(string spokenFragment)
        {
            var tokens = Tokenise(spokenFragment);
            if (tokens.Count == 0) return ((string)null, (string)null);

            // Single token → chapter only
            if (tokens.Count == 1)
            {
                int? val = WordsToNumber(tokens[0]);
                return val.HasValue ? (val.Value.ToString(), (string)null) : ((string)null, (string)null);
            }

            // Try each split point from left to right.
            // Prefer the first valid split (smallest chapter number).
            // NOTE: we do NOT blindly skip splits where the right side starts with a filler
            // like "and" — "twenty and six" needs split before "and", and WordsToNumber
            // handles interior "and" correctly ("one hundred and three" -> 103).
            for (int splitAt = 1; splitAt < tokens.Count; splitAt++)
            {
                string leftPhrase = string.Join(" ", tokens.Take(splitAt));
                string rightPhrase = string.Join(" ", tokens.Skip(splitAt));

                string firstRight = tokens[splitAt];

                // "hundred" alone on the right makes no sense for a verse
                if (firstRight == "hundred") continue;

                // If the right side is nothing but fillers, absorb into left — don't split here
                bool rightIsOnlyFiller = Fillers.Contains(firstRight) &&
                                         tokens.Skip(splitAt).All(t => Fillers.Contains(t));
                if (rightIsOnlyFiller) continue;

                int? leftNum = WordsToNumber(leftPhrase);
                int? rightNum = WordsToNumber(rightPhrase);

                if (leftNum.HasValue && rightNum.HasValue && leftNum.Value > 0 && rightNum.Value > 0)
                {
                    return (leftNum.Value.ToString(), rightNum.Value.ToString());
                }
            }

            // No valid split found — treat the whole thing as chapter-only
            int? wholeNum = WordsToNumber(spokenFragment);
            return wholeNum.HasValue ? (wholeNum.Value.ToString(), (string)null) : ((string)null, (string)null);
        }
    }
}
