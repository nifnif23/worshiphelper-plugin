// ============================================================================
// ReferenceValidator.cs
// Validates detected scripture references against real Bible data and repairs
// them when they don't exist, using a priority-ordered set of heuristics
// tuned for Nigerian-accented English speech patterns.
//
// Problems handled:
//
//   1. Verse as range   "Zechariah 5:49"  → "Zechariah 5:4-9"
//      Verse 49 doesn't exist in Zech 5 (max 11). Digits 4 and 9 both exist
//      as verses. Treat the verse number as a concatenated range.
//
//   2. Collapsed chapter+verse  "Zechariah 49" → "Zechariah 4:9"
//      Chapter 49 doesn't exist (max 14). Split digits into chapter:verse.
//
//   3. -teen → -ty confusion   "Zechariah 14:14" vs "Zechariah 40:14"
//      Nigerian accent often drops the "-teen" ending. fourteen → forty,
//      sixteen → sixty, etc. If X0 doesn't exist but X+teen does, swap.
//
//   4. Digit-by-digit high Psalms  "Psalm one one nine" → Psalm 119
//      Already handled upstream in SpokenNumberConverter, but we verify
//      the result exists and don't accidentally split 119 into 1:19.
//
//   5. Multi-digit verse range  "Zechariah 5:149" → "Zechariah 5:14-9"
//      Try all digit split points and pick the one yielding a valid range.
//
// All repairs are validated against the loaded Bible object — we never
// produce a reference that doesn't exist in the actual data.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using log4net;

namespace WorshipHelperVSTO
{
    /// <summary>
    /// Result from ReferenceValidator.Validate().
    /// </summary>
    public class ValidatedReference
    {
        /// <summary>Canonical book name, e.g. "Zechariah".</summary>
        public string BookName { get; set; }

        /// <summary>Normalised reference string, e.g. "5:4-9" or "4:9".</summary>
        public string ReferenceFragment { get; set; }

        /// <summary>Full normalised reference ready to pass downstream, e.g. "Zechariah 5:4-9".</summary>
        public string NormalisedReference => $"{BookName} {ReferenceFragment}";

        /// <summary>How the reference was resolved.</summary>
        public ValidationOutcome Outcome { get; set; }

        /// <summary>The original (pre-repair) reference string.</summary>
        public string OriginalReference { get; set; }
    }

    public enum ValidationOutcome
    {
        /// <summary>Reference existed as-is in the Bible data. No repair needed.</summary>
        Valid,

        /// <summary>Verse number was split into a range (e.g. 49 → 4-9).</summary>
        RepairedVerseRange,

        /// <summary>Chapter+verse were collapsed; digits split at best point (e.g. 49 → 4:9).</summary>
        RepairedChapterVerseSplit,

        /// <summary>-teen/-ty confusion repaired (e.g. forty → fourteen).</summary>
        RepairedTeenTySwap,

        /// <summary>Chapter or verse snapped to nearest existing value.</summary>
        RepairedSnapped,

        /// <summary>Could not repair — reference is unrecoverable.</summary>
        Invalid,
    }

    public static class ReferenceValidator
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(ReferenceValidator));

        // -teen counterparts for -ty numbers that might be misheard
        // key = the "wrong" -ty value, value = the "-teen" it probably was
        private static readonly Dictionary<int, int> TyToTeen = new Dictionary<int, int>
        {
            { 20, 12 },   // twenty  ← twelve  (less common but possible)
            { 30, 13 },   // thirty  ← thirteen
            { 40, 14 },   // forty   ← fourteen  (most common Nigerian accent swap)
            { 50, 15 },   // fifty   ← fifteen
            { 60, 16 },   // sixty   ← sixteen
            { 70, 17 },   // seventy ← seventeen
            { 80, 18 },   // eighty  ← eighteen
            { 90, 19 },   // ninety  ← nineteen
        };

        // -----------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------

        /// <summary>
        /// Validates a detected reference against the loaded Bible and attempts
        /// repairs if it doesn't exist. Returns null if totally unrecoverable.
        ///
        /// Parameters:
        ///   bible      — loaded Bible object (used for real chapter/verse counts)
        ///   bookName   — canonical book name (already resolved by detector)
        ///   refFragment — the "chapter:verse" or "chapter" part, e.g. "5:49", "49"
        /// </summary>
        public static ValidatedReference Validate(Bible bible, string bookName, string refFragment)
        {
            if (bible == null || string.IsNullOrWhiteSpace(bookName) ||
                string.IsNullOrWhiteSpace(refFragment))
                return null;

            var book = bible.books.FirstOrDefault(b =>
                b.name.Equals(bookName, StringComparison.OrdinalIgnoreCase));
            if (book == null) return null;

            string original = $"{bookName} {refFragment}";

            // Parse the fragment into its components
            if (!TryParseFragment(refFragment, out int chapter, out int? verse, out int? verseEnd))
            {
                log.Warn($"ReferenceValidator: could not parse fragment \"{refFragment}\"");
                return null;
            }

            // --- Attempt 1: reference is valid as-is ---
            if (ReferenceExists(book, chapter, verse, verseEnd))
            {
                return new ValidatedReference
                {
                    BookName = book.name,
                    ReferenceFragment = refFragment,
                    Outcome = ValidationOutcome.Valid,
                    OriginalReference = original,
                };
            }

            log.Debug($"ReferenceValidator: \"{original}\" not valid, attempting repairs...");

            // --- Attempt 2: verse-as-range (the main case: 5:49 → 5:4-9) ---
            if (verse.HasValue && verseEnd == null)
            {
                var rangeResult = TryRepairVerseAsRange(book, chapter, verse.Value, original);
                if (rangeResult != null) return rangeResult;
            }

            // --- Attempt 3: multi-digit verse range (5:149 → 5:14-9 or 5:1-49) ---
            if (verse.HasValue && verseEnd == null && verse.Value >= 10)
            {
                var multiResult = TryRepairMultiDigitRange(book, chapter, verse.Value, original);
                if (multiResult != null) return multiResult;
            }

            // --- Attempt 4: chapter+verse collapsed, no colon (e.g. "49" → 4:9) ---
            if (verse == null)
            {
                var splitResult = TryRepairChapterVerseSplit(book, chapter, original);
                if (splitResult != null) return splitResult;
            }

            // --- Attempt 5: also try splitting even when colon present but chapter invalid ---
            if (verse.HasValue && !ChapterExists(book, chapter))
            {
                // Maybe the whole "chapter:verse" is really a collapsed number
                // e.g. "Zech 1:49" where chapter 1 is valid but verse 49 isn't,
                // and "149" → 14:9 makes more sense
                string collapsed = $"{chapter}{verse.Value}";
                if (int.TryParse(collapsed, out int collapsedNum))
                {
                    var splitResult = TryRepairChapterVerseSplit(book, collapsedNum, original);
                    if (splitResult != null) return splitResult;
                }
            }

            // --- Attempt 6: -teen/-ty swap ---
            var teenResult = TryRepairTeenTySwap(book, chapter, verse, verseEnd, original);
            if (teenResult != null) return teenResult;

            // --- Attempt 7: snap to nearest existing reference ---
            var snapped = TrySnap(book, chapter, verse, verseEnd, original);
            if (snapped != null) return snapped;

            log.Warn($"ReferenceValidator: could not repair \"{original}\"");
            return null;
        }

        // -----------------------------------------------------------------------
        // Repair strategies
        // -----------------------------------------------------------------------

        /// <summary>
        /// Strategy 2: verse number as concatenated range digits.
        /// "Zech 5:49" — verse 49 invalid, but 4 and 9 both exist → "Zech 5:4-9"
        /// Works for 2-digit verse numbers only (XY → X-Y).
        /// </summary>
        private static ValidatedReference TryRepairVerseAsRange(
            Book book, int chapter, int verseNum, string original)
        {
            if (verseNum < 10 || verseNum > 99) return null;

            int lo = verseNum / 10;
            int hi = verseNum % 10;

            if (lo < 1 || hi < 1 || lo >= hi) return null;

            if (ReferenceExists(book, chapter, lo, hi))
            {
                string fragment = $"{chapter}:{lo}-{hi}";
                log.Info($"ReferenceValidator: repaired verse-as-range \"{original}\" → \"{book.name} {fragment}\"");
                return new ValidatedReference
                {
                    BookName = book.name,
                    ReferenceFragment = fragment,
                    Outcome = ValidationOutcome.RepairedVerseRange,
                    OriginalReference = original,
                };
            }
            return null;
        }

        /// <summary>
        /// Strategy 3: multi-digit verse split into range.
        /// "Zech 5:149" → try "5:14-9", "5:1-49" — pick the valid ascending one.
        /// </summary>
        private static ValidatedReference TryRepairMultiDigitRange(
            Book book, int chapter, int verseNum, string original)
        {
            string verseStr = verseNum.ToString();
            if (verseStr.Length < 2) return null;

            // Try each split point
            for (int split = 1; split < verseStr.Length; split++)
            {
                if (!int.TryParse(verseStr.Substring(0, split), out int lo)) continue;
                if (!int.TryParse(verseStr.Substring(split), out int hi)) continue;

                if (lo < 1 || hi < 1 || lo >= hi) continue;

                if (ReferenceExists(book, chapter, lo, hi))
                {
                    string fragment = $"{chapter}:{lo}-{hi}";
                    log.Info($"ReferenceValidator: repaired multi-digit range \"{original}\" → \"{book.name} {fragment}\"");
                    return new ValidatedReference
                    {
                        BookName = book.name,
                        ReferenceFragment = fragment,
                        Outcome = ValidationOutcome.RepairedVerseRange,
                        OriginalReference = original,
                    };
                }
            }
            return null;
        }

        /// <summary>
        /// Strategy 4: chapter+verse collapsed (no colon spoken).
        /// "Zech 49" → chapter 49 invalid, try splitting digits as chapter:verse.
        /// Prefers the split where BOTH parts exist. Tries left-to-right splits.
        /// Special case: Psalms 119 must not be split to 11:9 — if the chapter
        /// exists we never split it.
        /// </summary>
        private static ValidatedReference TryRepairChapterVerseSplit(
            Book book, int collapsed, string original)
        {
            // If the chapter actually exists (e.g. Psalms 119), don't split it
            if (ChapterExists(book, collapsed)) return null;

            string digits = collapsed.ToString();
            if (digits.Length < 2) return null;

            // Try each split point, prefer shortest chapter (leftmost split)
            for (int split = 1; split < digits.Length; split++)
            {
                if (!int.TryParse(digits.Substring(0, split), out int ch)) continue;
                if (!int.TryParse(digits.Substring(split), out int vs)) continue;

                if (ch < 1 || vs < 1) continue;

                if (ReferenceExists(book, ch, vs, null))
                {
                    string fragment = $"{ch}:{vs}";
                    log.Info($"ReferenceValidator: repaired chapter+verse split \"{original}\" → \"{book.name} {fragment}\"");
                    return new ValidatedReference
                    {
                        BookName = book.name,
                        ReferenceFragment = fragment,
                        Outcome = ValidationOutcome.RepairedChapterVerseSplit,
                        OriginalReference = original,
                    };
                }
            }
            return null;
        }

        /// <summary>
        /// Strategy 6: -teen/-ty swap.
        /// Chapter 40 doesn't exist but chapter 14 does? Swap forty → fourteen.
        /// Applied to both chapter and verse numbers.
        /// </summary>
        private static ValidatedReference TryRepairTeenTySwap(
            Book book, int chapter, int? verse, int? verseEnd, string original)
        {
            // Try swapping the chapter
            if (TyToTeen.TryGetValue(chapter, out int teenChapter))
            {
                if (ReferenceExists(book, teenChapter, verse, verseEnd))
                {
                    string fragment = BuildFragment(teenChapter, verse, verseEnd);
                    log.Info($"ReferenceValidator: repaired -teen/-ty chapter \"{original}\" → \"{book.name} {fragment}\"");
                    return new ValidatedReference
                    {
                        BookName = book.name,
                        ReferenceFragment = fragment,
                        Outcome = ValidationOutcome.RepairedTeenTySwap,
                        OriginalReference = original,
                    };
                }
            }

            // Try swapping the verse
            if (verse.HasValue && TyToTeen.TryGetValue(verse.Value, out int teenVerse))
            {
                if (ReferenceExists(book, chapter, teenVerse, verseEnd))
                {
                    string fragment = BuildFragment(chapter, teenVerse, verseEnd);
                    log.Info($"ReferenceValidator: repaired -teen/-ty verse \"{original}\" → \"{book.name} {fragment}\"");
                    return new ValidatedReference
                    {
                        BookName = book.name,
                        ReferenceFragment = fragment,
                        Outcome = ValidationOutcome.RepairedTeenTySwap,
                        OriginalReference = original,
                    };
                }
            }

            // Try swapping the verse end
            if (verseEnd.HasValue && TyToTeen.TryGetValue(verseEnd.Value, out int teenEnd))
            {
                if (ReferenceExists(book, chapter, verse, teenEnd))
                {
                    string fragment = BuildFragment(chapter, verse, teenEnd);
                    log.Info($"ReferenceValidator: repaired -teen/-ty range end \"{original}\" → \"{book.name} {fragment}\"");
                    return new ValidatedReference
                    {
                        BookName = book.name,
                        ReferenceFragment = fragment,
                        Outcome = ValidationOutcome.RepairedTeenTySwap,
                        OriginalReference = original,
                    };
                }
            }

            return null;
        }

        /// <summary>
        /// Strategy 7: snap chapter and/or verse to the nearest existing value.
        /// Last resort — clamps out-of-bounds numbers rather than dropping the reference.
        /// e.g. "Zech 5:20" → Zech 5 has 11 verses → snap to "Zech 5:11"
        /// </summary>
        private static ValidatedReference TrySnap(
            Book book, int chapter, int? verse, int? verseEnd, string original)
        {
            // Snap chapter
            int maxChapter = book.chapters.Max(c => c.number);
            int snappedChapter = Math.Max(1, Math.Min(chapter, maxChapter));

            var ch = book.chapters.FirstOrDefault(c => c.number == snappedChapter);
            if (ch == null) return null;

            if (verse == null)
            {
                // Chapter-only reference — chapter was out of range, snapped
                if (snappedChapter == chapter) return null; // was already valid, shouldn't be here
                string fragment = snappedChapter.ToString();
                log.Info($"ReferenceValidator: snapped chapter \"{original}\" → \"{book.name} {fragment}\"");
                return new ValidatedReference
                {
                    BookName = book.name,
                    ReferenceFragment = fragment,
                    Outcome = ValidationOutcome.RepairedSnapped,
                    OriginalReference = original,
                };
            }

            int maxVerse = ch.verses.Max(v => v.number);
            int snappedVerse = Math.Max(1, Math.Min(verse.Value, maxVerse));
            int? snappedEnd = verseEnd.HasValue
                ? (int?)Math.Max(snappedVerse, Math.Min(verseEnd.Value, maxVerse))
                : null;

            // Only snap if we actually changed something
            if (snappedChapter == chapter && snappedVerse == verse.Value &&
                snappedEnd == verseEnd) return null;

            string frag = BuildFragment(snappedChapter, snappedVerse, snappedEnd);
            log.Info($"ReferenceValidator: snapped \"{original}\" → \"{book.name} {frag}\"");
            return new ValidatedReference
            {
                BookName = book.name,
                ReferenceFragment = frag,
                Outcome = ValidationOutcome.RepairedSnapped,
                OriginalReference = original,
            };
        }

        // -----------------------------------------------------------------------
        // Helpers
        // -----------------------------------------------------------------------

        private static bool TryParseFragment(string fragment,
            out int chapter, out int? verse, out int? verseEnd)
        {
            chapter = 0; verse = null; verseEnd = null;

            // Expected formats: "5", "5:4", "5:4-9"
            var match = Regex.Match(fragment.Trim(),
                @"^(\d+)(?::(\d+)(?:-(\d+))?)?$");
            if (!match.Success) return false;

            chapter = int.Parse(match.Groups[1].Value);
            if (match.Groups[2].Success)
                verse = int.Parse(match.Groups[2].Value);
            if (match.Groups[3].Success)
                verseEnd = int.Parse(match.Groups[3].Value);

            return true;
        }

        private static bool ChapterExists(Book book, int chapter)
            => book.chapters.Any(c => c.number == chapter);

        private static bool VerseExists(Chapter ch, int verse)
            => ch.verses.Any(v => v.number == verse);

        private static bool ReferenceExists(Book book, int chapter, int? verse, int? verseEnd)
        {
            var ch = book.chapters.FirstOrDefault(c => c.number == chapter);
            if (ch == null) return false;
            if (verse == null) return true; // chapter-only

            if (!VerseExists(ch, verse.Value)) return false;
            if (verseEnd.HasValue)
            {
                if (verseEnd.Value <= verse.Value) return false;
                if (!VerseExists(ch, verseEnd.Value)) return false;
            }
            return true;
        }

        private static string BuildFragment(int chapter, int? verse, int? verseEnd)
        {
            if (verse == null) return chapter.ToString();
            if (verseEnd == null) return $"{chapter}:{verse}";
            return $"{chapter}:{verse}-{verseEnd}";
        }
    }
}
