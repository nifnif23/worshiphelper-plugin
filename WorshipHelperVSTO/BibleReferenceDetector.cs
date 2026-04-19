// ============================================================================
// BibleReferenceDetector.cs
// Detects and normalises Bible references from raw speech-to-text output.
//
// Takes input like:
//   "let's turn to first corinthians thirteen verse four"
// and produces:
//   "1 Corinthians 13:4"
//
// Drop into:  WorshipHelperVSTO/BibleReferenceDetector.cs
// Namespace:  WorshipHelperVSTO
// ============================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace WorshipHelperVSTO
{
    /// <summary>
    /// Result of a Bible reference detection attempt.
    /// </summary>
    public class DetectedReference
    {
        /// <summary>
        /// The fully normalised reference string, e.g. "John 3:16" or "1 Corinthians 13:4-7".
        /// Ready to be passed into InsertScripture / FullReferenceParser.ParseFullReference.
        /// </summary>
        public string NormalisedReference { get; set; }

        /// <summary>
        /// The canonical book name (e.g. "John", "1 Corinthians", "Psalms").
        /// </summary>
        public string BookName { get; set; }

        /// <summary>
        /// The numeric reference portion (e.g. "3:16", "13:4-7", "23").
        /// </summary>
        public string ReferenceFragment { get; set; }

        /// <summary>
        /// The raw text that was matched and consumed from the spoken input.
        /// Useful for logging / debugging.
        /// </summary>
        public string MatchedRawText { get; set; }

        /// <summary>
        /// Confidence score (0.0 – 1.0). Higher means more likely a real reference.
        /// Based on heuristics: book name match quality, presence of numbers, etc.
        /// </summary>
        public double Confidence { get; set; }
    }

    /// <summary>
    /// Stateless detector that scans raw recognised speech text for Bible references
    /// and returns normalised reference strings.
    ///
    /// Conservative by design: only fires when a Bible book name is clearly
    /// identified followed by plausible chapter/verse numbers.
    /// </summary>
    public static class BibleReferenceDetector
    {
        // -----------------------------------------------------------------------
        // Bible book name data
        // -----------------------------------------------------------------------

        /// <summary>
        /// All 66 canonical Bible book names.
        /// Multi-word entries (e.g. "song of solomon") are included.
        /// </summary>
        private static readonly List<BookEntry> Books = BuildBookList();

        private class BookEntry
        {
            public string Canonical;       // e.g. "1 Corinthians"
            public List<string> Variants;  // spoken forms: "first corinthians", "1 corinthians", "1st corinthians"

            /// <summary>
            /// Max number of tokens in any variant (used for greedy matching).
            /// </summary>
            public int MaxTokens => Variants.Max(v => v.Split(' ').Length);
        }

        // -----------------------------------------------------------------------
        // Filler / preamble words that commonly precede a Bible reference in speech
        // -----------------------------------------------------------------------
        private static readonly HashSet<string> PreambleWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "read", "reading", "turn", "turning", "go", "going", "look", "looking",
            "open", "opening", "find", "finding", "flip", "flipping",
            "let's", "lets", "let", "us", "please", "now", "okay", "ok",
            "to", "at", "in", "from", "the", "book", "passage",
            // NOTE: "of" intentionally removed — it breaks "Song of Solomon" matching
            "scripture", "text", "it's", "its", "says", "we're", "were",
            "i'm", "im", "today", "tonight", "this", "morning", "evening",
        };

        // -----------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------

        /// <summary>
        /// Minimum confidence threshold for a detection to be considered valid.
        /// Can be adjusted for noisier environments (raise) or controlled settings (lower).
        /// Default: 0.5
        /// </summary>
        public static double MinConfidence { get; set; } = 0.4; // Lowered from 0.5 — fuzzy book name matches score ~0.4-0.5

        /// <summary>
        /// Attempts to detect one or more Bible references in the given raw text.
        /// Returns an empty list if nothing is found.
        ///
        /// The detector is conservative: it requires a recognisable Bible book name
        /// followed by at least one number-like token.
        /// </summary>
        public static List<DetectedReference> Detect(string rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText)) return new List<DetectedReference>();

            var results = new List<DetectedReference>();

            // Normalise: lowercase, remove punctuation except hyphens and apostrophes
            string cleaned = rawText.Trim().ToLowerInvariant();
            cleaned = Regex.Replace(cleaned, @"[^\w\s'\-]", " ");
            cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();

            var tokens = cleaned.Split(' ').ToList();

            int pos = 0;
            while (pos < tokens.Count)
            {
                // Skip preamble words
                if (PreambleWords.Contains(tokens[pos]))
                {
                    pos++;
                    continue;
                }

                // Try to match a Bible book name starting at this position
                var (book, tokensConsumed, wasFuzzy) = TryMatchBook(tokens, pos);
                if (book != null)
                {
                    int afterBook = pos + tokensConsumed;

                    // Collect the remaining tokens that look like numbers / reference words
                    int refStart = afterBook;
                    int refEnd = refStart;

                    // Skip an optional "chapter" keyword right after the book name
                    if (refEnd < tokens.Count && (tokens[refEnd] == "chapter" || tokens[refEnd] == "chapters"))
                        refEnd++;

                    while (refEnd < tokens.Count && IsReferenceToken(tokens[refEnd]))
                    {
                        refEnd++;
                    }

                    string rawMatch = string.Join(" ", tokens.Skip(pos).Take(refEnd - pos));

                    if (refEnd > refStart)
                    {
                        // We have number tokens after the book name
                        string spokenRef = string.Join(" ", tokens.Skip(refStart).Take(refEnd - refStart));
                        string refFragment = SpokenNumberConverter.SpokenToReferenceFragment(spokenRef);

                        if (refFragment != null)
                        {
                            double confidence = ComputeConfidence(book, refFragment, tokensConsumed, wasFuzzy);
                            if (confidence >= MinConfidence)
                            {
                                string normalised = $"{book.Canonical} {refFragment}";
                                results.Add(new DetectedReference
                                {
                                    NormalisedReference = normalised,
                                    BookName = book.Canonical,
                                    ReferenceFragment = refFragment,
                                    MatchedRawText = rawMatch,
                                    Confidence = confidence,
                                });
                            }
                        }
                    }
                    else
                    {
                        // Book name only — could be a chapter-only reference for single-chapter books
                        // (Philemon, Jude, Obadiah, 2 John, 3 John)
                        // We skip these as they're too ambiguous in speech.
                    }

                    pos = Math.Max(pos + 1, refEnd);
                }
                else
                {
                    pos++;
                }
            }

            return results;
        }

        /// <summary>
        /// Convenience method: detects the single best (highest-confidence) reference.
        /// Returns null if nothing is detected.
        /// </summary>
        public static DetectedReference DetectBest(string rawText)
        {
            var all = Detect(rawText);
            return all.OrderByDescending(r => r.Confidence).FirstOrDefault();
        }

        // -----------------------------------------------------------------------
        // Book name matching
        // -----------------------------------------------------------------------

        /// <summary>
        /// Tries to match a Bible book name starting at the given token position.
        /// Uses greedy matching (longest match wins).
        /// Returns (BookEntry, tokensConsumed) or (null, 0).
        /// </summary>
        private static (BookEntry book, int tokensConsumed, bool wasFuzzy) TryMatchBook(List<string> tokens, int startPos)
        {
            BookEntry bestMatch = null;
            int bestLength = 0;
            bool bestWasExact = false;
            bool bestHadFuzzy = false;

            foreach (var book in Books)
            {
                foreach (var variant in book.Variants)
                {
                    var varTokens = variant.Split(' ');
                    int len = varTokens.Length;

                    if (startPos + len > tokens.Count) continue;
                    if (len < bestLength) continue; // Only interested in longer or equal matches

                    bool exactMatch = true;
                    bool fuzzyMatch = true;
                    int fuzzyErrors = 0;

                    for (int i = 0; i < len; i++)
                    {
                        string spoken = tokens[startPos + i];
                        string expected = varTokens[i];

                        if (!string.Equals(spoken, expected, StringComparison.OrdinalIgnoreCase))
                        {
                            exactMatch = false;

                            // Fuzzy matching for book name words.
                            // Short tokens (digits, "of", "1st") must be exact — only fuzz words of 5+ chars.
                            // Allowed edit distance scales with word length so obscure long book names
                            // like "Habakkuk" (8) or "Thessalonians" (13) get more slack than short ones.
                            //   5–7 chars  → allow 2 edits  (e.g. "isaiah" → "isaian")
                            //   8–10 chars → allow 3 edits  (e.g. "habakkuk" → "habacuc")
                            //   11+ chars  → allow 4 edits  (e.g. "thessalonians" → "thesalonians")
                            int allowedEdits = expected.Length >= 11 ? 4
                                            : expected.Length >= 8  ? 3
                                            : expected.Length >= 5  ? 2
                                            : 0;

                            if (allowedEdits > 0 && LevenshteinDistance(
                                    spoken.ToLowerInvariant(), expected.ToLowerInvariant()) <= allowedEdits)
                            {
                                fuzzyErrors++;
                                // Allow at most one fuzzy token per variant to avoid false positives
                                if (fuzzyErrors > 1) { fuzzyMatch = false; break; }
                            }
                            else
                            {
                                fuzzyMatch = false;
                                break;
                            }
                        }
                    }

                    bool isMatch = exactMatch || fuzzyMatch;
                    if (!isMatch) continue;

                    // Prefer exact over fuzzy; prefer longer over shorter
                    bool betterThanCurrent = (len > bestLength) ||
                                            (len == bestLength && exactMatch && !bestWasExact);
                    if (betterThanCurrent)
                    {
                        bestMatch = book;
                        bestLength = len;
                        bestWasExact = exactMatch;
                        bestHadFuzzy = !exactMatch;
                    }
                }
            }

            return (bestMatch, bestLength, bestHadFuzzy);
        }

        /// <summary>
        /// Standard Levenshtein edit distance. Used for fuzzy book name matching
        /// to handle speech engine mishearings like "corinthian" vs "corinthians".
        /// </summary>
        private static int LevenshteinDistance(string a, string b)
        {
            if (string.IsNullOrEmpty(a)) return b?.Length ?? 0;
            if (string.IsNullOrEmpty(b)) return a.Length;

            var d = new int[a.Length + 1, b.Length + 1];
            for (int i = 0; i <= a.Length; i++) d[i, 0] = i;
            for (int j = 0; j <= b.Length; j++) d[0, j] = j;

            for (int i = 1; i <= a.Length; i++)
            for (int j = 1; j <= b.Length; j++)
            {
                int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + cost);
            }

            return d[a.Length, b.Length];
        }

        // -----------------------------------------------------------------------
        // Reference token identification
        // -----------------------------------------------------------------------

        /// <summary>
        /// Returns true if the token looks like it could be part of a chapter:verse reference.
        /// </summary>
        private static bool IsReferenceToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return false;
            token = token.ToLowerInvariant();

            // Pure digit
            if (int.TryParse(token, out _)) return true;

            // Known number words
            if (SpokenNumberConverter.IsNumberWord(token)) return true;

            // Structural keywords in references
            var refKeywords = new HashSet<string> { "chapter", "chapters", "verse", "verses", "to", "through", "from", "and" };
            return refKeywords.Contains(token);
        }

        // -----------------------------------------------------------------------
        // Confidence scoring
        // -----------------------------------------------------------------------

        private static double ComputeConfidence(BookEntry book, string refFragment, int bookTokenCount, bool wasFuzzy = false)
        {
            double score = 0.4; // Base score for having a recognised book name

            // Longer book name matches are more confident
            if (bookTokenCount >= 2) score += 0.1;
            if (bookTokenCount >= 3) score += 0.1;

            // Having a colon (chapter:verse) increases confidence
            if (refFragment.Contains(":")) score += 0.2;

            // Having a range (dash) is very specific
            if (refFragment.Contains("-")) score += 0.1;

            // The reference fragment should have at least one digit
            if (refFragment.Any(char.IsDigit)) score += 0.1;

            // Fuzzy book name match: small penalty, but still accept it
            // (the engine misheard an uncommon word — trust it)
            if (wasFuzzy) score -= 0.1;

            return Math.Min(score, 1.0);
        }

        // -----------------------------------------------------------------------
        // Book list builder
        // -----------------------------------------------------------------------

        private static List<BookEntry> BuildBookList()
        {
            var list = new List<BookEntry>();

            // Helper to create a book entry with all its spoken variants
            void Add(string canonical, params string[] extraVariants)
            {
                var variants = new List<string> { canonical.ToLowerInvariant() };
                foreach (var v in extraVariants)
                    variants.Add(v.ToLowerInvariant());

                // For numbered books, auto-generate ordinal variants
                if (canonical.Length > 2 && char.IsDigit(canonical[0]) && canonical[1] == ' ')
                {
                    char digit = canonical[0];
                    string rest = canonical.Substring(2).ToLowerInvariant();

                    // "1 Corinthians" → "first corinthians", "1st corinthians"
                    string ordinalWord = digit == '1' ? "first" : digit == '2' ? "second" : digit == '3' ? "third" : null;
                    string ordinalSuffix = digit == '1' ? "1st" : digit == '2' ? "2nd" : digit == '3' ? "3rd" : null;

                    if (ordinalWord != null)
                    {
                        variants.Add($"{ordinalWord} {rest}");
                        variants.Add($"{ordinalSuffix} {rest}");
                    }
                }

                list.Add(new BookEntry { Canonical = canonical, Variants = variants.Distinct().ToList() });
            }

            // --- Old Testament ---
            Add("Genesis", "gen");
            Add("Exodus", "exod");
            Add("Leviticus", "lev", "leviticus");
            Add("Numbers", "num");
            Add("Deuteronomy", "deut", "deuteronomy", "deutronomy", "duteronomy");
            Add("Joshua", "josh");
            Add("Judges", "judg");
            Add("Ruth");
            Add("1 Samuel", "1 sam", "i samuel", "i sam");
            Add("2 Samuel", "2 sam", "ii samuel", "ii sam");
            Add("1 Kings", "1 king", "i kings", "i king");
            Add("2 Kings", "2 king", "ii kings", "ii king");
            Add("1 Chronicles", "1 chron", "i chronicles", "i chron");
            Add("2 Chronicles", "2 chron", "ii chronicles", "ii chron");
            Add("Ezra");
            Add("Nehemiah", "neh", "nehemia", "nehimiah", "nehimia");
            Add("Esther");
            Add("Job");
            Add("Psalms", "psalm", "psa", "salms", "sams");
            Add("Proverbs", "prov", "proverb", "proverbs");
            Add("Ecclesiastes", "eccl", "eccles", "ecclesiaste");
            Add("Song of Solomon", "song of songs", "song of sol", "songs of solomon", "solomon's song");
            Add("Isaiah", "isa", "isaiah", "esaiah", "isaia");
            Add("Jeremiah", "jer", "jeremia", "jerimiah", "jerimia");
            Add("Lamentations", "lam", "lamentation");
            Add("Ezekiel", "ezek", "ezekiel", "ezekia", "ezekel");
            Add("Daniel", "dan");
            Add("Hosea", "hos", "hosea", "hosia");
            Add("Joel");
            Add("Amos");
            Add("Obadiah", "obad", "obadia", "obadiya");
            Add("Jonah");
            Add("Micah", "mic", "mica");
            Add("Nahum", "nah");
            Add("Habakkuk", "hab", "habakuk", "habacuc", "habakuk", "habacuk");
            Add("Zephaniah", "zeph", "zefaniah", "zefania", "zephaniah");
            Add("Haggai", "hag", "hagai", "hagai");
            // Zechariah — most commonly mispronounced/misheard book in the OT.
            // Nigerian English: /zɛkəˈraɪə/ → often comes out as "zakariah",
            // "zakaria", "zekaria", "zacharia". The small Vosk model maps it to
            // "zech" because that's shorter. We add all phonetic variants so the
            // fuzzy matcher catches whatever Vosk produces.
            Add("Zechariah", "zech", "zachariah", "zacharia", "zakaria",
                "zakariah", "zekaria", "zekariah", "zacharia", "zecharia",
                "zecharias", "zacharias");
            Add("Malachi", "mal", "malaki", "malachi");

            // --- New Testament ---
            Add("Matthew", "matt", "mat", "mathew", "mathieu");
            Add("Mark");
            Add("Luke");
            Add("John", "jn");
            Add("Acts", "act");
            Add("Romans", "rom");
            Add("1 Corinthians", "1 cor", "i corinthians", "i cor", "1 corinthian", "corinthian");
            Add("2 Corinthians", "2 cor", "ii corinthians", "ii cor", "2 corinthian");
            Add("Galatians", "gal", "galatian");
            Add("Ephesians", "eph", "ephesian");
            Add("Philippians", "phil", "php", "philippian", "philipians", "philipian");
            Add("Colossians", "col", "colossian", "colosians", "colosian");
            Add("1 Thessalonians", "1 thess", "i thessalonians", "i thess", "1 thessalonian", "thessalonian");
            Add("2 Thessalonians", "2 thess", "ii thessalonians", "ii thess", "2 thessalonian");
            Add("1 Timothy", "1 tim", "i timothy", "i tim", "1 timoty", "timoty");
            Add("2 Timothy", "2 tim", "ii timothy", "ii tim", "2 timoty");
            Add("Titus", "tit");
            Add("Philemon", "phlm", "philem", "filemon", "philemon");
            Add("Hebrews", "heb", "hebrew");
            Add("James", "jas");
            Add("1 Peter", "1 pet", "i peter", "i pet");
            Add("2 Peter", "2 pet", "ii peter", "ii pet");
            Add("1 John", "1 jn", "i john", "i jn");
            Add("2 John", "2 jn", "ii john", "ii jn");
            Add("3 John", "3 jn", "iii john", "iii jn");
            Add("Jude");
            Add("Revelation", "rev", "revelations", "revelacion", "revelasion");

            // Sort by longest variant first so greedy matching works correctly
            return list.OrderByDescending(b => b.MaxTokens).ToList();
        }
    }
}
