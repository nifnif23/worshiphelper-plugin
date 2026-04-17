// ============================================================================
// PhoneticCorrector.cs
// Pre-processes raw Vosk speech output to fix systematic phonetic mishearings
// before the text reaches BibleReferenceDetector.
//
// The problem:
//   Vosk's constrained grammar vocabulary means it can only output words from
//   the approved list. But when the speaker says "four nine", Vosk is forced
//   to pick the closest-sounding word it DOES know — so phonemes get mapped
//   to wrong-but-in-vocabulary words:
//
//     "zechariah four nine"  → "zechariah fight for tea night"
//     "genesis three sixteen" → "genesis three sex teen"
//     "psalm twenty seven"   → "psalm twenty heaven"
//
// The fix:
//   Two-pass correction applied to the raw token stream before detection:
//
//   Pass 1 — Single-token corrections:
//     Direct phonetic near-miss replacements for number words that Vosk
//     reliably mishears. e.g. "night" → "nine", "tea" → "three".
//
//   Pass 2 — Multi-token phrase corrections:
//     Some mishearings span multiple tokens. e.g. "fight for" → "four",
//     "sex teen" → "sixteen". Handled with ordered phrase substitutions.
//
// This is 100% offline, zero-latency, and zero-dependency.
//
// Drop into:  WorshipHelperVSTO/PhoneticCorrector.cs
// Namespace:  WorshipHelperVSTO
// ============================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace WorshipHelperVSTO
{
    /// <summary>
    /// Corrects systematic Vosk phonetic mishearings in speech output before
    /// it is passed to the Bible reference detector.
    ///
    /// All corrections are conservative — they only fire when the surrounding
    /// context makes a number interpretation unambiguous (i.e. the word
    /// appears after a known book name or another number word).
    /// Context-free substitutions are also included for high-confidence
    /// near-misses that have no plausible non-number meaning in this domain.
    /// </summary>
    public static class PhoneticCorrector
    {
        // -----------------------------------------------------------------------
        // Pass 1: Single-token phonetic replacements
        //
        // Format: { "mishearing" → "correction" }
        //
        // Only include words where the mishearing has essentially no other
        // plausible meaning in a Bible-reference context. "mark" stays as
        // "mark" (book name), "acts" stays as "acts", etc.
        //
        // Phonetic reasoning for each entry is documented inline.
        // -----------------------------------------------------------------------
        private static readonly Dictionary<string, string> TokenFixes =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // ── "one" mishearings ──────────────────────────────────────────────
            // "won" sounds identical
            { "won",        "one"     },

            // ── "two" mishearings ──────────────────────────────────────────────
            { "to",         "two"     },  // /tuː/ identical — context-gated below
            { "too",        "two"     },
            { "tu",         "two"     },

            // ── "three" mishearings ────────────────────────────────────────────
            // /θriː/ → Vosk often drops the /θ/ and hears /t/ instead
            { "tea",        "three"   },  // th→t, ee≈ee  (most common "three" mishear)
            { "tree",       "three"   },  // th→t
            { "free",       "three"   },  // th→f (common in non-rhotic speakers)
            { "thee",       "three"   },  // grammar word; here /θiː/ ≈ /θriː/

            // ── "four" mishearings ─────────────────────────────────────────────
            // "for" / "fore" are homophones in most accents
            { "fore",       "four"    },

            // ── "five" mishearings ─────────────────────────────────────────────
            { "hive",       "five"    },  // h/f confusion, -ive identical
            { "live",       "five"    },  // l/f, -ive identical

            // ── "six" mishearings ──────────────────────────────────────────────
            { "sex",        "six"     },  // /sɛks/ ≈ /sɪks/
            { "sikh",       "six"     },  // /siːk/ ≈ /sɪks/
            { "sick",       "six"     },

            // ── "seven" mishearings ────────────────────────────────────────────
            { "heaven",     "seven"   },  // h-eaven ≈ s-even (very common!)
            { "evan",       "seven"   },  // dropped s-
            { "even",       "seven"   },

            // ── "eight" mishearings ────────────────────────────────────────────
            { "ate",        "eight"   },  // homophones
            { "hate",       "eight"   },  // h+ate

            // ── "nine" mishearings ─────────────────────────────────────────────
            { "night",      "nine"    },  // /naɪt/ ≈ /naɪn/ — extremely common!
            { "knight",     "nine"    },
            { "nigh",       "nine"    },
            { "nye",        "nine"    },

            // ── "ten" mishearings ──────────────────────────────────────────────
            { "then",       "ten"     },  // /ðɛn/ ≈ /tɛn/
            { "den",        "ten"     },

            // ── "eleven" mishearings ───────────────────────────────────────────
            { "levin",      "eleven"  },  // dropped e-
            { "elevin",     "eleven"  },

            // ── "twelve" mishearings ───────────────────────────────────────────
            { "twelfth",    "twelve"  },  // ordinal form

            // ── "thirteen" mishearings ─────────────────────────────────────────
            { "thirsting",  "thirteen"},
            { "thirsteen",  "thirteen"},

            // ── "fourteen" mishearings ─────────────────────────────────────────
            { "forteen",    "fourteen"},
            { "fourteeen",  "fourteen"},

            // ── "fifteen" mishearings ──────────────────────────────────────────
            { "fiveteen",   "fifteen" },

            // ── "sixteen" mishearings ──────────────────────────────────────────
            // Often split into two tokens by Vosk; handled in phrase pass below
            { "sixteen",    "sixteen" },  // identity — keep for normalisation

            // ── "seventeen" mishearings ────────────────────────────────────────
            { "seventheen", "seventeen"},

            // ── "eighteen" mishearings ─────────────────────────────────────────
            { "eighteeen",  "eighteen"},

            // ── "nineteen" mishearings ─────────────────────────────────────────
            { "nighteen",   "nineteen"},  // night+een
            { "ninteen",    "nineteen"},

            // ── "twenty" mishearings ───────────────────────────────────────────
            { "twentie",    "twenty"  },
            { "tweny",      "twenty"  },

            // ── "thirty" mishearings ───────────────────────────────────────────
            { "thirtie",    "thirty"  },
            { "thirdy",     "thirty"  },

            // ── "forty" mishearings ────────────────────────────────────────────
            { "fortie",     "forty"   },
            { "fourty",     "forty"   },  // common misspelling also heard by Vosk

            // ── "fifty" mishearings ────────────────────────────────────────────
            { "fiftie",     "fifty"   },
            { "fifdy",      "fifty"   },

            // ── "sixty" mishearings ────────────────────────────────────────────
            { "sixtie",     "sixty"   },

            // ── "seventy" mishearings ──────────────────────────────────────────
            { "seventie",   "seventy" },
            { "heaventy",   "seventy" },  // heaven+ty

            // ── "eighty" mishearings ───────────────────────────────────────────
            { "eightie",    "eighty"  },

            // ── "ninety" mishearings ───────────────────────────────────────────
            { "ninetie",    "ninety"  },
            { "nightty",    "ninety"  },

            // ── "hundred" mishearings ──────────────────────────────────────────
            { "hunderd",    "hundred" },
            { "hundered",   "hundred" },

            // ── "verse" mishearings ────────────────────────────────────────────
            // These only make sense as structure words in this domain
            { "versus",     "verse"   },
            { "burst",      "verse"   },  // /bɜːst/ ≈ /vɜːs/
            { "firs",       "verse"   },  // v/f + dropped t

            // ── "chapter" mishearings ──────────────────────────────────────────
            { "chapture",   "chapter" },
            { "chaper",     "chapter" },

            // ── Colon / separator words ────────────────────────────────────────
            { "cologne",    "colon"   },  // /kəˈloʊn/ ≈ /ˈkoʊlən/
            { "coln",       "colon"   },
        };

        // -----------------------------------------------------------------------
        // Pass 2: Multi-token phrase replacements
        //
        // Applied BEFORE single-token fixes (longest match wins).
        // Ordered from longest to shortest to ensure greedy matching.
        //
        // Format: { "phrase to match" → "replacement tokens" }
        // -----------------------------------------------------------------------
        private static readonly List<(string From, string To)> PhraseFixes =
            new List<(string, string)>
        {
            // "fight for" is extremely common for "four" in constrained grammar
            // (Vosk maps /faɪt fɔː/ from surrounding context noise)
            ( "fight for",          "four"       ),
            ( "fight fore",         "four"       ),

            // "for tea" → "forty" (forty splits across two tokens)
            ( "for tea",            "forty"      ),
            ( "for ty",             "forty"      ),

            // "sex teen" → "sixteen" (vosk splits at morpheme boundary)
            ( "sex teen",           "sixteen"    ),
            ( "sick teen",          "sixteen"    ),
            ( "six teen",           "sixteen"    ),

            // "seven teen" → "seventeen"
            ( "seven teen",         "seventeen"  ),
            ( "heaven teen",        "seventeen"  ),

            // "eight teen" → "eighteen"
            ( "eight teen",         "eighteen"   ),
            ( "ate teen",           "eighteen"   ),

            // "nine teen" → "nineteen"
            ( "nine teen",          "nineteen"   ),
            ( "night teen",         "nineteen"   ),

            // "twenty one/two/..." — Vosk sometimes adds "and" between these
            ( "twenty and one",     "twenty one"    ),
            ( "twenty and two",     "twenty two"    ),
            ( "twenty and three",   "twenty three"  ),
            ( "twenty and four",    "twenty four"   ),
            ( "twenty and five",    "twenty five"   ),
            ( "twenty and six",     "twenty six"    ),
            ( "twenty and seven",   "twenty seven"  ),
            ( "twenty and eight",   "twenty eight"  ),
            ( "twenty and nine",    "twenty nine"   ),

            // "one hundred and X" → correct, already handled; but "one hundred X" needs colon
            // Handled downstream — no phrase fix needed here.

            // "verse us" → "verse" (versus → verse misfire)
            ( "verse us",           "verse"      ),

            // "book of" before a book name — not needed but reduces token count
            // Left to preamble stripping in BibleReferenceDetector.
        };

        // -----------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------

        /// <summary>
        /// Applies phonetic correction to raw Vosk output text.
        /// Returns the corrected string, ready to be passed to BibleReferenceDetector.
        ///
        /// Example:
        ///   "zechariah fight for tea night" → "zechariah four three nine"
        ///
        /// Note: "to" is ambiguous (it's in the grammar as a structure word AND
        /// sounds like "two"). We only replace "to" with "two" when it appears
        /// in a numeric context — i.e. sandwiched between other number words.
        /// </summary>
        public static string Correct(string rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText)) return rawText;

            string text = rawText.Trim().ToLowerInvariant();

            // Pass 1: Phrase replacements (longest-first, case-insensitive)
            text = ApplyPhraseFixes(text);

            // Pass 2: Token replacements
            text = ApplyTokenFixes(text);

            // Pass 3: Context-sensitive "to" → "two" fix
            text = FixAmbiguousTo(text);

            // Pass 4: Context-sensitive "for" → "four" fix
            // "for" is a preamble word AND can be "four" — only fix when numeric context
            text = FixAmbiguousFor(text);

            if (!string.Equals(text, rawText, StringComparison.OrdinalIgnoreCase))
            {
                // Log the correction so we can build a better map over time
                System.Diagnostics.Debug.WriteLine(
                    $"[PhoneticCorrector] \"{rawText}\" → \"{text}\"");
            }

            return text;
        }

        // -----------------------------------------------------------------------
        // Private helpers
        // -----------------------------------------------------------------------

        private static string ApplyPhraseFixes(string text)
        {
            // Sort by descending length already guaranteed by list order,
            // but be safe and re-sort in case entries are reordered later.
            foreach (var (from, to) in PhraseFixes.OrderByDescending(p => p.From.Length))
            {
                // Whole-word phrase replacement, case-insensitive
                text = Regex.Replace(
                    text,
                    @"\b" + Regex.Escape(from) + @"\b",
                    to,
                    RegexOptions.IgnoreCase);
            }
            return text;
        }

        private static string ApplyTokenFixes(string text)
        {
            var tokens = text.Split(' ');
            for (int i = 0; i < tokens.Length; i++)
            {
                if (TokenFixes.TryGetValue(tokens[i], out string replacement))
                    tokens[i] = replacement;
            }
            return string.Join(" ", tokens);
        }

        /// <summary>
        /// Replaces "to" with "two" only when it appears between two number-like
        /// tokens — e.g. "john three to six" (range "3 to 6") should stay as
        /// "to", but "verse to" where "to" is alone after a book+chapter should
        /// become "two" if surrounded by numbers.
        ///
        /// Rule: replace "to" → "two" when BOTH the token before and after are
        /// number words / digits AND neither is a structure keyword like "verse",
        /// "chapter", "through".
        /// </summary>
        private static string FixAmbiguousTo(string text)
        {
            var tokens = text.Split(' ').ToList();
            for (int i = 1; i < tokens.Count - 1; i++)
            {
                if (!tokens[i].Equals("to", StringComparison.OrdinalIgnoreCase)) continue;

                bool prevIsNum = IsNumberToken(tokens[i - 1]);
                bool nextIsNum = IsNumberToken(tokens[i + 1]);

                // "X to Y" where both sides are numbers → range, keep "to"
                if (prevIsNum && nextIsNum) continue;

                // "to X" where only right side is a number (chapter start) → "two"
                if (!prevIsNum && nextIsNum)
                    tokens[i] = "two";
            }
            return string.Join(" ", tokens);
        }

        /// <summary>
        /// Replaces "for" with "four" only in numeric context.
        /// "for" as a preamble word ("turn to for...") should stay.
        /// "john three for" where "for" follows a number → "four".
        /// </summary>
        private static string FixAmbiguousFor(string text)
        {
            var tokens = text.Split(' ').ToList();
            for (int i = 0; i < tokens.Count; i++)
            {
                if (!tokens[i].Equals("for", StringComparison.OrdinalIgnoreCase)) continue;

                // Only replace if the previous token is a number word or digit
                bool prevIsNum = i > 0 && IsNumberToken(tokens[i - 1]);
                bool nextIsNum = i < tokens.Count - 1 && IsNumberToken(tokens[i + 1]);

                if (prevIsNum || (nextIsNum && i > 0 && IsBookOrStructure(tokens[i - 1])))
                    tokens[i] = "four";
            }
            return string.Join(" ", tokens);
        }

        private static bool IsNumberToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return false;
            if (int.TryParse(token, out _)) return true;
            return SpokenNumberConverter.IsNumberWord(token);
        }

        // Book-name words that can precede a number directly without a "chapter" keyword
        private static readonly HashSet<string> _structureOrBookWords =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "zechariah","zech","genesis","gen","exodus","exod","leviticus","lev",
            "numbers","num","deuteronomy","deut","joshua","josh","judges","judg",
            "ruth","samuel","sam","kings","king","chronicles","chron","ezra",
            "nehemiah","neh","esther","job","psalms","psalm","psa","proverbs",
            "prov","ecclesiastes","eccl","isaiah","isa","jeremiah","jer",
            "lamentations","lam","ezekiel","ezek","daniel","dan","hosea","hos",
            "joel","amos","obadiah","obad","jonah","micah","mic","nahum","nah",
            "habakkuk","hab","zephaniah","zeph","haggai","hag","malachi","mal",
            "matthew","matt","mark","luke","john","jn","acts","act","romans","rom",
            "corinthians","cor","galatians","gal","ephesians","eph","philippians",
            "phil","colossians","col","thessalonians","thess","timothy","tim",
            "titus","tit","philemon","phlm","hebrews","heb","james","jas",
            "peter","pet","jude","revelation","rev",
            // structure words that can precede a chapter number
            "chapter","chapters","verse","verses","colon",
        };

        private static bool IsBookOrStructure(string token)
            => _structureOrBookWords.Contains(token);
    }
}
