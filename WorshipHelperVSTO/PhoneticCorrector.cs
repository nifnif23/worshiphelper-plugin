// ============================================================================
// PhoneticCorrector.cs  —  v3
// Pre-processes raw speech output to fix systematic phonetic mishearings
// before the text reaches BibleReferenceDetector.
//
// v3 changes:
//   • Open-dictation SCAN mode (the new Vosk setup) means the text now
//     contains plenty of natural English — "let's turn to John three
//     verse sixteen" rather than gibberish. The old phonetic corrector
//     was written for tight-grammar mode where EVERY out-of-vocabulary
//     sound was mapped to the nearest wordbank word. That meant many
//     token fixes (e.g. "ate" → "eight") used to fire unconditionally.
//     In open dictation that would corrupt normal English.
//
//   • Safe fixes (no plausible non-scripture meaning) stay in TokenFixes
//     and fire everywhere.
//
//   • Risky fixes ("ate" → "eight", "night" → "nine", etc.) moved to
//     ContextualTokenFixes and now ONLY fire when the token is surrounded
//     by unambiguous scripture context (a book name, structure word, or
//     another number token).
// ============================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace WorshipHelperVSTO
{
    public static class PhoneticCorrector
    {
        // -------------------------------------------------------------------
        // Pass 0: Book name normalisation
        // -------------------------------------------------------------------
        private static readonly List<(string From, string To)> BookNameFixes =
            new List<(string, string)>
        {
            // Zechariah — most commonly misheard OT book (Nigerian accent)
            ( "zachariah",   "zechariah" ),
            ( "zacharia",    "zechariah" ),
            ( "zakaria",     "zechariah" ),
            ( "zakariah",    "zechariah" ),
            ( "zekaria",     "zechariah" ),
            ( "zekariah",    "zechariah" ),
            ( "zecharia",    "zechariah" ),
            ( "zecharias",   "zechariah" ),
            ( "zacharias",   "zechariah" ),

            // Other commonly mispronounced OT books
            ( "nehemia",     "nehemiah"  ),
            ( "nehimiah",    "nehemiah"  ),
            ( "nehimia",     "nehemiah"  ),
            ( "jerimiah",    "jeremiah"  ),
            ( "jerimia",     "jeremiah"  ),
            ( "jeremia",     "jeremiah"  ),
            ( "obadia",      "obadiah"   ),
            ( "obadiya",     "obadiah"   ),
            ( "zefaniah",    "zephaniah" ),
            ( "zefania",     "zephaniah" ),
            ( "habakuk",     "habakkuk"  ),
            ( "habacuc",     "habakkuk"  ),
            ( "habacuk",     "habakkuk"  ),
            ( "hagai",       "haggai"    ),
            ( "malaki",      "malachi"   ),
            ( "esaiah",      "isaiah"    ),
            ( "isaia",       "isaiah"    ),
            ( "ezekel",      "ezekiel"   ),
            ( "ezekia",      "ezekiel"   ),
            ( "hosia",       "hosea"     ),
            ( "mica",        "micah"     ),
            ( "salms",       "psalms"    ),
            ( "sams",        "psalms"    ),

            // NT books
            ( "mathew",      "matthew"   ),
            ( "mathieu",     "matthew"   ),
            ( "corinthian",  "corinthians" ),
            ( "galatian",    "galatians"   ),
            ( "ephesian",    "ephesians"   ),
            ( "philipians",  "philippians" ),
            ( "philipian",   "philippians" ),
            ( "colosians",   "colossians"  ),
            ( "colosian",    "colossians"  ),
            ( "thessalonian","thessalonians" ),
            ( "timoty",      "timothy"   ),
            ( "filemon",     "philemon"  ),
            ( "hebrew",      "hebrews"   ),
            ( "revelacion",  "revelation"),
            ( "revelasion",  "revelation"),
        };

        // -------------------------------------------------------------------
        // Pass 2a: Safe token-level fixes.
        // These fire unconditionally because the mishearing has no plausible
        // non-scripture meaning in normal English dictation.
        // -------------------------------------------------------------------
        private static readonly Dictionary<string, string> TokenFixes =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // "won" sounds identical to "one" and rarely appears in scripture context
            { "won",        "one"     },

            // "too" / "tu" -> "two" (unambiguous)
            { "too",        "two"     },
            { "tu",         "two"     },

            // "fore" -> "four" (archaic outside of golf)
            { "fore",       "four"    },

            // Number spelling variants — safe
            { "sikh",       "six"     },
            { "sick",       "six"     },
            { "evan",       "seven"   },
            { "nigh",       "nine"    },
            { "nye",        "nine"    },
            { "knight",     "nine"    },

            // Spelling variants for compound numbers
            { "levin",      "eleven"  },
            { "elevin",     "eleven"  },
            { "twelfth",    "twelve"  },
            { "thirsting",  "thirteen"},
            { "thirsteen",  "thirteen"},
            { "forteen",    "fourteen"},
            { "fourteeen",  "fourteen"},
            { "fiveteen",   "fifteen" },
            { "seventheen", "seventeen"},
            { "eighteeen",  "eighteen"},
            { "nighteen",   "nineteen"},
            { "ninteen",    "nineteen"},
            { "twentie",    "twenty"  },
            { "tweny",      "twenty"  },
            { "thirtie",    "thirty"  },
            { "thirdy",     "thirty"  },
            { "fortie",     "forty"   },
            { "fourty",     "forty"   },
            { "fiftie",     "fifty"   },
            { "fifdy",      "fifty"   },
            { "sixtie",     "sixty"   },
            { "seventie",   "seventy" },
            { "eightie",    "eighty"  },
            { "ninetie",    "ninety"  },
            { "nightty",    "ninety"  },
            { "hunderd",    "hundred" },
            { "hundered",   "hundred" },
            { "hunderds",   "hundred" },

            // Structure-word misspellings
            { "versus",     "verse"   },
            { "chapture",   "chapter" },
            { "chaper",     "chapter" },

            // Colon separator
            { "cologne",    "colon"   },
            { "coln",       "colon"   },
        };

        // -------------------------------------------------------------------
        // Pass 2b: Contextual token fixes — only apply when the token is
        // surrounded by unambiguous scripture context (book name, structure
        // word, or another number). Prevents "they ate lunch" turning into
        // "they eight lunch" in open-dictation mode.
        // -------------------------------------------------------------------
        private static readonly Dictionary<string, string> ContextualTokenFixes =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "to",         "two"     },
            { "tea",        "three"   },
            { "tree",       "three"   },
            { "free",       "three"   },
            { "thee",       "three"   },
            { "hive",       "five"    },
            { "live",       "five"    },
            { "sex",        "six"     },
            { "heaven",     "seven"   },
            { "even",       "seven"   },
            { "heaventy",   "seventy" },
            { "ate",        "eight"   },
            { "hate",       "eight"   },
            { "night",      "nine"    },
            { "then",       "ten"     },
            { "den",        "ten"     },
            // "verse" mishearings (post-rationalisation tokens)
            { "burst",      "verse"   },
            { "firs",       "verse"   },
        };

        // -------------------------------------------------------------------
        // Pass 1: Multi-token phrase replacements
        // -------------------------------------------------------------------
        private static readonly List<(string From, string To)> PhraseFixes =
            new List<(string, string)>
        {
            // "fight for" -> "four" (Vosk artefact in tight-grammar mode;
            // kept for backwards compatibility even though v3 rarely sees it)
            ( "fight for",          "four"       ),
            ( "fight fore",         "four"       ),

            // "for tea" -> "forty"
            ( "for tea",            "forty"      ),
            ( "for ty",             "forty"      ),

            // "sex teen" -> "sixteen"
            ( "sex teen",           "sixteen"    ),
            ( "sick teen",          "sixteen"    ),
            ( "six teen",           "sixteen"    ),

            // "seven teen" -> "seventeen"
            ( "seven teen",         "seventeen"  ),
            ( "heaven teen",        "seventeen"  ),

            // "eight teen" -> "eighteen"
            ( "eight teen",         "eighteen"   ),
            ( "ate teen",           "eighteen"   ),

            // "nine teen" -> "nineteen"
            ( "nine teen",          "nineteen"   ),
            ( "night teen",         "nineteen"   ),

            // "twenty and X" -> "twenty X"
            ( "twenty and one",     "twenty one"    ),
            ( "twenty and two",     "twenty two"    ),
            ( "twenty and three",   "twenty three"  ),
            ( "twenty and four",    "twenty four"   ),
            ( "twenty and five",    "twenty five"   ),
            ( "twenty and six",     "twenty six"    ),
            ( "twenty and seven",   "twenty seven"  ),
            ( "twenty and eight",   "twenty eight"  ),
            ( "twenty and nine",    "twenty nine"   ),

            // "verse us" -> "verse" (versus -> verse misfire)
            ( "verse us",           "verse"      ),

            // "of us" / "of as" -> "verse" (old tight-grammar artefact,
            // rarely seen in v3 but harmless to leave in)
            ( "of us",              "verse"      ),
            ( "of as",              "verse"      ),
            ( "of is",              "verse"      ),
            ( "office",             "verse"      ),
            ( "of asses",           "verses"     ),
            ( "offices",            "verses"     ),

            // Common preamble variants — collapse to a single "turn to"
            ( "turn on to",         "turn to"    ),
            ( "turning to",         "turn to"    ),
            ( "turn over to",       "turn to"    ),
            ( "go to the book of",  "book of"    ),
        };

        // -------------------------------------------------------------------
        // Book-name / structure words used to detect scripture context
        // -------------------------------------------------------------------
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
            // structure words
            "chapter","chapters","verse","verses","colon","first","second","third",
        };

        // -------------------------------------------------------------------
        // Public API
        // -------------------------------------------------------------------
        public static string Correct(string rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText)) return rawText;

            string text = rawText.Trim().ToLowerInvariant();

            // Pass 0: book name normalisation
            text = ApplyBookNameFixes(text);

            // Pass 1: multi-token phrase replacements
            text = ApplyPhraseFixes(text);

            // Pass 2a: safe single-token fixes
            text = ApplyTokenFixes(text);

            // Pass 2b: context-sensitive token fixes
            text = ApplyContextualTokenFixes(text);

            // Pass 3: still-needed ambiguous fixes ("to"/"for"/"amos") with
            // smart numeric-context gating.
            text = FixAmbiguousTo(text);
            text = FixAmbiguousFor(text);
            text = FixAmosAsEight(text);

            if (!string.Equals(text, rawText, StringComparison.OrdinalIgnoreCase))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[PhoneticCorrector] \"{rawText}\" -> \"{text}\"");
            }

            return text;
        }

        // -------------------------------------------------------------------
        // Implementation helpers
        // -------------------------------------------------------------------
        private static string ApplyBookNameFixes(string text)
        {
            foreach (var (from, to) in BookNameFixes)
                text = Regex.Replace(text, @"\b" + Regex.Escape(from) + @"\b", to, RegexOptions.IgnoreCase);
            return text;
        }

        private static string ApplyPhraseFixes(string text)
        {
            foreach (var (from, to) in PhraseFixes.OrderByDescending(p => p.From.Length))
                text = Regex.Replace(text, @"\b" + Regex.Escape(from) + @"\b", to, RegexOptions.IgnoreCase);
            return text;
        }

        private static string ApplyTokenFixes(string text)
        {
            var tokens = text.Split(' ');
            for (int i = 0; i < tokens.Length; i++)
                if (TokenFixes.TryGetValue(tokens[i], out string replacement))
                    tokens[i] = replacement;
            return string.Join(" ", tokens);
        }

        private static string ApplyContextualTokenFixes(string text)
        {
            var tokens = text.Split(' ');
            for (int i = 0; i < tokens.Length; i++)
            {
                if (!ContextualTokenFixes.TryGetValue(tokens[i], out string replacement))
                    continue;

                bool leftCtx  = i > 0 && IsScriptureContextToken(tokens[i - 1]);
                bool rightCtx = i < tokens.Length - 1 && IsScriptureContextToken(tokens[i + 1]);

                if (leftCtx || rightCtx)
                    tokens[i] = replacement;
            }
            return string.Join(" ", tokens);
        }

        private static bool IsScriptureContextToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return false;
            token = token.Trim().ToLowerInvariant();
            if (int.TryParse(token, out _)) return true;
            if (SpokenNumberConverter.IsNumberWord(token)) return true;
            return _structureOrBookWords.Contains(token);
        }

        private static string FixAmbiguousTo(string text)
        {
            var tokens = text.Split(' ').ToList();
            for (int i = 1; i < tokens.Count - 1; i++)
            {
                if (!tokens[i].Equals("to", StringComparison.OrdinalIgnoreCase)) continue;

                bool prevIsNum = IsNumberToken(tokens[i - 1]);
                bool nextIsNum = IsNumberToken(tokens[i + 1]);

                if (prevIsNum && nextIsNum) continue;                       // range "X to Y"
                if (!prevIsNum && nextIsNum) tokens[i] = "two";             // "to X" -> "two X"
            }
            return string.Join(" ", tokens);
        }

        private static string FixAmbiguousFor(string text)
        {
            var tokens = text.Split(' ').ToList();
            for (int i = 0; i < tokens.Count; i++)
            {
                if (!tokens[i].Equals("for", StringComparison.OrdinalIgnoreCase)) continue;

                bool prevIsNum = i > 0 && IsNumberToken(tokens[i - 1]);
                bool nextIsNum = i < tokens.Count - 1 && IsNumberToken(tokens[i + 1]);

                if (prevIsNum || (nextIsNum && i > 0 && _structureOrBookWords.Contains(tokens[i - 1])))
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

        /// <summary>
        /// Fixes "amos" appearing in a non-leading position as a mishearing of "eight".
        /// </summary>
        private static string FixAmosAsEight(string text)
        {
            var tokens = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length <= 1) return text;

            bool changed = false;
            for (int i = 1; i < tokens.Length; i++)
            {
                if (string.Equals(tokens[i], "amos", StringComparison.OrdinalIgnoreCase))
                {
                    // Only swap if surrounded by scripture context — prevents
                    // "reading amos chapter five" (book of Amos) getting wrecked.
                    // But if this is position >= 1 and the previous token is a
                    // book or structure word or number, this is almost certainly
                    // the "eight" mishearing.
                    bool prevCtx = IsScriptureContextToken(tokens[i - 1]);
                    bool nextCtx = i < tokens.Length - 1 && IsScriptureContextToken(tokens[i + 1]);

                    // Avoid rewriting if "amos" is the book itself, which would
                    // have been the FIRST content token. After rationalisation
                    // preamble words are stripped, so position 0 = likely the
                    // book. Still, require context to fire.
                    if (prevCtx || nextCtx)
                    {
                        tokens[i] = "eight";
                        changed = true;
                    }
                }
            }
            return changed ? string.Join(" ", tokens) : text;
        }
    }
}
