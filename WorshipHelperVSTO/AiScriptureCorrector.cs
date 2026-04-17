// ============================================================================
// AiScriptureCorrector.cs
// Uses an AI model (via HTTP) to recover valid scripture references from
// garbled speech-to-text output that the rule-based detector missed or
// mis-parsed.
//
// Problem it solves:
//   Vosk outputs something like "zechariah fight for tea night" when the
//   speaker said "Zechariah 4:9" — the grammar constrains the vocabulary
//   but chapter/verse tokens still get mangled if the phonemes are odd.
//
//   The rule-based BibleReferenceDetector handles the clean case well.
//   This class is the fallback for anything that looked like a scripture
//   reference (book name matched) but the number fragment was garbled or
//   missing — OR where confidence came out low.
//
// Strategy:
//   1. BibleReferenceDetector runs first (fast, offline, no cost).
//   2. If no result, OR confidence < AiConfidenceThreshold, this class
//      sends a compact prompt to the AI model to recover the reference.
//   3. AI output is validated against the known book list before use.
//   4. Results are cached so repeated queries are free.
//
// Drop into:  WorshipHelperVSTO/AiScriptureCorrector.cs
// Namespace:  WorshipHelperVSTO
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using log4net;

namespace WorshipHelperVSTO
{
    /// <summary>
    /// Attempts to recover a valid Bible reference from garbled speech output
    /// using an AI model, falling back gracefully if unavailable.
    /// </summary>
    public class AiScriptureCorrector
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(AiScriptureCorrector));

        // -----------------------------------------------------------------------
        // Configuration
        // -----------------------------------------------------------------------

        /// <summary>
        /// Path to a plain-text file containing your API key (one line).
        /// Stored in %APPDATA%\WorshipHelper\ai-key.txt.
        /// If the file does not exist, AI correction is silently disabled.
        /// </summary>
        public static string KeyFilePath { get; set; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "WorshipHelper", "ai-key.txt");

        /// <summary>
        /// AI API endpoint.  Defaults to Anthropic's Messages API.
        /// Override to point at a local proxy or different provider.
        /// </summary>
        public static string ApiEndpoint { get; set; } =
            "https://api.anthropic.com/v1/messages";

        /// <summary>
        /// Model to use.  claude-haiku is fast and cheap enough for real-time use.
        /// </summary>
        public static string Model { get; set; } = "claude-haiku-4-5-20251001";

        /// <summary>
        /// Minimum BibleReferenceDetector confidence below which we ask the AI
        /// to verify / correct the reference.
        /// At or above this threshold, the rule-based result is used as-is.
        /// Default: 0.65 — ask AI whenever the rule-based result is uncertain.
        /// </summary>
        public static double AiConfidenceThreshold { get; set; } = 0.65;

        /// <summary>
        /// HTTP timeout in milliseconds.  Keep this low — this is live presentation
        /// mode and we'd rather skip a correction than freeze the UI.
        /// Default: 2500 ms.
        /// </summary>
        public static int TimeoutMs { get; set; } = 2500;

        // -----------------------------------------------------------------------
        // Result cache — same garbled input almost always produces the same output
        // -----------------------------------------------------------------------
        private static readonly Dictionary<string, string> _cache =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private static string _apiKey; // null means disabled
        private static bool _keyLoaded;

        // -----------------------------------------------------------------------
        // All 66 canonical book names — used to validate AI output.
        // Must match the canonical names in BibleReferenceDetector exactly.
        // -----------------------------------------------------------------------
        private static readonly HashSet<string> CanonicalBooks =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Genesis","Exodus","Leviticus","Numbers","Deuteronomy","Joshua","Judges","Ruth",
            "1 Samuel","2 Samuel","1 Kings","2 Kings","1 Chronicles","2 Chronicles",
            "Ezra","Nehemiah","Esther","Job","Psalms","Proverbs","Ecclesiastes",
            "Song of Solomon","Isaiah","Jeremiah","Lamentations","Ezekiel","Daniel",
            "Hosea","Joel","Amos","Obadiah","Jonah","Micah","Nahum","Habakkuk",
            "Zephaniah","Haggai","Zechariah","Malachi",
            "Matthew","Mark","Luke","John","Acts","Romans",
            "1 Corinthians","2 Corinthians","Galatians","Ephesians","Philippians",
            "Colossians","1 Thessalonians","2 Thessalonians","1 Timothy","2 Timothy",
            "Titus","Philemon","Hebrews","James","1 Peter","2 Peter",
            "1 John","2 John","3 John","Jude","Revelation",
        };

        // -----------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------

        /// <summary>
        /// Main entry point.  Given the raw speech text and the result from
        /// BibleReferenceDetector (which may be null), returns the best available
        /// normalised reference string, or null if nothing can be recovered.
        ///
        /// Decision logic:
        ///   - If detected != null AND confidence >= threshold  →  use rule-based result (no AI call)
        ///   - If detected != null AND confidence < threshold   →  ask AI to verify/correct
        ///   - If detected == null                              →  ask AI to recover from scratch
        ///
        /// The AI call is best-effort: if it fails (network, key missing, timeout),
        /// the rule-based result (or null) is returned silently.
        /// </summary>
        public static string Correct(string rawSpeech, DetectedReference ruleBasedResult)
        {
            if (string.IsNullOrWhiteSpace(rawSpeech))
                return ruleBasedResult?.NormalisedReference;

            // Fast path: rule-based was confident enough
            if (ruleBasedResult != null && ruleBasedResult.Confidence >= AiConfidenceThreshold)
            {
                log.Debug($"AiCorrector: rule-based confident ({ruleBasedResult.Confidence:F2}), skipping AI.");
                return ruleBasedResult.NormalisedReference;
            }

            // Load API key once
            if (!_keyLoaded) LoadApiKey();

            if (_apiKey == null)
            {
                // AI unavailable — fall back to rule-based result (may be null)
                log.Debug("AiCorrector: no API key, skipping AI correction.");
                return ruleBasedResult?.NormalisedReference;
            }

            // Cache lookup
            string cacheKey = rawSpeech.Trim().ToLowerInvariant();
            if (_cache.TryGetValue(cacheKey, out string cached))
            {
                log.Debug($"AiCorrector: cache hit for \"{rawSpeech}\" → \"{cached}\"");
                return cached;
            }

            // Build context hint from the rule-based guess (if any)
            string hint = ruleBasedResult != null
                ? $"The rule-based parser's best guess was \"{ruleBasedResult.NormalisedReference}\" " +
                  $"(confidence {ruleBasedResult.Confidence:F2}). Correct it if wrong, or confirm it if right."
                : "The rule-based parser found no match.";

            string aiResult = CallAi(rawSpeech, hint);

            // Validate and cache
            if (aiResult != null)
            {
                log.Info($"AiCorrector: \"{rawSpeech}\" → \"{aiResult}\"");
                _cache[cacheKey] = aiResult;
            }
            else
            {
                log.Debug($"AiCorrector: AI returned no valid reference for \"{rawSpeech}\".");
                // Cache the null so we don't re-query the same bad input
                _cache[cacheKey] = null;
            }

            return aiResult ?? ruleBasedResult?.NormalisedReference;
        }

        // -----------------------------------------------------------------------
        // Private helpers
        // -----------------------------------------------------------------------

        private static void LoadApiKey()
        {
            _keyLoaded = true;
            try
            {
                if (File.Exists(KeyFilePath))
                {
                    string key = File.ReadAllText(KeyFilePath).Trim();
                    if (!string.IsNullOrEmpty(key))
                    {
                        _apiKey = key;
                        log.Info("AiCorrector: API key loaded.");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                log.Warn($"AiCorrector: could not read key file: {ex.Message}");
            }
            log.Info("AiCorrector: no API key found — AI correction disabled.");
        }

        /// <summary>
        /// Sends a compact prompt to the AI model and returns a validated
        /// normalised reference string, or null.
        /// </summary>
        private static string CallAi(string rawSpeech, string hint)
        {
            try
            {
                // Compact, token-efficient prompt.
                // We tell the model exactly the vocabulary it's working with and
                // ask for a terse, machine-parseable response.
                string prompt =
                    "You are correcting garbled speech-to-text output from a worship presentation tool. " +
                    "The user spoke a Bible scripture reference into a microphone. " +
                    "The speech engine may have mangled book names or chapter/verse numbers. " +
                    "\n\n" +
                    "Raw speech output: \"" + rawSpeech + "\"\n" +
                    hint + "\n\n" +
                    "Valid canonical book names (use EXACTLY these spellings):\n" +
                    string.Join(", ", CanonicalBooks.OrderBy(b => b)) + "\n\n" +
                    "Rules:\n" +
                    "- Reply with ONLY the normalised reference, e.g. \"Zechariah 4:9\" or \"John 3:16-18\".\n" +
                    "- If the speech contains a chapter+verse reference, always include both: \"Book chapter:verse\".\n" +
                    "- If only a chapter is identifiable, reply \"Book chapter\".\n" +
                    "- If you cannot identify any plausible scripture reference, reply with the single word: NONE\n" +
                    "- No explanation, no punctuation other than the colon and optional hyphen in the reference.\n" +
                    "- Do not guess wildly — only reply if you are reasonably confident.";

                string requestBody =
                    "{\"model\":\"" + Model + "\"," +
                    "\"max_tokens\":50," +
                    "\"messages\":[{\"role\":\"user\",\"content\":" + JsonString(prompt) + "}]}";

                var request = (HttpWebRequest)WebRequest.Create(ApiEndpoint);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.Headers.Add("x-api-key", _apiKey);
                request.Headers.Add("anthropic-version", "2023-06-01");
                request.Timeout = TimeoutMs;

                byte[] body = Encoding.UTF8.GetBytes(requestBody);
                request.ContentLength = body.Length;

                using (var stream = request.GetRequestStream())
                    stream.Write(body, 0, body.Length);

                using (var response = (HttpWebResponse)request.GetResponse())
                using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    string json = reader.ReadToEnd();
                    return ParseAiResponse(json);
                }
            }
            catch (WebException wex) when (wex.Status == WebExceptionStatus.Timeout)
            {
                log.Warn("AiCorrector: API call timed out.");
                return null;
            }
            catch (Exception ex)
            {
                log.Warn($"AiCorrector: API call failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Extracts and validates the text content from a raw Anthropic API JSON response.
        /// Returns null if the response is "NONE" or doesn't look like a valid reference.
        /// </summary>
        private static string ParseAiResponse(string json)
        {
            // Extract the text content from the response JSON.
            // Anthropic response: {"content":[{"type":"text","text":"Zechariah 4:9"}],...}
            var match = Regex.Match(json, "\"text\"\\s*:\\s*\"([^\"]+)\"");
            if (!match.Success) return null;

            string text = match.Groups[1].Value.Trim();

            // Unescape JSON string escapes
            text = text.Replace("\\n", " ").Replace("\\\"", "\"").Replace("\\\\", "\\");

            if (string.IsNullOrWhiteSpace(text) ||
                text.Equals("NONE", StringComparison.OrdinalIgnoreCase))
                return null;

            return ValidateReference(text);
        }

        /// <summary>
        /// Validates that the AI's output looks like a real Bible reference.
        /// Returns the normalised string if valid, null otherwise.
        ///
        /// Format: "Book chapter" or "Book chapter:verse" or "Book chapter:verse-end"
        /// Multi-word books: "1 Corinthians 13:4"
        /// </summary>
        private static string ValidateReference(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;

            // Strip trailing punctuation that a verbose model might add
            raw = raw.TrimEnd('.', ',', ';', '!', '?');

            // Try to split into "book" + "numbers" portions.
            // The numbers part starts at the first digit token.
            // e.g. "Zechariah 4:9" → book="Zechariah", numbers="4:9"
            // e.g. "1 Corinthians 13:4" → book="1 Corinthians", numbers="13:4"

            var tokens = raw.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length < 2) return null;

            string bookName = null;
            string refPart = null;

            // Find the longest prefix that matches a canonical book name
            for (int len = Math.Min(tokens.Length - 1, 4); len >= 1; len--)
            {
                string candidate = string.Join(" ", tokens.Take(len));
                if (CanonicalBooks.Contains(candidate))
                {
                    bookName = candidate;
                    refPart = string.Join(" ", tokens.Skip(len)).Trim();
                    break;
                }
            }

            if (bookName == null || string.IsNullOrWhiteSpace(refPart)) return null;

            // Validate the reference fragment: digits, colon, optional hyphen
            if (!Regex.IsMatch(refPart, @"^\d+([:\-]\d+)?$")) return null;

            return $"{bookName} {refPart}";
        }

        /// <summary>
        /// Returns the string JSON-encoded (quoted, with internal characters escaped).
        /// Minimal implementation — only handles characters that appear in our prompts.
        /// </summary>
        private static string JsonString(string s)
        {
            return "\"" +
                s.Replace("\\", "\\\\")
                 .Replace("\"", "\\\"")
                 .Replace("\n", "\\n")
                 .Replace("\r", "") +
                "\"";
        }
    }
}
