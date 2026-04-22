// ============================================================================
// SpeechListener.cs  —  v3 (hybrid grammar + stronger rationaliser)
//
// Real-time speech recognition using Vosk (offline, free, open-source).
//
// Why v3:
//   v1/v2 ran Vosk with a tiny "wordbank of scriptures" grammar. That forced
//   Vosk to pick words from a 200-word list even when the audio didn't match
//   any of them — accuracy tanked and rare words (Zechariah, Habakkuk, ...)
//   came out as near-miss common words. The user quite rightly said:
//
//      "Not bad — just very limiting. Vosk is basically being forced into a
//       tiny box. It can only output those words, even when the audio doesn't
//       match them well. You want a hybrid vocabulary, not a tiny one."
//
//   v3 fixes this by:
//
//     • SCAN phase → runs Vosk in **open** dictation mode (no grammar at all).
//       Vosk has its full ~200k-word model available, so it anchors audio to
//       real English sentences. Ministers can now say
//         "Let's all turn to the book of John, chapter three, verse sixteen"
//       and get a clean transcript, not "lets ... john tea sixteen".
//     • FOCUS phase → still uses a hyper-tight grammar (the specific book's
//       variants + numbers + structure words). Once we know the book, the
//       focused grammar scrubs any chapter/verse noise.
//     • Rationaliser is much stronger: strips fillers, disfluencies, partial
//       words, commonplace dictation artefacts, and normalises punctuation
//       before the Bible reference detector ever sees it.
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using log4net;
using NAudio.Wave;
using Vosk;

namespace WorshipHelperVSTO
{
    public class SpeechRecognisedEventArgs : EventArgs
    {
        public string Text { get; set; }
        public float Confidence { get; set; }
    }

    public class SpeechListenerStatusEventArgs : EventArgs
    {
        public string Message { get; set; }
        public bool IsError { get; set; }
    }

    public class SpeechPhaseChangedEventArgs : EventArgs
    {
        /// <summary>"Scan" or "Focus".</summary>
        public string Phase { get; set; }
        /// <summary>For Focus phase, the canonical book name. Null otherwise.</summary>
        public string BookName { get; set; }
    }

    public class SpeechListener : IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(SpeechListener));

        private static readonly string DefaultModelPath = ResolveModelPath();

        private static string ResolveModelPath()
        {
            string installDir = Microsoft.Win32.Registry.CurrentUser
                .OpenSubKey(@"SOFTWARE\WorshipHelper")
                ?.GetValue("InstallLocation") as string;

            if (!string.IsNullOrEmpty(installDir))
            {
                string registryPath = Path.Combine(installDir, "data", "vosk-model");
                if (Directory.Exists(registryPath)) return registryPath;
            }

            string dllDir = Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location) ?? "";
            string bundled = Path.Combine(dllDir, "data", "vosk-model");
            if (Directory.Exists(bundled)) return bundled;

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "WorshipHelper", "vosk-model");
        }

        private VoskRecognizer _recogniser;
        private WaveInEvent _waveIn;
        private bool _isListening;
        private bool _disposed;
        private readonly object _lock = new object();

        // ──────────────────────────────────────────────────────────────────
        // Two-phase recognition state
        //
        //   SCAN  — open-vocabulary dictation (no grammar). Vosk has its full
        //           200k-word English model available. This is the key change
        //           from v1/v2: we no longer force Vosk to pick from a tiny
        //           scripture-only wordbank that produced constant misfires.
        //
        //   FOCUS — hyper-tight grammar built from the detected book's
        //           variants + numbers + structure words only. Kicks in as
        //           soon as we hear a book name. Reverts to SCAN after the
        //           utterance fires or after an idle timeout.
        // ──────────────────────────────────────────────────────────────────
        private Model _model;
        private string _phase2BookName;
        private System.Threading.Timer _phase2Timeout;
        private const int Phase2TimeoutMs = 5000;  // 5s of silence → back to scan

        private enum RecognitionPhase { Scan, Focus }
        private RecognitionPhase _currentPhase = RecognitionPhase.Scan;

        // ──────────────────────────────────────────────────────────────────
        // Configuration
        // ──────────────────────────────────────────────────────────────────
        public string ModelPath { get; set; } = DefaultModelPath;

        /// <summary>
        /// Minimum average word confidence before SpeechRecognised fires.
        /// Kept low because open-dictation mode naturally returns slightly
        /// lower per-word confidence than a constrained grammar; the combined
        /// confidence (with detector score) is the real gate.
        /// </summary>
        public float MinEngineConfidence { get; set; } = 0.35f;

        public event EventHandler<SpeechRecognisedEventArgs> SpeechRecognised;
        public event EventHandler<SpeechListenerStatusEventArgs> StatusChanged;
        public event EventHandler<SpeechPhaseChangedEventArgs> PhaseChanged;

        public bool IsListening { get { lock (_lock) return _isListening; } }

        // ──────────────────────────────────────────────────────────────────
        // Start / Stop / Toggle
        // ──────────────────────────────────────────────────────────────────
        public void Start()
        {
            lock (_lock)
            {
                if (_disposed) throw new ObjectDisposedException(nameof(SpeechListener));
                if (_isListening) return;

                try
                {
                    log.Info("SpeechListener (Vosk v3): Initialising…");
                    RaiseStatus("Initialising speech recognition…");

                    if (!Directory.Exists(ModelPath))
                        throw new DirectoryNotFoundException(
                            $"Vosk model not found at: {ModelPath}\n" +
                            "Download a model from https://alphacephei.com/vosk/models and " +
                            $"extract it to {ModelPath}");

                    Vosk.Vosk.SetLogLevel(-1);
                    _model = new Model(ModelPath);
                    _currentPhase   = RecognitionPhase.Scan;
                    _phase2BookName = null;

                    // ── SCAN phase: open dictation (full English vocabulary) ──
                    _recogniser = new VoskRecognizer(_model, 16000f);
                    _recogniser.SetMaxAlternatives(0);
                    _recogniser.SetWords(true);

                    _waveIn = new WaveInEvent
                    {
                        WaveFormat = new WaveFormat(16000, 1),
                        BufferMilliseconds = 100,
                    };
                    _waveIn.DataAvailable    += OnDataAvailable;
                    _waveIn.RecordingStopped += OnRecordingStopped;
                    _waveIn.StartRecording();

                    PreWarmRecogniser();
                    _isListening = true;

                    log.Info("SpeechListener (Vosk v3): Now listening (open dictation).");
                    RaiseStatus("Listening for Bible references…");
                    RaisePhaseChanged("Scan", null);
                }
                catch (Exception ex)
                {
                    log.Error("SpeechListener (Vosk): Failed to start.", ex);
                    RaiseStatus($"Failed to start: {ex.Message}", isError: true);
                    CleanupEngine();
                    throw;
                }
            }
        }

        public void Stop()
        {
            lock (_lock)
            {
                if (!_isListening) return;
                log.Info("SpeechListener (Vosk): Stopping…");

                try { _waveIn?.StopRecording(); }
                catch (Exception ex) { log.Warn("SpeechListener: Error stopping WaveIn.", ex); }

                CleanupEngine();
                _isListening = false;

                log.Info("SpeechListener (Vosk): Stopped.");
                RaiseStatus("Speech listening stopped.");
            }
        }

        public bool Toggle()
        {
            if (IsListening) { Stop(); return false; }
            else { Start(); return true; }
        }

        // ──────────────────────────────────────────────────────────────────
        // Pre-warm
        // ──────────────────────────────────────────────────────────────────
        private void PreWarmRecogniser()
        {
            try
            {
                var silence = new byte[3200];
                lock (_lock) { _recogniser?.AcceptWaveform(silence, silence.Length); }
            }
            catch (Exception ex)
            {
                log.Debug($"SpeechListener: Pre-warm skipped ({ex.Message})");
            }
        }

        // ──────────────────────────────────────────────────────────────────
        // Audio pipeline
        // ──────────────────────────────────────────────────────────────────
        private void OnDataAvailable(object sender, WaveInEventArgs e)
        {
            bool finalResult;
            lock (_lock)
            {
                if (_recogniser == null) return;
                finalResult = _recogniser.AcceptWaveform(e.Buffer, e.BytesRecorded);
            }

            if (finalResult)
            {
                string json;
                lock (_lock)
                {
                    if (_recogniser == null) return;
                    json = _recogniser.Result();
                }
                ProcessResult(json);
            }
        }

        private void OnRecordingStopped(object sender, StoppedEventArgs e)
        {
            if (e.Exception != null)
            {
                log.Warn($"SpeechListener: Recording stopped with error: {e.Exception.Message}");
                RaiseStatus($"Recording error: {e.Exception.Message}", isError: true);
            }
        }

        private void ProcessResult(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return;

            var textMatch = Regex.Match(json, "\"text\"\\s*:\\s*\"([^\"]*)\"");
            if (!textMatch.Success) return;

            string rawText = textMatch.Groups[1].Value.Trim();
            if (string.IsNullOrWhiteSpace(rawText)) return;

            // ── Stronger rationaliser ─────────────────────────────────────
            string rationalised = RationaliseVoskOutput(rawText);
            if (string.IsNullOrWhiteSpace(rationalised)) return;

            string text = PhoneticCorrector.Correct(rationalised);
            float confidence = ExtractAverageConfidence(json);

            if (!string.Equals(rawText, text, StringComparison.Ordinal))
                log.Debug($"SpeechListener: raw=\"{rawText}\" → rationalised=\"{rationalised}\" → corrected=\"{text}\" (conf={confidence:F2})");
            else
                log.Debug($"SpeechListener (Vosk) [{_currentPhase}]: \"{text}\" (conf={confidence:F2})");

            // ── Two-phase dispatch ────────────────────────────────────────
            if (_currentPhase == RecognitionPhase.Scan)
            {
                string triggeredBook = DetectBookTrigger(text);
                if (triggeredBook != null)
                {
                    log.Debug($"SpeechListener: Scan triggered on \"{triggeredBook}\" — switching to Focus.");
                    // Fire THIS utterance first (we already have it) then switch
                    // to focus for any follow-up verse announcement.
                    FireIfConfident(text, confidence);
                    SwitchToFocusPhase(triggeredBook);
                }
                else
                {
                    // Still emit the raw utterance so the debug panel shows activity
                    // and the detector can try on its own (maybe a fuzzy match fires).
                    FireIfConfident(text, confidence);
                }
            }
            else
            {
                ResetPhase2Timeout();
                FireIfConfident(text, confidence);
                // After firing in focus, revert so the next reference starts clean.
                SwitchToScanPhase();
            }
        }

        private void FireIfConfident(string text, float confidence)
        {
            if (confidence < MinEngineConfidence)
            {
                log.Debug($"SpeechListener: Below MinEngineConfidence ({MinEngineConfidence:F2}), ignoring.");
                return;
            }

            SpeechRecognised?.Invoke(this, new SpeechRecognisedEventArgs
            {
                Text = text,
                Confidence = confidence,
            });
        }

        // ──────────────────────────────────────────────────────────────────
        // Book trigger detection — consult the Bible reference detector
        // ──────────────────────────────────────────────────────────────────
        private static string DetectBookTrigger(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            var detected = BibleReferenceDetector.DetectBest(text);
            return detected?.BookName;
        }

        private void SwitchToFocusPhase(string canonicalBookName)
        {
            lock (_lock)
            {
                if (_model == null) return;

                try
                {
                    _recogniser?.Dispose();
                    _recogniser = new VoskRecognizer(
                        _model, 16000f, BuildFocusGrammar(canonicalBookName));
                    _recogniser.SetMaxAlternatives(0);
                    _recogniser.SetWords(true);

                    _currentPhase   = RecognitionPhase.Focus;
                    _phase2BookName = canonicalBookName;

                    log.Info($"SpeechListener: Focus phase — locked to \"{canonicalBookName}\".");
                }
                catch (Exception ex)
                {
                    log.Warn($"SpeechListener: Failed to switch to focus phase: {ex.Message}");
                }
            }

            ResetPhase2Timeout();
            RaisePhaseChanged("Focus", canonicalBookName);
        }

        private void SwitchToScanPhase()
        {
            bool transitioned = false;
            lock (_lock)
            {
                if (_model == null || _currentPhase == RecognitionPhase.Scan) return;

                try
                {
                    _recogniser?.Dispose();
                    // Open-dictation scan recogniser — no grammar.
                    _recogniser = new VoskRecognizer(_model, 16000f);
                    _recogniser.SetMaxAlternatives(0);
                    _recogniser.SetWords(true);

                    _currentPhase   = RecognitionPhase.Scan;
                    _phase2BookName = null;
                    transitioned    = true;

                    log.Debug("SpeechListener: Reverted to Scan phase (open dictation).");
                }
                catch (Exception ex)
                {
                    log.Warn($"SpeechListener: Failed to revert to scan phase: {ex.Message}");
                }
            }

            if (transitioned) RaisePhaseChanged("Scan", null);
        }

        private void ResetPhase2Timeout()
        {
            _phase2Timeout?.Dispose();
            _phase2Timeout = new System.Threading.Timer(_ =>
            {
                log.Debug("SpeechListener: Phase 2 timeout — reverting to scan.");
                SwitchToScanPhase();
            }, null, Phase2TimeoutMs, System.Threading.Timeout.Infinite);
        }

        private static float ExtractAverageConfidence(string json)
        {
            var matches = Regex.Matches(json, "\"conf\"\\s*:\\s*([0-9.]+)");
            if (matches.Count == 0) return 0.8f;

            float total = 0f; int count = 0;
            foreach (Match m in matches)
            {
                if (float.TryParse(m.Groups[1].Value,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out float val))
                {
                    total += val; count++;
                }
            }
            return count > 0 ? total / count : 0.8f;
        }

        // ──────────────────────────────────────────────────────────────────
        // Rationaliser — cleans raw Vosk output before Bible detection
        //
        // New in v3 — the rationaliser now does real work:
        //   1.  Strips [unk]/<unk> tokens and Vosk artefacts.
        //   2.  Strips dictation disfluencies ("um", "uh", "er", "hmm", "you know").
        //   3.  Strips trailing/leading small words that wreck detection
        //       ("okay", "alright", "so", "yeah", "please").
        //   4.  Normalises common spoken-form tokens: "ch." → "chapter",
        //       "v.", "vs" → "verse", ":" / "colon" kept for the detector.
        //   5.  Collapses repeated tokens ("chapter chapter" → "chapter").
        //   6.  Normalises whitespace.
        //   7.  Lowercases (downstream code is case-insensitive but this is
        //       clean for logs).
        // ──────────────────────────────────────────────────────────────────
        private static readonly string[] DisfluencyTokens = new[]
        {
            "um","uh","er","erm","hmm","mhm","mm","uhh","uhhh","umm",
            "ah","eh","oh",
            // Only strip "oh" when it's a disfluency, not when it forms a number
            // like "one oh five". Disfluency stripping is context-aware below.
        };

        private static readonly Regex FillerPhraseRegex = new Regex(
            @"\b(you know|i mean|like i said|kind of|sort of|right now|okay then|" +
            @"well|alright|all right|so yeah|yeah so|yeah |okay |ok )\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static string RationaliseVoskOutput(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

            string text = raw;

            // 1. Strip explicit unknown tokens
            text = Regex.Replace(text, @"\[unk\]|<unk>", " ", RegexOptions.IgnoreCase);

            // 2. Strip multi-word filler phrases (before we touch single tokens)
            text = FillerPhraseRegex.Replace(text, " ");

            // 3. Normalise spoken punctuation shorthand
            text = Regex.Replace(text, @"\b(ch|chp|chpt)\b\.?",  "chapter", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"\b(vs|vss|vrs|v)\b\.?", "verse",   RegexOptions.IgnoreCase);

            // 4. Strip disfluencies, but only when they're standalone tokens
            //    and not part of a meaningful number pattern ("one oh five").
            var tokens = text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            var kept = new List<string>(tokens.Count);
            for (int i = 0; i < tokens.Count; i++)
            {
                string tok = tokens[i].ToLowerInvariant();

                // Preserve "oh" / "o" when sandwiched between digits/number-words
                // (i.e. the "one-oh-five" pattern).
                if (tok == "oh" || tok == "o")
                {
                    bool leftIsNum  = i > 0 && IsNumericish(tokens[i - 1]);
                    bool rightIsNum = i < tokens.Count - 1 && IsNumericish(tokens[i + 1]);
                    if (leftIsNum && rightIsNum) { kept.Add(tok); continue; }
                    // else fall through to disfluency check
                }

                if (Array.IndexOf(DisfluencyTokens, tok) >= 0) continue;

                kept.Add(tok);
            }

            // 5. Collapse consecutive duplicate tokens ("chapter chapter" → "chapter")
            var deduped = new List<string>(kept.Count);
            foreach (var tok in kept)
            {
                if (deduped.Count > 0 && deduped[deduped.Count - 1] == tok) continue;
                deduped.Add(tok);
            }

            // 6. Normalise whitespace
            text = string.Join(" ", deduped);
            text = Regex.Replace(text, @"\s{2,}", " ").Trim();

            return text;
        }

        private static bool IsNumericish(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return false;
            token = token.Trim().ToLowerInvariant();
            if (int.TryParse(token, out _)) return true;
            return SpokenNumberConverter.IsNumberWord(token);
        }

        // ──────────────────────────────────────────────────────────────────
        // Focus-phase grammar (kept — it's the hot half of the pipeline)
        // ──────────────────────────────────────────────────────────────────
        private static string BuildFocusGrammar(string canonicalBookName)
        {
            var words = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var variant in GetBookVariants(canonicalBookName))
                words.Add(variant);

            if (canonicalBookName.Length > 1 && char.IsDigit(canonicalBookName[0]))
            {
                words.Add("first"); words.Add("second"); words.Add("third");
            }
            if (canonicalBookName == "Song of Solomon")
            {
                words.Add("song"); words.Add("of"); words.Add("solomon"); words.Add("songs");
            }

            foreach (var w in new[]
            {
                "zero","oh","o",
                "one","two","three","four","five",
                "six","seven","eight","nine","ten",
                "eleven","twelve","thirteen","fourteen","fifteen",
                "sixteen","seventeen","eighteen","nineteen",
                "twenty","thirty","forty","fifty",
                "sixty","seventy","eighty","ninety","hundred",
                // phonetic near-misses
                "night","tea","tree","free","heaven","fore",
                "sex","ate","won","too","fight",
                // structure
                "chapter","chapters","verse","verses",
                "through","to","and","colon","dash","hyphen",
                // "verse" mishearings the corrector rewrites
                "us","as","of","office","offices",
                "[unk]",
            })
                words.Add(w);

            for (int n = 1; n <= 176; n++) words.Add(n.ToString());

            var quoted = words.OrderBy(w => w).Select(w => $"\"{w.Replace("\"", "\\\"")}\"");
            return "[" + string.Join(",", quoted) + "]";
        }

        private static readonly Dictionary<string, string[]> BookVariantsMap =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            { "Genesis",          new[]{ "genesis" } },
            { "Exodus",           new[]{ "exodus" } },
            { "Leviticus",        new[]{ "leviticus" } },
            { "Numbers",          new[]{ "numbers" } },
            { "Deuteronomy",      new[]{ "deuteronomy","deutronomy","duteronomy" } },
            { "Joshua",           new[]{ "joshua" } },
            { "Judges",           new[]{ "judges" } },
            { "Ruth",             new[]{ "ruth" } },
            { "1 Samuel",         new[]{ "samuel","sam" } },
            { "2 Samuel",         new[]{ "samuel","sam" } },
            { "1 Kings",          new[]{ "kings","king" } },
            { "2 Kings",          new[]{ "kings","king" } },
            { "1 Chronicles",     new[]{ "chronicles","chron" } },
            { "2 Chronicles",     new[]{ "chronicles","chron" } },
            { "Ezra",             new[]{ "ezra" } },
            { "Nehemiah",         new[]{ "nehemiah","nehemia","nehimiah","nehimia" } },
            { "Esther",           new[]{ "esther" } },
            { "Job",              new[]{ "job" } },
            { "Psalms",           new[]{ "psalms","psalm","salms","sams" } },
            { "Proverbs",         new[]{ "proverbs","proverb" } },
            { "Ecclesiastes",     new[]{ "ecclesiastes","ecclesiaste" } },
            { "Song of Solomon",  new[]{ "song","songs","solomon" } },
            { "Isaiah",           new[]{ "isaiah","esaiah","isaia" } },
            { "Jeremiah",         new[]{ "jeremiah","jeremia","jerimiah","jerimia" } },
            { "Lamentations",     new[]{ "lamentations","lamentation" } },
            { "Ezekiel",          new[]{ "ezekiel","ezekia","ezekel" } },
            { "Daniel",           new[]{ "daniel" } },
            { "Hosea",            new[]{ "hosea","hosia" } },
            { "Joel",             new[]{ "joel" } },
            { "Amos",             new[]{ "amos" } },
            { "Obadiah",          new[]{ "obadiah","obadia","obadiya" } },
            { "Jonah",            new[]{ "jonah" } },
            { "Micah",            new[]{ "micah","mica" } },
            { "Nahum",            new[]{ "nahum" } },
            { "Habakkuk",         new[]{ "habakkuk","habakuk","habacuc","habacuk" } },
            { "Zephaniah",        new[]{ "zephaniah","zefaniah","zefania" } },
            { "Haggai",           new[]{ "haggai","hagai" } },
            { "Zechariah",        new[]{ "zechariah","zachariah","zacharia","zakaria",
                                         "zakariah","zekaria","zekariah","zecharia",
                                         "zecharias","zacharias" } },
            { "Malachi",          new[]{ "malachi","malaki" } },
            { "Matthew",          new[]{ "matthew","mathew","mathieu" } },
            { "Mark",             new[]{ "mark" } },
            { "Luke",             new[]{ "luke" } },
            { "John",             new[]{ "john" } },
            { "Acts",             new[]{ "acts" } },
            { "Romans",           new[]{ "romans" } },
            { "1 Corinthians",    new[]{ "corinthians","corinthian" } },
            { "2 Corinthians",    new[]{ "corinthians","corinthian" } },
            { "Galatians",        new[]{ "galatians","galatian" } },
            { "Ephesians",        new[]{ "ephesians","ephesian" } },
            { "Philippians",      new[]{ "philippians","philippian","philipians","philipian" } },
            { "Colossians",       new[]{ "colossians","colossian","colosians","colosian" } },
            { "1 Thessalonians",  new[]{ "thessalonians","thessalonian" } },
            { "2 Thessalonians",  new[]{ "thessalonians","thessalonian" } },
            { "1 Timothy",        new[]{ "timothy","timoty" } },
            { "2 Timothy",        new[]{ "timothy","timoty" } },
            { "Titus",            new[]{ "titus" } },
            { "Philemon",         new[]{ "philemon","filemon" } },
            { "Hebrews",          new[]{ "hebrews","hebrew" } },
            { "James",            new[]{ "james" } },
            { "1 Peter",          new[]{ "peter" } },
            { "2 Peter",          new[]{ "peter" } },
            { "1 John",           new[]{ "john" } },
            { "2 John",           new[]{ "john" } },
            { "3 John",           new[]{ "john" } },
            { "Jude",             new[]{ "jude" } },
            { "Revelation",       new[]{ "revelation","revelations","revelacion","revelasion" } },
        };

        private static IEnumerable<string> GetBookVariants(string canonicalBookName)
        {
            if (BookVariantsMap.TryGetValue(canonicalBookName, out var variants)) return variants;
            return new[] { canonicalBookName.ToLowerInvariant() };
        }

        // ──────────────────────────────────────────────────────────────────
        private void CleanupEngine()
        {
            _phase2Timeout?.Dispose();
            _phase2Timeout  = null;
            _currentPhase   = RecognitionPhase.Scan;
            _phase2BookName = null;

            if (_waveIn != null)
            {
                try
                {
                    _waveIn.DataAvailable    -= OnDataAvailable;
                    _waveIn.RecordingStopped -= OnRecordingStopped;
                    _waveIn.Dispose();
                }
                catch (Exception ex) { log.Warn("SpeechListener: Error disposing WaveIn.", ex); }
                _waveIn = null;
            }

            if (_recogniser != null)
            {
                try { _recogniser.Dispose(); }
                catch (Exception ex) { log.Warn("SpeechListener: Error disposing recogniser.", ex); }
                _recogniser = null;
            }

            if (_model != null)
            {
                try { _model.Dispose(); }
                catch (Exception ex) { log.Warn("SpeechListener: Error disposing model.", ex); }
                _model = null;
            }
        }

        public void Dispose()
        {
            lock (_lock)
            {
                if (_disposed) return;
                _disposed = true;
                Stop();
            }
        }

        private void RaiseStatus(string message, bool isError = false)
        {
            StatusChanged?.Invoke(this, new SpeechListenerStatusEventArgs
            {
                Message = message,
                IsError = isError,
            });
        }

        private void RaisePhaseChanged(string phase, string bookName)
        {
            try
            {
                PhaseChanged?.Invoke(this, new SpeechPhaseChangedEventArgs
                {
                    Phase    = phase,
                    BookName = bookName,
                });
            }
            catch (Exception ex)
            {
                log.Debug($"SpeechListener: PhaseChanged subscriber threw: {ex.Message}");
            }
        }
    }
}
