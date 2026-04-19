// ============================================================================
// SpeechListener.cs
// Real-time speech recognition using Vosk (offline, free, open-source).
//
// Replaces the old System.Speech implementation. Drop-in compatible —
// same public API, same events. No cloud, no cost, much better accuracy.
//
// Prerequisites (NuGet):
//   Vosk                  — add via NuGet Package Manager
//   NAudio                — add via NuGet Package Manager (audio capture)
//
// Model download (one-time, manual step — see SETUP.md):
//   Recommended (~1.8 GB, best accuracy):
//     https://alphacephei.com/vosk/models/vosk-model-en-us-0.22.zip
//   Fallback (~50 MB, fast load):
//     https://alphacephei.com/vosk/models/vosk-model-small-en-us-0.15.zip
//   Extract whichever you choose to: %APPDATA%\WorshipHelper\vosk-model\
//
//   The large model is strongly preferred — it significantly improves accuracy
//   for accented speech and reduces phonetic misfires (e.g. "eight" → "amos").
//   On modern hardware (8 GB+ RAM, SSD) the load time difference is ~2 seconds.
//
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

    /// <summary>
    /// Carries phase-change notifications so the UI can show whether the
    /// listener is in the broad SCAN phase or the book-locked FOCUS phase.
    /// </summary>
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

        // Model path resolution order:
        // 1. Next to the installed DLL (data\vosk-model\) — bundled by the MSI
        // 2. %APPDATA%\WorshipHelper\vosk-model\ — manual install fallback
        private static readonly string DefaultModelPath = ResolveModelPath();

        private static string ResolveModelPath()
        {
            // VSTO shadow-copies the DLL to a temp folder, so
            // Assembly.Location doesn't point to the install directory.
            // Read InstallLocation from the registry — the MSI writes it.
            string installDir = Microsoft.Win32.Registry.CurrentUser
                .OpenSubKey(@"SOFTWARE\WorshipHelper")
                ?.GetValue("InstallLocation") as string;

            if (!string.IsNullOrEmpty(installDir))
            {
                string registryPath = Path.Combine(installDir, "data", "vosk-model");
                if (Directory.Exists(registryPath)) return registryPath;
            }

            // Dev/debug fallback: next to the DLL (works without VSTO shadow-copying)
            string dllDir = Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location) ?? "";
            string bundled = Path.Combine(dllDir, "data", "vosk-model");
            if (Directory.Exists(bundled)) return bundled;

            // Last resort: manual install in AppData
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "WorshipHelper", "vosk-model");
        }

        private VoskRecognizer _recogniser;
        private WaveInEvent _waveIn;
        private bool _isListening;
        private bool _disposed;
        private readonly object _lock = new object();

        // -------------------------------------------------------------------------
        // Two-phase recognition state
        //
        // Phase 1 (SCAN): broad grammar — all book names + number words.
        //   Vosk listens for any book name. As soon as one fires, switch to phase 2.
        //
        // Phase 2 (FOCUS): hyper-tight grammar — ONLY the detected book's name
        //   variants + number words + structure words. ~50-70 words max.
        //   Vosk has almost nowhere else to go, so chapter/verse numbers come
        //   out cleanly.
        //
        // After phase 2 fires (or after a timeout), revert to phase 1.
        //
        // Why this works:
        //   Big vocab → Vosk picks nearest-sounding common word for rare names.
        //   e.g. "zechariah four nine" → "zakariah four night" (near-miss).
        //   Focused vocab → only zechariah variants + numbers in scope.
        //   "zechariah four nine" → "zechariah four nine". Done.
        // -------------------------------------------------------------------------
        private Model _model;                    // loaded once, shared across phases
        private string _phase2BookName;          // canonical book name we locked onto
        private System.Threading.Timer _phase2Timeout; // revert to phase 1 after idle
        private const int Phase2TimeoutMs = 4000;      // 4s with no result → back to scan

        private enum RecognitionPhase { Scan, Focus }
        private RecognitionPhase _currentPhase = RecognitionPhase.Scan;

        // -------------------------------------------------------------------------
        // Configuration
        // -------------------------------------------------------------------------

        /// <summary>
        /// Path to the extracted Vosk model directory.
        /// Defaults to %APPDATA%\WorshipHelper\vosk-model\
        /// </summary>
        public string ModelPath { get; set; } = DefaultModelPath;

        /// <summary>
        /// Minimum confidence (0.0-1.0) before SpeechRecognised fires.
        /// Default: 0.45
        /// </summary>
        public float MinEngineConfidence { get; set; } = 0.45f;

        // -------------------------------------------------------------------------
        // Events
        // -------------------------------------------------------------------------

        public event EventHandler<SpeechRecognisedEventArgs> SpeechRecognised;
        public event EventHandler<SpeechListenerStatusEventArgs> StatusChanged;
        public event EventHandler<SpeechPhaseChangedEventArgs> PhaseChanged;

        public bool IsListening
        {
            get { lock (_lock) return _isListening; }
        }

        // -------------------------------------------------------------------------
        // Start / Stop / Toggle
        // -------------------------------------------------------------------------

        public void Start()
        {
            lock (_lock)
            {
                if (_disposed) throw new ObjectDisposedException(nameof(SpeechListener));
                if (_isListening) return;

                try
                {
                    log.Info("SpeechListener (Vosk): Initialising...");
                    RaiseStatus("Initialising speech recognition...");

                    if (!Directory.Exists(ModelPath))
                        throw new DirectoryNotFoundException(
                            $"Vosk model not found at: {ModelPath}\n" +
                            "Download a model from https://alphacephei.com/vosk/models and " +
                            $"extract it to {ModelPath}");

                    Vosk.Vosk.SetLogLevel(-1);

                    _model = new Model(ModelPath);
                    _currentPhase = RecognitionPhase.Scan;
                    _phase2BookName = null;
                    _recogniser = new VoskRecognizer(_model, 16000f, BuildScanGrammar());
                    _recogniser.SetMaxAlternatives(0);
                    _recogniser.SetWords(true);

                    _waveIn = new WaveInEvent
                    {
                        WaveFormat = new WaveFormat(16000, 1),
                        BufferMilliseconds = 100,
                    };

                    _waveIn.DataAvailable += OnDataAvailable;
                    _waveIn.RecordingStopped += OnRecordingStopped;
                    _waveIn.StartRecording();

                    // Pre-warm the acoustic model: the first AcceptWaveform call after
                    // load is slow (JIT + model init). Feeding a silent buffer here eats
                    // that cost during startup rather than on the first real utterance,
                    // which prevents a noticeable freeze mid-service.
                    PreWarmRecogniser();

                    _isListening = true;

                    log.Info("SpeechListener (Vosk): Now listening.");
                    RaiseStatus("Listening for Bible references...");
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

                log.Info("SpeechListener (Vosk): Stopping...");

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

        // -------------------------------------------------------------------------
        // Pre-warm
        // -------------------------------------------------------------------------

        /// <summary>
        /// Feeds 100ms of silence through the recogniser immediately after startup.
        /// The large Vosk model (vosk-model-en-us-0.22) defers some initialisation
        /// to the first AcceptWaveform call. Without pre-warming this causes a ~1s
        /// freeze on the very first real utterance. Running it here moves that cost
        /// to the background start sequence where it is invisible to the user.
        /// </summary>
        private void PreWarmRecogniser()
        {
            try
            {
                // 16000 samples/sec × 1 channel × 2 bytes/sample × 0.1s = 3200 bytes of silence
                var silence = new byte[3200];
                lock (_lock)
                {
                    _recogniser?.AcceptWaveform(silence, silence.Length);
                }
            }
            catch (Exception ex)
            {
                log.Debug($"SpeechListener: Pre-warm skipped ({ex.Message})");
            }
        }

        // -------------------------------------------------------------------------
        // Audio pipeline
        // -------------------------------------------------------------------------

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

        /// <summary>
        /// Parses a Vosk JSON result and fires SpeechRecognised if confidence passes.
        ///
        /// Vosk result JSON example:
        /// {
        ///   "result": [
        ///     {"conf": 0.98, "word": "john"},
        ///     {"conf": 0.91, "word": "three"},
        ///     {"conf": 0.87, "word": "sixteen"}
        ///   ],
        ///   "text": "john three sixteen"
        /// }
        /// </summary>
        private void ProcessResult(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return;

            var textMatch = Regex.Match(json, "\"text\"\\s*:\\s*\"([^\"]+)\"");
            if (!textMatch.Success) return;

            string rawText = textMatch.Groups[1].Value.Trim();
            if (string.IsNullOrWhiteSpace(rawText)) return;

            string rationalised = RationaliseVoskOutput(rawText);
            if (string.IsNullOrWhiteSpace(rationalised)) return;

            string text = PhoneticCorrector.Correct(rationalised);
            float confidence = ExtractAverageConfidence(json);

            if (rawText != text)
                log.Debug($"SpeechListener: raw=\"{rawText}\" → corrected=\"{text}\" (conf={confidence:F2})");
            else
                log.Debug($"SpeechListener (Vosk) [{_currentPhase}]: \"{text}\" (conf={confidence:F2})");

            // -----------------------------------------------------------------------
            // Two-phase dispatch
            // -----------------------------------------------------------------------
            if (_currentPhase == RecognitionPhase.Scan)
            {
                // Phase 1: look for any book name trigger in the output.
                // If found, switch to hyper-focused phase 2 for that book and
                // re-process this same utterance immediately — we already have it.
                string triggeredBook = DetectBookTrigger(text);
                if (triggeredBook != null)
                {
                    log.Debug($"SpeechListener: Scan triggered on \"{triggeredBook}\" — switching to Focus phase.");
                    SwitchToFocusPhase(triggeredBook);

                    // Re-process this utterance now that we're in focus phase.
                    // The focus grammar would have produced cleaner output, but we
                    // already have the corrected text so run it through the pipeline.
                    FireIfConfident(text, confidence);
                }
                // If no book detected in scan phase, do nothing — wait for next utterance.
            }
            else
            {
                // Phase 2: we're focused on a specific book.
                // Whatever comes out is our best shot at the reference.
                // Reset the timeout, fire if confident, then snap back to scan.
                ResetPhase2Timeout();
                FireIfConfident(text, confidence);
                // Revert immediately after firing so we're ready for the next reference
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

        // -----------------------------------------------------------------------
        // Two-phase helpers
        // -----------------------------------------------------------------------

        /// <summary>
        /// Scans corrected text for any known book name or phonetic variant.
        /// Returns the canonical book name if found, null otherwise.
        /// Uses the same BibleReferenceDetector book matching so it's consistent.
        /// </summary>
        private static string DetectBookTrigger(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            var detected = BibleReferenceDetector.DetectBest(text);
            return detected?.BookName;
        }

        /// <summary>
        /// Switches the recogniser to a hyper-tight grammar focused on one book.
        /// The new recogniser uses the same loaded model — no disk I/O.
        /// </summary>
        private void SwitchToFocusPhase(string canonicalBookName)
        {
            lock (_lock)
            {
                if (_model == null) return;

                try
                {
                    _recogniser?.Dispose();
                    _recogniser = new VoskRecognizer(
                        _model, 16000f,
                        BuildFocusGrammar(canonicalBookName));
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

        /// <summary>
        /// Reverts to scan phase (broad grammar).
        /// Called after phase 2 fires or after the idle timeout.
        /// </summary>
        private void SwitchToScanPhase()
        {
            bool transitioned = false;
            lock (_lock)
            {
                if (_model == null || _currentPhase == RecognitionPhase.Scan) return;

                try
                {
                    _recogniser?.Dispose();
                    _recogniser = new VoskRecognizer(_model, 16000f, BuildScanGrammar());
                    _recogniser.SetMaxAlternatives(0);
                    _recogniser.SetWords(true);

                    _currentPhase   = RecognitionPhase.Scan;
                    _phase2BookName = null;
                    transitioned    = true;

                    log.Debug("SpeechListener: Reverted to Scan phase.");
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

            float total = 0f;
            int count = 0;
            foreach (System.Text.RegularExpressions.Match m in matches)
            {
                if (float.TryParse(m.Groups[1].Value,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out float val))
                {
                    total += val;
                    count++;
                }
            }
            return count > 0 ? total / count : 0.8f;
        }

        // -------------------------------------------------------------------------
        // Rationaliser
        // -------------------------------------------------------------------------

        /// <summary>
        /// Cleans raw Vosk output before it enters the Bible reference detection
        /// pipeline.  Strips [unk] tokens produced by the grammar for out-of-
        /// vocabulary phonemes, then collapses any resulting extra whitespace.
        /// </summary>
        private static string RationaliseVoskOutput(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            string text = Regex.Replace(raw, @"\[unk\]", " ", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"\s{2,}", " ").Trim();
            return text;
        }

        // -------------------------------------------------------------------------
        // Grammar
        // -------------------------------------------------------------------------

        // -------------------------------------------------------------------------
        // Grammars
        // -------------------------------------------------------------------------

        /// <summary>
        /// Phase 1 (SCAN) grammar: all book name variants + number words.
        /// Broad enough to catch any book name, including phonetic near-misses.
        /// When a book name fires here, we switch to the focused phase 2 grammar.
        /// </summary>
        private static string BuildScanGrammar()
        {
            var words = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                // ── All book name words + phonetic variants (OT) ───────────────
                "genesis",
                "exodus",
                "leviticus",
                "numbers",
                "deuteronomy","deutronomy","duteronomy",
                "joshua",
                "judges",
                "ruth",
                "samuel",
                "kings","king",
                "chronicles",
                "ezra",
                "nehemiah","nehemia","nehimiah","nehimia",
                "esther",
                "job",
                "psalms","psalm","salms","sams",
                "proverbs","proverb",
                "ecclesiastes","ecclesiaste",
                "song","solomon","songs",
                "isaiah","esaiah","isaia",
                "jeremiah","jeremia","jerimiah","jerimia",
                "lamentations","lamentation",
                "ezekiel","ezekia","ezekel",
                "daniel",
                "hosea","hosia",
                "joel",
                "amos",
                "obadiah","obadia","obadiya",
                "jonah",
                "micah","mica",
                "nahum",
                "habakkuk","habakuk","habacuc","habacuk",
                "zephaniah","zefaniah","zefania",
                "haggai","hagai",
                // Zechariah — all phonetic variants so ANY mishearing triggers phase 2
                "zechariah","zachariah","zacharia","zakaria",
                "zakariah","zekaria","zekariah","zecharia","zecharias","zacharias",
                "malachi","malaki",

                // ── All book name words (NT) ───────────────────────────────────
                "matthew","mathew","mathieu",
                "mark",
                "luke",
                "john",
                "acts",
                "romans",
                "corinthians","corinthian",
                "galatians","galatian",
                "ephesians","ephesian",
                "philippians","philippian","philipians","philipian",
                "colossians","colossian","colosians","colosian",
                "thessalonians","thessalonian",
                "timothy","timoty",
                "titus",
                "philemon","filemon",
                "hebrews","hebrew",
                "james",
                "peter",
                "jude",
                "revelation","revelations","revelacion","revelasion",

                // ── Numbered-book prefixes ─────────────────────────────────────
                "first","second","third",
                "of",

                // ── Number words (needed so whole references in one utterance work) ──
                "zero","oh","o",
                "one","two","three","four","five",
                "six","seven","eight","nine","ten",
                "eleven","twelve","thirteen","fourteen","fifteen",
                "sixteen","seventeen","eighteen","nineteen",
                "twenty","thirty","forty","fifty",
                "sixty","seventy","eighty","ninety",
                "hundred",

                // ── Phonetic near-misses for numbers ──────────────────────────
                "night","tea","tree","free","heaven","fore",
                "sex","ate","won","too","fight",

                // ── Structure words ────────────────────────────────────────────
                "chapter","chapters","verse","verses",
                "through","to","and","colon","dash","hyphen",
                "scripture",

                // ── Common "verse" mishearings ────────────────────────────────
                // Without these in the grammar, Vosk maps /vɜːs/ to [unk] and
                // the rationaliser strips it, turning "chapter 1 verse 1" into
                // the ambiguous "chapter 1 1". Including "us"/"as"/"office"
                // lets Vosk output e.g. "of us" which PhoneticCorrector then
                // rewrites back to "verse".
                "us","as","office","offices",

                "[unk]",
            };

            for (int n = 1; n <= 176; n++) words.Add(n.ToString());
            var quoted = words.OrderBy(w => w).Select(w => $"\"{w.Replace("\"", "\\\"")}\"");
            return "[" + string.Join(",", quoted) + "]";
        }

        /// <summary>
        /// Phase 2 (FOCUS) grammar: hyper-tight — only the specific book's name
        /// variants + number words + structure words. ~50-70 words max.
        ///
        /// Vosk has almost nowhere else to go, so chapter/verse numbers come out
        /// cleanly. e.g. for Zechariah the grammar is just zechariah variants +
        /// number words — "four nine" can only map to "four" and "nine".
        ///
        /// This answers the "what if zechariah itself isn't heard in scan phase"
        /// question: the scan grammar includes ALL phonetic variants of zechariah
        /// (zachariah, zacharia, zakaria, zekaria, etc.). Any of them triggers
        /// phase 2. Phase 2 then uses those same variants so whatever vosk outputs
        /// for the book name still matches correctly.
        /// </summary>
        private static string BuildFocusGrammar(string canonicalBookName)
        {
            var words = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Add the specific book's canonical name and all known variants
            foreach (var variant in GetBookVariants(canonicalBookName))
                words.Add(variant);

            // Numbered book prefix words if needed
            if (canonicalBookName.Length > 1 && char.IsDigit(canonicalBookName[0]))
            {
                words.Add("first"); words.Add("second"); words.Add("third");
            }

            // "of" for Song of Solomon
            if (canonicalBookName == "Song of Solomon")
            {
                words.Add("song"); words.Add("of"); words.Add("solomon"); words.Add("songs");
            }

            // Number words — the whole point of phase 2
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
                // "verse" mishearings — rewritten by PhoneticCorrector
                "us","as","of","office","offices",
                "[unk]",
            })
                words.Add(w);

            for (int n = 1; n <= 176; n++) words.Add(n.ToString());

            var quoted = words.OrderBy(w => w).Select(w => $"\"{w.Replace("\"", "\\\"")}\"");
            return "[" + string.Join(",", quoted) + "]";
        }

        /// <summary>
        /// Returns all spoken variants for a canonical book name.
        /// Mirrors BibleReferenceDetector's variant list so they're always in sync.
        /// </summary>
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
            if (BookVariantsMap.TryGetValue(canonicalBookName, out var variants))
                return variants;
            // Fallback: just lowercase canonical
            return new[] { canonicalBookName.ToLowerInvariant() };
        }

        // -------------------------------------------------------------------------
        // -------------------------------------------------------------------------

        private void CleanupEngine()
        {
            _phase2Timeout?.Dispose();
            _phase2Timeout = null;
            _currentPhase  = RecognitionPhase.Scan;
            _phase2BookName = null;

            if (_waveIn != null)
            {
                try
                {
                    _waveIn.DataAvailable -= OnDataAvailable;
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
                    Phase = phase,
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
