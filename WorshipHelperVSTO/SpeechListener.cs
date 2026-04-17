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
//   Small (~50 MB, fast):  https://alphacephei.com/vosk/models/vosk-model-small-en-us-0.15.zip
//   Large (~1.8 GB, best): https://alphacephei.com/vosk/models/vosk-model-en-us-0.22.zip
//   Extract to:            %APPDATA%\WorshipHelper\vosk-model\
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

                    var model = new Model(ModelPath);
                    _recogniser = new VoskRecognizer(model, 16000f, BuildBibleGrammar());
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

                    _isListening = true;

                    log.Info("SpeechListener (Vosk): Now listening.");
                    RaiseStatus("Listening for Bible references...");
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

            // Rationalise: strip [unk] tokens, normalise spoken punctuation, etc.
            string rationalised = RationaliseVoskOutput(rawText);
            if (string.IsNullOrWhiteSpace(rationalised)) return;

            // Phonetic correction: fix systematic Vosk mishearings before detection.
            // e.g. "zechariah fight for tea night" → "zechariah four three nine"
            string text = PhoneticCorrector.Correct(rationalised);

            float confidence = ExtractAverageConfidence(json);

            if (rawText != text)
                log.Debug($"SpeechListener: raw=\"{rawText}\" → corrected=\"{text}\" (conf={confidence:F2})");
            else
                log.Debug($"SpeechListener (Vosk): \"{text}\" (conf={confidence:F2})");

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

        /// <summary>
        /// Builds a JSON word-list that constrains Vosk to only Bible-relevant
        /// vocabulary.  Without this, Vosk maps "Zechariah" phonemes to random
        /// English words ("zachariah night birth a to her").
        ///
        /// "[unk]" is included so that off-topic speech produces an [unk] token
        /// instead of a false-positive book name match.
        ///
        /// Digit strings ("1"–"176") are included because Vosk sometimes outputs
        /// them even in grammar mode, and WordsToNumber already handles them.
        /// </summary>
        private static string BuildBibleGrammar()
        {
            var words = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                // ── Book name words (OT) ───────────────────────────────────────
                "genesis","gen",
                "exodus","exod",
                "leviticus","lev",
                "numbers","num",
                "deuteronomy","deut",
                "joshua","josh",
                "judges","judg",
                "ruth",
                "samuel","sam",
                "kings","king",
                "chronicles","chron",
                "ezra",
                "nehemiah","neh",
                "esther",
                "job",
                "psalms","psalm","psa",
                "proverbs","proverb","prov",
                "ecclesiastes","eccl","eccles",
                "song","solomon","songs",
                "isaiah","isa",
                "jeremiah","jer",
                "lamentations","lam",
                "ezekiel","ezek",
                "daniel","dan",
                "hosea","hos",
                "joel",
                "amos",
                "obadiah","obad",
                "jonah",
                "micah","mic",
                "nahum","nah",
                "habakkuk","hab",
                "zephaniah","zeph",
                "haggai","hag",
                "zechariah","zech",
                "malachi","mal",

                // ── Book name words (NT) ───────────────────────────────────────
                "matthew","matt","mat",
                "mark",
                "luke",
                "john","jn",
                "acts","act",
                "romans","rom",
                "corinthians","cor",
                "galatians","gal",
                "ephesians","eph",
                "philippians","phil","php",
                "colossians","col",
                "thessalonians","thess",
                "timothy","tim",
                "titus","tit",
                "philemon","phlm","philem",
                "hebrews","heb",
                "james","jas",
                "peter","pet",
                "jude",
                "revelation","revelations","rev",

                // ── Numbered-book prefixes ─────────────────────────────────────
                "first","second","third",
                "i","ii","iii",
                "of",   // "song of solomon"

                // ── Number words ───────────────────────────────────────────────
                "zero","oh","o",
                "one","two","three","four","five",
                "six","seven","eight","nine","ten",
                "eleven","twelve","thirteen","fourteen","fifteen",
                "sixteen","seventeen","eighteen","nineteen",
                "twenty","thirty","forty","fifty",
                "sixty","seventy","eighty","ninety",
                "hundred",

                // ── Reference connector / structure words ──────────────────────
                "chapter","chapters",
                "verse","verses",
                "through","to","and","colon","dash","hyphen",

                // ── Common preamble / filler words ────────────────────────────
                "read","reading","turn","turning","open","opening",
                "look","looking","go","going","find","finding",
                "let","lets","us","please","now","okay","ok",
                "the","book","passage","scripture","text",
                "today","tonight","this","morning","evening",
                "says","we're","i'm","from","at","in",

                // ── Unknown-word escape hatch ──────────────────────────────────
                // Without this Vosk forces every phoneme into our vocabulary,
                // causing false positives from non-Bible speech.
                "[unk]",
            };

            // Also add digit strings 1–176 (max Bible verse number).
            // Vosk can output these even in grammar mode, and WordsToNumber
            // already handles digit strings via int.TryParse.
            for (int n = 1; n <= 176; n++)
                words.Add(n.ToString());

            var quoted = words.OrderBy(w => w).Select(w => $"\"{w.Replace("\"", "\\\"")}\"");
            return "[" + string.Join(",", quoted) + "]";
        }

        // -------------------------------------------------------------------------
        // -------------------------------------------------------------------------

        private void CleanupEngine()
        {
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
    }
}
