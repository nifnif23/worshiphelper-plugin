// ============================================================================
// SpeechListener.cs  --  v5
//
// v4 was a thin facade; v5 makes it a smart facade:
//   * Forwards the engine's quality metrics (logprob / compression / no_speech)
//     so downstream gates (HallucinationGuard, SpeechToScriptureService) can
//     use them.
//   * Own small dedup cache -- if the server somehow forwards two identical
//     transcripts inside 2s we absorb the repeat here.
//   * Hotword push: on Start() we ship the full book-name list so Whisper
//     knows those are real words without getting biased toward example refs.
//
// Public API preserved (SpeechRecognised, StatusChanged, PhaseChanged,
// Start/Stop/Toggle, IsListening, MinEngineConfidence, ServerUri).
// Consumers of v4 keep working; new consumers can opt in to richer signals
// via the RawTranscriptReceived event.
// ============================================================================
using System;
using log4net;
using WorshipHelperVSTO.Audio;
using WorshipHelperVSTO.Networking;

namespace WorshipHelperVSTO
{
    public class SpeechRecognisedEventArgs : EventArgs
    {
        public string Text { get; set; }
        public float  Confidence { get; set; }

        // NEW in v5 -- richer trust metrics from the Whisper engine.
        public double AvgLogProb { get; set; }
        public double NoSpeechProb { get; set; }
        public double CompressionRatio { get; set; }
        public double DurationSeconds { get; set; }
    }

    public class SpeechListenerStatusEventArgs : EventArgs
    {
        public string Message { get; set; }
        public bool   IsError { get; set; }
    }

    public class SpeechPhaseChangedEventArgs : EventArgs
    {
        public string Phase { get; set; }       // "Scan" | "Focus"
        public string BookName { get; set; }
    }

    public class SegmentDroppedEventArgs : EventArgs
    {
        public string Reason { get; set; }
        public string Text { get; set; }
        public double DurationSeconds { get; set; }
    }

    public sealed class SpeechListener : IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(SpeechListener));

        private readonly MicrophoneCapture _mic     = new MicrophoneCapture();
        private readonly Chunker           _chunker = new Chunker();
        private readonly PythonClient      _client  = new PythonClient();

        private readonly object _lock = new object();
        private bool _isListening;
        private bool _disposed;
        private string _focusBook;
        private System.Threading.Timer _phaseTimeout;
        private const int FocusTimeoutMs = 5_000;

        // Local dedup -- defensive. Server already dedups; this is the belt on
        // top of the braces.
        private string _lastEmittedNormalised;
        private DateTime _lastEmittedUtc;
        private static readonly TimeSpan _dedupWindow = TimeSpan.FromSeconds(2);

        public float MinEngineConfidence { get; set; } = 0.15f;

        public Uri ServerUri
        {
            get => _client.ServerUri;
            set => _client.ServerUri = value;
        }

        public int LastPingMs => _client.LastPingMs;

        // -- Events ----------------------------------------------------------
        public event EventHandler<SpeechRecognisedEventArgs>     SpeechRecognised;
        public event EventHandler<SpeechListenerStatusEventArgs> StatusChanged;
        public event EventHandler<SpeechPhaseChangedEventArgs>   PhaseChanged;
        public event EventHandler<SegmentDroppedEventArgs>       SegmentDropped;

        public bool IsListening { get { lock (_lock) return _isListening; } }

        // ---------------------------------------------------------------
        public SpeechListener()
        {
            _mic.PcmFrame        += (s, pcm) => _chunker.Feed(pcm);
            _mic.CaptureError    += (s, ex)  => RaiseStatus("Mic error: " + ex.Message, true);
            _chunker.ChunkReady  += OnChunkReady;
            _client.TranscriptReceived += OnTranscript;
            _client.SegmentDropped     += OnSegmentDropped;
            _client.Connected          += async (s, e) =>
            {
                RaiseStatus("Connected to STT server.");
                // Push canonical book names so Whisper recognises them.
                try { await _client.SendHotwordsAsync(BookHotwords.Default); }
                catch (Exception ex) { log.Debug("hotwords push failed: " + ex.Message); }
            };
            _client.Disconnected       += (s, e) =>
                RaiseStatus("Disconnected from STT server -- reconnecting...", true);
            _client.ConnectionError    += (s, ex) => log.Debug("STT connect err: " + ex.Message);
            _client.StatusReceived     += (s, msg) => log.Debug("STT status: " + msg);
        }

        // ---------------------------------------------------------------
        public void Start()
        {
            lock (_lock)
            {
                if (_disposed) throw new ObjectDisposedException(nameof(SpeechListener));
                if (_isListening) return;

                try
                {
                    log.Info("SpeechListener v5: starting...");
                    RaiseStatus("Connecting to speech engine...");
                    _client.Start();
                    _mic.Start();
                    _isListening = true;
                    RaisePhaseChanged("Scan", null);
                    RaiseStatus("Listening for Bible references...");
                }
                catch (Exception ex)
                {
                    log.Error("SpeechListener v5: failed to start.", ex);
                    RaiseStatus("Failed to start: " + ex.Message, true);
                    SafeStop();
                    throw;
                }
            }
        }

        public void Stop()
        {
            lock (_lock)
            {
                if (!_isListening) return;
                log.Info("SpeechListener v5: stopping...");
                SafeStop();
                _isListening = false;
                RaiseStatus("Speech listening stopped.");
            }
        }

        public bool Toggle()
        {
            if (IsListening) { Stop(); return false; }
            Start(); return true;
        }

        // ---------------------------------------------------------------
        private void SafeStop()
        {
            try { _mic.Stop(); }                    catch (Exception ex) { log.Debug("mic stop: " + ex.Message); }
            try { _chunker.Flush(); }               catch { }
            try { _client.StopAsync().Wait(1000); } catch (Exception ex) { log.Debug("ws stop: " + ex.Message); }
            _phaseTimeout?.Dispose(); _phaseTimeout = null;
            _focusBook = null;
        }

        private async void OnChunkReady(object sender, byte[] pcm)
        {
            try { await _client.SendAudioAsync(pcm); }
            catch (Exception ex) { log.Debug("chunk send: " + ex.Message); }
        }

        // ---------------------------------------------------------------
        private void OnSegmentDropped(object sender, DroppedSegmentEventArgs e)
        {
            SegmentDropped?.Invoke(this, new SegmentDroppedEventArgs
            {
                Reason = e.Reason,
                Text = e.Text,
                DurationSeconds = e.Duration,
            });
        }

        private void OnTranscript(object sender, TranscriptEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(e.Text)) return;

            if (e.Confidence < MinEngineConfidence)
            {
                log.Debug($"SpeechListener: conf {e.Confidence:F2} < threshold {MinEngineConfidence:F2} -- drop.");
                return;
            }

            string corrected = PhoneticCorrector.Correct(e.Text);
            string normalised = Normalise(corrected);

            // Local dedup (belt + braces after server-side dedup).
            if (_lastEmittedNormalised == normalised &&
                (DateTime.UtcNow - _lastEmittedUtc) < _dedupWindow)
            {
                log.Debug($"SpeechListener: local dedup -- ignoring repeat \"{corrected}\".");
                return;
            }
            _lastEmittedNormalised = normalised;
            _lastEmittedUtc = DateTime.UtcNow;

            log.Debug($"SpeechListener v5: \"{e.Text}\" -> \"{corrected}\" " +
                      $"(conf={e.Confidence:F2} logprob={e.AvgLogProb:F2} " +
                      $"cr={e.CompressionRatio:F2} nsp={e.NoSpeechProb:F2} dur={e.DurationSeconds:F2}s)");

            UpdatePhaseFrom(corrected);

            SpeechRecognised?.Invoke(this, new SpeechRecognisedEventArgs
            {
                Text             = corrected,
                Confidence       = e.Confidence,
                AvgLogProb       = e.AvgLogProb,
                NoSpeechProb     = e.NoSpeechProb,
                CompressionRatio = e.CompressionRatio,
                DurationSeconds  = e.DurationSeconds,
            });
        }

        private void UpdatePhaseFrom(string text)
        {
            var det = BibleReferenceDetector.DetectBest(text);
            if (det != null && det.BookName != _focusBook)
            {
                _focusBook = det.BookName;
                RaisePhaseChanged("Focus", _focusBook);
                _phaseTimeout?.Dispose();
                _phaseTimeout = new System.Threading.Timer(_ =>
                {
                    _focusBook = null;
                    RaisePhaseChanged("Scan", null);
                }, null, FocusTimeoutMs, System.Threading.Timeout.Infinite);
            }
        }

        private static string Normalise(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            return System.Text.RegularExpressions.Regex
                .Replace(text.ToLowerInvariant(), @"[^a-z0-9 ]+", " ")
                .Trim();
        }

        private void RaiseStatus(string msg, bool isError = false) =>
            StatusChanged?.Invoke(this, new SpeechListenerStatusEventArgs
            { Message = msg, IsError = isError });

        private void RaisePhaseChanged(string phase, string book) =>
            PhaseChanged?.Invoke(this, new SpeechPhaseChangedEventArgs
            { Phase = phase, BookName = book });

        // ---------------------------------------------------------------
        public void Dispose()
        {
            lock (_lock)
            {
                if (_disposed) return;
                _disposed = true;
                SafeStop();
                _mic.Dispose();
                _client.Dispose();
            }
        }
    }

    /// <summary>
    /// Centralised list of default Whisper hotwords (canonical book names).
    /// Kept here so SpeechListener.Start() can push them on connect.
    /// </summary>
    internal static class BookHotwords
    {
        public static readonly string[] Default = new[]
        {
            "Genesis","Exodus","Leviticus","Numbers","Deuteronomy",
            "Joshua","Judges","Ruth","Samuel","Kings","Chronicles",
            "Ezra","Nehemiah","Esther","Job","Psalm","Psalms",
            "Proverbs","Ecclesiastes","Isaiah","Jeremiah",
            "Lamentations","Ezekiel","Daniel","Hosea","Joel","Amos",
            "Obadiah","Jonah","Micah","Nahum","Habakkuk","Zephaniah",
            "Haggai","Zechariah","Malachi",
            "Matthew","Mark","Luke","John","Acts","Romans",
            "Corinthians","Galatians","Ephesians","Philippians",
            "Colossians","Thessalonians","Timothy","Titus","Philemon",
            "Hebrews","James","Peter","Jude","Revelation",
            "chapter","verse",
        };
    }
}
