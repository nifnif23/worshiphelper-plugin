// ============================================================================
// SpeechListener.cs  --  v4 (Faster-Whisper edition)
//
// v3 ran Vosk in-process and juggled two recognition phases. v4 delegates all
// speech-to-text to a Python Faster-Whisper server and keeps the phase-state
// purely for UI (so the SpeechDebugPanel still shows Scan / Focus badges).
//
// Public API is preserved for backwards compatibility:
//   * event SpeechRecognised
//   * event StatusChanged
//   * event PhaseChanged          (Scan / Focus -- derived from detector output)
//   * Start(), Stop(), Toggle()
//   * IsListening, MinEngineConfidence
//
// Pipeline:
//   MicrophoneCapture --> Chunker --> PythonClient --> transcript
//                                                      |
//                              PhoneticCorrector <-----+
//                                   |
//                              SpeechRecognised  (consumers take over)
// ============================================================================
using System;
using System.IO;
using log4net;
using WorshipHelperVSTO.Audio;
using WorshipHelperVSTO.Networking;

namespace WorshipHelperVSTO
{
    // -- Event arg shapes kept identical to v3 --------------------------------
    public class SpeechRecognisedEventArgs : EventArgs
    {
        public string Text { get; set; }
        public float  Confidence { get; set; }
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

        // -- Config (kept for source-compat with the old listener) -----------
        public float MinEngineConfidence { get; set; } = 0.35f;

        /// <summary>Override if you run the server on another host/port.</summary>
        public Uri ServerUri
        {
            get => _client.ServerUri;
            set => _client.ServerUri = value;
        }

        // -- Events ----------------------------------------------------------
        public event EventHandler<SpeechRecognisedEventArgs>       SpeechRecognised;
        public event EventHandler<SpeechListenerStatusEventArgs>   StatusChanged;
        public event EventHandler<SpeechPhaseChangedEventArgs>     PhaseChanged;

        public bool IsListening { get { lock (_lock) return _isListening; } }

        // ---------------------------------------------------------------
        public SpeechListener()
        {
            _mic.PcmFrame      += (s, pcm) => _chunker.Feed(pcm);
            _mic.CaptureError  += (s, ex)  => RaiseStatus("Mic error: " + ex.Message, true);
            _chunker.ChunkReady += OnChunkReady;
            _client.TranscriptReceived += OnTranscript;
            _client.Connected          += (s, e) => RaiseStatus("Connected to STT server.");
            _client.Disconnected       += (s, e) => RaiseStatus("Disconnected -- reconnecting...", true);
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
                    log.Info("SpeechListener v4: starting...");
                    RaiseStatus("Connecting to speech engine...");
                    _client.Start();
                    _mic.Start();
                    _isListening = true;
                    RaisePhaseChanged("Scan", null);
                    RaiseStatus("Listening for Bible references...");
                }
                catch (Exception ex)
                {
                    log.Error("SpeechListener v4: failed to start.", ex);
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
                log.Info("SpeechListener v4: stopping...");
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
            try { _mic.Stop(); }                     catch (Exception ex) { log.Debug("mic stop: " + ex.Message); }
            try { _chunker.Flush(); }                catch { }
            try { _client.StopAsync().Wait(1000); }  catch (Exception ex) { log.Debug("ws stop: " + ex.Message); }
            _phaseTimeout?.Dispose(); _phaseTimeout = null;
            _focusBook = null;
        }

        private async void OnChunkReady(object sender, byte[] pcm)
        {
            try { await _client.SendAudioAsync(pcm); }
            catch (Exception ex) { log.Debug("chunk send: " + ex.Message); }
        }

        // ---------------------------------------------------------------
        private void OnTranscript(object sender, TranscriptEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(e.Text)) return;
            if (e.Confidence < MinEngineConfidence)
            {
                log.Debug($"SpeechListener: conf {e.Confidence:F2} < threshold {MinEngineConfidence:F2} -- drop.");
                return;
            }

            string corrected = PhoneticCorrector.Correct(e.Text);
            log.Debug($"SpeechListener v4 transcript: \"{e.Text}\" -> \"{corrected}\" (conf={e.Confidence:F2})");

            UpdatePhaseFrom(corrected);

            SpeechRecognised?.Invoke(this, new SpeechRecognisedEventArgs
            {
                Text = corrected,
                Confidence = e.Confidence,
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
}
