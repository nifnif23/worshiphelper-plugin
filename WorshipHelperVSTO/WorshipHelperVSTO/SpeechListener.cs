// ============================================================================
// SpeechListener.cs
// Windows speech-to-text listener using System.Speech (built-in, free, offline).
//
// Wraps SpeechRecognitionEngine in an easy start/stop interface with events
// for recognised phrases.
//
// Prerequisites:
//   - Reference: System.Speech (GAC assembly, no NuGet needed)
//   - Target: .NET Framework 4.7.2+
//   - Windows OS with speech recognition support
//
// Drop into:  WorshipHelperVSTO/SpeechListener.cs
// Namespace:  WorshipHelperVSTO
// ============================================================================

using System;
using System.Globalization;
using System.Speech.Recognition;
using log4net;

namespace WorshipHelperVSTO
{
    // -----------------------------------------------------------------------
    // Event args
    // -----------------------------------------------------------------------

    /// <summary>
    /// Carries raw speech recognition results to subscribers.
    /// </summary>
    public class SpeechRecognisedEventArgs : EventArgs
    {
        /// <summary>
        /// The raw text recognised by the speech engine.
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Engine confidence (0.0–1.0) in the recognition result.
        /// </summary>
        public float Confidence { get; set; }
    }

    /// <summary>
    /// Carries status/error information from the speech listener.
    /// </summary>
    public class SpeechListenerStatusEventArgs : EventArgs
    {
        public string Message { get; set; }
        public bool IsError { get; set; }
    }

    // -----------------------------------------------------------------------
    // SpeechListener
    // -----------------------------------------------------------------------

    /// <summary>
    /// Provides a managed wrapper around the Windows built-in speech recognition
    /// engine (System.Speech.Recognition.SpeechRecognitionEngine).
    ///
    /// Features:
    ///   - Fully offline / free — uses the Windows Desktop speech engine.
    ///   - Start / Stop / IsListening API.
    ///   - Continuous dictation mode (free-form speech, not fixed grammar).
    ///   - Fires SpeechRecognised event for every recognised phrase.
    ///   - Configurable minimum confidence threshold to filter low-quality results.
    ///   - Thread-safe disposal.
    ///
    /// Usage:
    ///   var listener = new SpeechListener();
    ///   listener.SpeechRecognised += (s, e) => Console.WriteLine(e.Text);
    ///   listener.Start();
    ///   // … later …
    ///   listener.Stop();
    ///   listener.Dispose();
    /// </summary>
    public class SpeechListener : IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(SpeechListener));

        private SpeechRecognitionEngine _engine;
        private bool _isListening;
        private bool _disposed;
        private readonly object _lock = new object();

        // -----------------------------------------------------------------------
        // Configuration
        // -----------------------------------------------------------------------

        /// <summary>
        /// Minimum engine confidence required before the SpeechRecognised event fires.
        /// Range: 0.0 – 1.0.  Default: 0.3 (fairly permissive — the Bible reference
        /// detector adds its own confidence filter on top).
        ///
        /// Raise this value if the environment is very noisy and you're getting
        /// too many false recognitions.  Lower it in quiet settings.
        /// </summary>
        public float MinEngineConfidence { get; set; } = 0.3f;

        /// <summary>
        /// The culture/language for speech recognition.
        /// Default: en-US.  Change to en-GB, en-AU, etc. as needed.
        /// The corresponding Windows language pack must be installed.
        /// </summary>
        public CultureInfo Culture { get; set; } = new CultureInfo("en-US");

        // -----------------------------------------------------------------------
        // Events
        // -----------------------------------------------------------------------

        /// <summary>
        /// Fired every time the engine recognises a spoken phrase above the
        /// MinEngineConfidence threshold.
        /// </summary>
        public event EventHandler<SpeechRecognisedEventArgs> SpeechRecognised;

        /// <summary>
        /// Fired when the listener encounters an error or notable status change
        /// (e.g., audio source problems, engine restart).
        /// </summary>
        public event EventHandler<SpeechListenerStatusEventArgs> StatusChanged;

        // -----------------------------------------------------------------------
        // Properties
        // -----------------------------------------------------------------------

        /// <summary>
        /// Returns true if the listener is currently active and processing audio.
        /// </summary>
        public bool IsListening
        {
            get { lock (_lock) return _isListening; }
        }

        // -----------------------------------------------------------------------
        // Start / Stop
        // -----------------------------------------------------------------------

        /// <summary>
        /// Initialises the speech recognition engine and starts listening
        /// to the default audio input device (microphone).
        /// Safe to call multiple times — subsequent calls are no-ops if already listening.
        /// </summary>
        public void Start()
        {
            lock (_lock)
            {
                if (_disposed) throw new ObjectDisposedException(nameof(SpeechListener));
                if (_isListening) return;

                try
                {
                    log.Info("SpeechListener: Initialising speech recognition engine…");
                    RaiseStatus("Initialising speech recognition…");

                    _engine = new SpeechRecognitionEngine(Culture);

                    // Use free-form dictation grammar — we don't want a fixed grammar
                    // because the speaker may say anything.  The Bible reference detector
                    // layer handles the filtering.
                    _engine.LoadGrammar(new DictationGrammar());

                    // Wire up events
                    _engine.SpeechRecognized += OnSpeechRecognized;
                    _engine.SpeechRecognitionRejected += OnSpeechRejected;
                    _engine.RecognizeCompleted += OnRecognizeCompleted;
                    _engine.AudioStateChanged += OnAudioStateChanged;

                    // Use the default system microphone
                    _engine.SetInputToDefaultAudioDevice();

                    // Start continuous recognition (asynchronous, non-blocking)
                    _engine.RecognizeAsync(RecognizeMode.Multiple);

                    _isListening = true;

                    log.Info("SpeechListener: Now listening.");
                    RaiseStatus("Listening for Bible references…");
                }
                catch (Exception ex)
                {
                    log.Error("SpeechListener: Failed to start speech recognition.", ex);
                    RaiseStatus($"Failed to start: {ex.Message}", isError: true);
                    CleanupEngine();
                    throw;
                }
            }
        }

        /// <summary>
        /// Stops listening and releases the audio device.
        /// Safe to call multiple times — subsequent calls are no-ops.
        /// The listener can be restarted after stopping.
        /// </summary>
        public void Stop()
        {
            lock (_lock)
            {
                if (!_isListening) return;

                try
                {
                    log.Info("SpeechListener: Stopping…");
                    _engine?.RecognizeAsyncCancel();
                }
                catch (Exception ex)
                {
                    log.Warn("SpeechListener: Error during RecognizeAsyncCancel.", ex);
                }

                CleanupEngine();
                _isListening = false;

                log.Info("SpeechListener: Stopped.");
                RaiseStatus("Speech listening stopped.");
            }
        }

        /// <summary>
        /// Toggles between listening and not listening.
        /// Returns the new state (true = listening, false = stopped).
        /// </summary>
        public bool Toggle()
        {
            if (IsListening)
            {
                Stop();
                return false;
            }
            else
            {
                Start();
                return true;
            }
        }

        // -----------------------------------------------------------------------
        // Engine event handlers
        // -----------------------------------------------------------------------

        private void OnSpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result == null) return;

            string text = e.Result.Text;
            float confidence = e.Result.Confidence;

            log.Debug($"SpeechListener: Recognised \"{text}\" (confidence={confidence:F2})");

            if (confidence < MinEngineConfidence)
            {
                log.Debug($"SpeechListener: Below MinEngineConfidence ({MinEngineConfidence:F2}), ignoring.");
                return;
            }

            // Fire the event on the current thread (engine's internal thread).
            // Subscribers should marshal to UI thread if needed.
            SpeechRecognised?.Invoke(this, new SpeechRecognisedEventArgs
            {
                Text = text,
                Confidence = confidence
            });
        }

        private void OnSpeechRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            // This fires when the engine hears something but cannot match it to any grammar.
            // With DictationGrammar this is rare, but can happen in very noisy environments.
            log.Debug("SpeechListener: Speech rejected (could not recognise).");
        }

        private void OnRecognizeCompleted(object sender, RecognizeCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                log.Warn($"SpeechListener: Recognition completed with error: {e.Error.Message}");
                RaiseStatus($"Recognition error: {e.Error.Message}", isError: true);
            }

            if (e.Cancelled)
            {
                log.Debug("SpeechListener: Recognition was cancelled.");
            }
        }

        private void OnAudioStateChanged(object sender, AudioStateChangedEventArgs e)
        {
            log.Debug($"SpeechListener: Audio state → {e.AudioState}");
        }

        // -----------------------------------------------------------------------
        // Cleanup
        // -----------------------------------------------------------------------

        private void CleanupEngine()
        {
            if (_engine != null)
            {
                try
                {
                    _engine.SpeechRecognized -= OnSpeechRecognized;
                    _engine.SpeechRecognitionRejected -= OnSpeechRejected;
                    _engine.RecognizeCompleted -= OnRecognizeCompleted;
                    _engine.AudioStateChanged -= OnAudioStateChanged;
                    _engine.Dispose();
                }
                catch (Exception ex)
                {
                    log.Warn("SpeechListener: Error during engine cleanup.", ex);
                }
                _engine = null;
            }
        }

        // -----------------------------------------------------------------------
        // IDisposable
        // -----------------------------------------------------------------------

        public void Dispose()
        {
            lock (_lock)
            {
                if (_disposed) return;
                _disposed = true;
                Stop();
            }
        }

        // -----------------------------------------------------------------------
        // Helpers
        // -----------------------------------------------------------------------

        private void RaiseStatus(string message, bool isError = false)
        {
            StatusChanged?.Invoke(this, new SpeechListenerStatusEventArgs
            {
                Message = message,
                IsError = isError
            });
        }
    }
}
