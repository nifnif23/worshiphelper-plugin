// ============================================================================
// Audio/MicrophoneCapture.cs
// Thin NAudio wrapper that produces 16-kHz mono int16 PCM frames.
// Decoupled from recognition so the rest of the pipeline is transport-agnostic.
// ============================================================================
using System;
using log4net;
using NAudio.Wave;

namespace WorshipHelperVSTO.Audio
{
    public sealed class MicrophoneCapture : IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(MicrophoneCapture));

        public const int SampleRate = 16_000;
        public const int Channels   = 1;

        private WaveInEvent _waveIn;
        private readonly object _lock = new object();
        private bool _running;
        private bool _disposed;

        /// <summary>Fires for every mic buffer -- int16 PCM @ 16 kHz mono.</summary>
        public event EventHandler<byte[]> PcmFrame;

        /// <summary>Fires if the underlying capture device raises an error.</summary>
        public event EventHandler<Exception> CaptureError;

        public bool IsRunning { get { lock (_lock) return _running; } }

        public void Start()
        {
            lock (_lock)
            {
                if (_disposed) throw new ObjectDisposedException(nameof(MicrophoneCapture));
                if (_running) return;

                _waveIn = new WaveInEvent
                {
                    WaveFormat = new WaveFormat(SampleRate, 16, Channels),
                    BufferMilliseconds = 100,
                    NumberOfBuffers = 4,
                };
                _waveIn.DataAvailable    += OnData;
                _waveIn.RecordingStopped += OnStopped;
                _waveIn.StartRecording();
                _running = true;
                log.Info("MicrophoneCapture: started (16 kHz mono int16).");
            }
        }

        public void Stop()
        {
            lock (_lock)
            {
                if (!_running) return;
                try { _waveIn?.StopRecording(); }
                catch (Exception ex) { log.Warn("Error stopping WaveIn: " + ex.Message); }
                _running = false;
            }
        }

        private void OnData(object sender, WaveInEventArgs e)
        {
            if (e.BytesRecorded <= 0) return;
            var buf = new byte[e.BytesRecorded];
            Buffer.BlockCopy(e.Buffer, 0, buf, 0, e.BytesRecorded);
            PcmFrame?.Invoke(this, buf);
        }

        private void OnStopped(object sender, StoppedEventArgs e)
        {
            if (e.Exception != null)
            {
                log.Warn("MicrophoneCapture: recording stopped with error: " + e.Exception.Message);
                CaptureError?.Invoke(this, e.Exception);
            }
        }

        public void Dispose()
        {
            lock (_lock)
            {
                if (_disposed) return;
                _disposed = true;
                try { _waveIn?.StopRecording(); } catch { /* best-effort */ }
                _waveIn?.Dispose();
                _waveIn = null;
                _running = false;
            }
        }
    }
}
