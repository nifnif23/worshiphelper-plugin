// ============================================================================
// Audio/Chunker.cs
// Aggregates short NAudio frames (~100 ms) into larger chunks before sending
// them to the Python server. This reduces WebSocket overhead and lets the
// server window transcription naturally.
//
// Also provides simple voice-activity gating: chunks that are pure silence
// are dropped so we don't ship quiet audio over the wire. Faster-Whisper
// still applies its own VAD, but gating here saves bandwidth & CPU.
// ============================================================================
using System;
using System.IO;
using log4net;

namespace WorshipHelperVSTO.Audio
{
    public sealed class Chunker
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Chunker));

        public int TargetChunkMs { get; set; } = 500;
        /// <summary>RMS threshold (0..1). Below = assumed silence, skipped.</summary>
        public double SilenceRmsThreshold { get; set; } = 0.003;

        private readonly MemoryStream _buffer = new MemoryStream();
        private readonly object _lock = new object();

        public event EventHandler<byte[]> ChunkReady;

        public void Feed(byte[] pcmInt16)
        {
            if (pcmInt16 == null || pcmInt16.Length == 0) return;

            byte[] toEmit = null;
            lock (_lock)
            {
                _buffer.Write(pcmInt16, 0, pcmInt16.Length);
                int targetBytes =
                    MicrophoneCapture.SampleRate * 2 * TargetChunkMs / 1000;
                if (_buffer.Length >= targetBytes)
                {
                    toEmit = _buffer.ToArray();
                    _buffer.SetLength(0);
                }
            }
            if (toEmit == null) return;

            if (IsSilence(toEmit))
            {
                log.Debug("Chunker: skipping silent chunk.");
                return;
            }

            ChunkReady?.Invoke(this, toEmit);
        }

        public void Flush()
        {
            byte[] toEmit = null;
            lock (_lock)
            {
                if (_buffer.Length > 0)
                {
                    toEmit = _buffer.ToArray();
                    _buffer.SetLength(0);
                }
            }
            if (toEmit != null && !IsSilence(toEmit))
                ChunkReady?.Invoke(this, toEmit);
        }

        private bool IsSilence(byte[] pcm)
        {
            long sum = 0; int n = pcm.Length / 2;
            if (n == 0) return true;
            for (int i = 0; i < pcm.Length; i += 2)
            {
                short s = (short)(pcm[i] | (pcm[i + 1] << 8));
                sum += s * s;
            }
            double rms = Math.Sqrt(sum / (double)n) / 32768.0;
            return rms < SilenceRmsThreshold;
        }
    }
}
