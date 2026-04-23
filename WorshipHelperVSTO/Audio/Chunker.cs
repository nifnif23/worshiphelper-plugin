// ============================================================================
// Audio/Chunker.cs  --  v5
//
// v4 problem: RMS silence threshold of 0.003 was essentially "any room tone
// passes". Silent chunks streamed to the server all day long, and because the
// v4 server ran Whisper on every chunk, that directly caused the hallucination
// loop the user was seeing.
//
// v5 behaviour:
//   * Chunker is now a THIN passthrough with energy + voiced-ratio gating.
//     It does NOT try to detect utterance boundaries -- the Python server's
//     Silero VAD does that. Our job here is just to avoid wasting bandwidth
//     on obvious silence.
//   * Target chunk size dropped to 200ms so the server's VAD has fresh
//     audio to reason about. The server aggregates into utterances.
//   * Two-stage gate:
//        (1) RMS >= MinRms  (absolute energy floor, ~ -42 dBFS)
//        (2) VoicedFrameRatio >= MinVoicedRatio   (reject steady hums)
//     We estimate "voiced frame" cheaply with a short-time RMS over 20ms
//     sub-frames. Doesn't try to be a real VAD; just kills noise.
// ============================================================================
using System;
using System.IO;
using log4net;

namespace WorshipHelperVSTO.Audio
{
    public sealed class Chunker
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Chunker));

        /// <summary>Target outbound chunk size in milliseconds.</summary>
        public int TargetChunkMs { get; set; } = 200;

        /// <summary>Absolute RMS floor (0..1). Below this = silence.</summary>
        public double MinRms { get; set; } = 0.010;

        /// <summary>Fraction of 20-ms sub-frames that must be above MinRms.</summary>
        public double MinVoicedRatio { get; set; } = 0.15;

        private readonly MemoryStream _buffer = new MemoryStream();
        private readonly object _lock = new object();

        /// <summary>Counts for the debug panel.</summary>
        public long TotalChunks   { get; private set; }
        public long SilentChunks  { get; private set; }
        public double LastRms     { get; private set; }
        public double LastVoicedRatio { get; private set; }

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

            TotalChunks++;

            var stats = AnalyseChunk(toEmit);
            LastRms         = stats.rms;
            LastVoicedRatio = stats.voicedRatio;

            if (stats.rms < MinRms || stats.voicedRatio < MinVoicedRatio)
            {
                SilentChunks++;
                // Emit a rare trace once every ~5s so we know the gate is
                // alive without spamming the log.
                if ((SilentChunks % 25) == 0)
                    log.Debug($"Chunker: silence gate dropped {SilentChunks:N0} chunks so far " +
                              $"(last rms={stats.rms:F4}, voiced={stats.voicedRatio:P0}).");
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
            if (toEmit == null) return;

            var stats = AnalyseChunk(toEmit);
            if (stats.rms >= MinRms && stats.voicedRatio >= MinVoicedRatio)
                ChunkReady?.Invoke(this, toEmit);
        }

        // --------------------------------------------------------------
        private (double rms, double voicedRatio) AnalyseChunk(byte[] pcm)
        {
            int n = pcm.Length / 2;
            if (n == 0) return (0, 0);

            // Full-chunk RMS
            long sumSq = 0;
            for (int i = 0; i < pcm.Length; i += 2)
            {
                short s = (short)(pcm[i] | (pcm[i + 1] << 8));
                sumSq += s * s;
            }
            double rms = Math.Sqrt(sumSq / (double)n) / 32768.0;

            // Sub-frame voiced ratio (20ms sub-frames @ 16kHz = 320 samples = 640 bytes)
            int subBytes = MicrophoneCapture.SampleRate * 2 * 20 / 1000;
            int total = 0, voiced = 0;
            for (int off = 0; off + subBytes <= pcm.Length; off += subBytes)
            {
                long ss = 0;
                int sn = subBytes / 2;
                for (int i = off; i < off + subBytes; i += 2)
                {
                    short s = (short)(pcm[i] | (pcm[i + 1] << 8));
                    ss += s * s;
                }
                double sr = Math.Sqrt(ss / (double)sn) / 32768.0;
                total++;
                if (sr >= MinRms) voiced++;
            }
            double voicedRatio = total == 0 ? 0.0 : (double)voiced / total;
            return (rms, voicedRatio);
        }
    }
}
