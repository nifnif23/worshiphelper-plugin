// ============================================================================
// Networking/PythonClient.cs  --  v5
//
// New in v5:
//   * Parses the richer transcript payload (avg_logprob, compression_ratio,
//     no_speech_prob, duration) so the C# detection pipeline can use them
//     as additional trust signals.
//   * "dropped" message type -- the server tells us when it guard-railed a
//     segment, so the debug panel can show "mic is live, but Whisper
//     threw away: Thanks for watching".
//   * SendHotwordsAsync() -- optional per-session bias (book names).
//   * Heartbeat: sends ping every 15s, surfaces latency for diagnostics.
// ============================================================================
using System;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using log4net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WorshipHelperVSTO.Networking
{
    public sealed class TranscriptEventArgs : EventArgs
    {
        public string Text { get; set; }
        public float  Confidence { get; set; }
        public bool   IsFinal { get; set; }
        public double DurationSeconds { get; set; }
        public double AvgLogProb { get; set; }
        public double CompressionRatio { get; set; }
        public double NoSpeechProb { get; set; }
    }

    public sealed class DroppedSegmentEventArgs : EventArgs
    {
        public string Reason { get; set; }
        public string Text { get; set; }
        public double Duration { get; set; }
        public double PeakRms { get; set; }
    }

    public sealed class PythonClient : IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(PythonClient));

        public Uri ServerUri { get; set; } = new Uri("ws://127.0.0.1:8765/");
        public TimeSpan InitialRetryDelay { get; set; } = TimeSpan.FromSeconds(1);
        public TimeSpan MaxRetryDelay     { get; set; } = TimeSpan.FromSeconds(15);
        public TimeSpan PingInterval      { get; set; } = TimeSpan.FromSeconds(15);

        public event EventHandler<TranscriptEventArgs>      TranscriptReceived;
        public event EventHandler<DroppedSegmentEventArgs>  SegmentDropped;
        public event EventHandler<string>                   StatusReceived;
        public event EventHandler<Exception>                ConnectionError;
        public event EventHandler                           Connected;
        public event EventHandler                           Disconnected;

        private ClientWebSocket _ws;
        private CancellationTokenSource _cts;
        private Task _rxTask;
        private Task _pingTask;
        private readonly SemaphoreSlim _sendLock = new SemaphoreSlim(1, 1);
        private bool _disposed;

        public bool IsConnected =>
            _ws != null && _ws.State == WebSocketState.Open;

        /// <summary>Last round-trip latency in ms (from ping/pong). 0 if unknown.</summary>
        public int LastPingMs { get; private set; }

        // ---------------------------------------------------------------
        public void Start()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(PythonClient));
            if (_cts != null) return;

            _cts = new CancellationTokenSource();
            _rxTask = Task.Run(() => RunAsync(_cts.Token));
            _pingTask = Task.Run(() => PingLoopAsync(_cts.Token));
        }

        public async Task StopAsync()
        {
            if (_cts == null) return;
            _cts.Cancel();
            try { await Task.WhenAll(_rxTask ?? Task.CompletedTask,
                                     _pingTask ?? Task.CompletedTask)
                                .ConfigureAwait(false); }
            catch { /* swallow */ }
            _cts.Dispose(); _cts = null; _rxTask = null; _pingTask = null;
        }

        // ---------------------------------------------------------------
        public async Task SendAudioAsync(byte[] pcm)
        {
            if (!IsConnected || pcm == null || pcm.Length == 0) return;
            await _sendLock.WaitAsync().ConfigureAwait(false);
            try
            {
                await _ws.SendAsync(
                    new ArraySegment<byte>(pcm),
                    WebSocketMessageType.Binary,
                    endOfMessage: true,
                    cancellationToken: _cts.Token).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                log.Debug("PythonClient: send failed: " + ex.Message);
            }
            finally { _sendLock.Release(); }
        }

        public Task SendFlushAsync() => SendJsonAsync("{\"type\":\"flush\"}");
        public Task SendResetAsync() => SendJsonAsync("{\"type\":\"reset\"}");

        public async Task SendHotwordsAsync(string[] words)
        {
            if (words == null || words.Length == 0) return;
            var payload = JsonConvert.SerializeObject(new { type = "hotwords", words });
            await SendJsonAsync(payload).ConfigureAwait(false);
        }

        private async Task SendJsonAsync(string json)
        {
            if (!IsConnected) return;
            await _sendLock.WaitAsync().ConfigureAwait(false);
            try
            {
                var bytes = Encoding.UTF8.GetBytes(json);
                await _ws.SendAsync(
                    new ArraySegment<byte>(bytes),
                    WebSocketMessageType.Text, true, _cts.Token).ConfigureAwait(false);
            }
            catch (Exception ex) { log.Debug("PythonClient: json send failed: " + ex.Message); }
            finally { _sendLock.Release(); }
        }

        // ---------------------------------------------------------------
        private async Task RunAsync(CancellationToken token)
        {
            TimeSpan delay = InitialRetryDelay;
            while (!token.IsCancellationRequested)
            {
                try
                {
                    _ws = new ClientWebSocket();
                    _ws.Options.KeepAliveInterval = TimeSpan.FromSeconds(20);
                    log.Info("PythonClient: connecting to " + ServerUri);
                    await _ws.ConnectAsync(ServerUri, token).ConfigureAwait(false);
                    log.Info("PythonClient: connected.");
                    Connected?.Invoke(this, EventArgs.Empty);
                    delay = InitialRetryDelay;

                    await ReceiveLoopAsync(token).ConfigureAwait(false);
                }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    log.Debug("PythonClient: connection error: " + ex.Message);
                    ConnectionError?.Invoke(this, ex);
                }
                finally
                {
                    try { _ws?.Dispose(); } catch { }
                    _ws = null;
                    Disconnected?.Invoke(this, EventArgs.Empty);
                }

                if (token.IsCancellationRequested) break;
                log.Debug($"PythonClient: reconnecting in {delay.TotalSeconds:F0}s...");
                try { await Task.Delay(delay, token).ConfigureAwait(false); }
                catch { break; }
                delay = TimeSpan.FromSeconds(Math.Min(
                    MaxRetryDelay.TotalSeconds, delay.TotalSeconds * 2));
            }
        }

        private async Task PingLoopAsync(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    await Task.Delay(PingInterval, token).ConfigureAwait(false);
                    if (IsConnected)
                    {
                        _pingSentUtc = DateTime.UtcNow;
                        await SendJsonAsync("{\"type\":\"ping\"}").ConfigureAwait(false);
                    }
                }
                catch (OperationCanceledException) { break; }
                catch (Exception ex) { log.Debug("ping loop: " + ex.Message); }
            }
        }
        private DateTime _pingSentUtc;

        private async Task ReceiveLoopAsync(CancellationToken token)
        {
            var buffer = new byte[16 * 1024];
            var sb = new StringBuilder();

            while (_ws != null && _ws.State == WebSocketState.Open && !token.IsCancellationRequested)
            {
                WebSocketReceiveResult result;
                try
                {
                    result = await _ws.ReceiveAsync(
                        new ArraySegment<byte>(buffer), token).ConfigureAwait(false);
                }
                catch (OperationCanceledException) { break; }

                if (result.MessageType == WebSocketMessageType.Close)
                {
                    await _ws.CloseAsync(WebSocketCloseStatus.NormalClosure,
                        "bye", CancellationToken.None).ConfigureAwait(false);
                    break;
                }
                if (result.MessageType != WebSocketMessageType.Text) continue;

                sb.Append(Encoding.UTF8.GetString(buffer, 0, result.Count));
                if (!result.EndOfMessage) continue;

                string json = sb.ToString();
                sb.Clear();
                try { HandleServerMessage(json); }
                catch (Exception ex) { log.Debug("PythonClient: bad msg: " + ex.Message); }
            }
        }

        private void HandleServerMessage(string json)
        {
            var obj = JObject.Parse(json);
            string type = (string)obj["type"];
            switch (type)
            {
                case "transcript":
                    TranscriptReceived?.Invoke(this, new TranscriptEventArgs
                    {
                        Text             = (string)obj["text"] ?? "",
                        Confidence       = (float?) obj["confidence"] ?? 0.0f,
                        IsFinal          = (bool?)  obj["final"] ?? true,
                        DurationSeconds  = (double?)obj["duration"] ?? 0,
                        AvgLogProb       = (double?)obj["avg_logprob"] ?? 0,
                        CompressionRatio = (double?)obj["compression_ratio"] ?? 0,
                        NoSpeechProb     = (double?)obj["no_speech_prob"] ?? 0,
                    });
                    break;

                case "dropped":
                    SegmentDropped?.Invoke(this, new DroppedSegmentEventArgs
                    {
                        Reason   = (string)obj["reason"] ?? "",
                        Text     = (string)obj["text"] ?? "",
                        Duration = (double?)obj["duration"] ?? 0,
                        PeakRms  = (double?)obj["peak_rms"] ?? 0,
                    });
                    break;

                case "status":
                    StatusReceived?.Invoke(this, (string)obj["message"] ?? "");
                    break;

                case "pong":
                    if (_pingSentUtc != default)
                        LastPingMs = (int)(DateTime.UtcNow - _pingSentUtc).TotalMilliseconds;
                    break;

                case "error":
                    log.Warn("Server error: " + (string)obj["message"]);
                    break;
            }
        }

        // ---------------------------------------------------------------
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { StopAsync().Wait(TimeSpan.FromSeconds(2)); } catch { }
            _sendLock.Dispose();
        }
    }
}
