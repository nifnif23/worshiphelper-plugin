// ============================================================================
// Networking/PythonClient.cs
// WebSocket client for the Faster-Whisper Python sidecar.
//
// Responsibilities:
//   * Connect to ws://127.0.0.1:8765/stt
//   * Stream binary PCM chunks as they arrive from Chunker
//   * Parse transcript messages and raise TranscriptReceived
//   * Auto-reconnect with capped exponential backoff
//   * Graceful shutdown (Close handshake + cancellation)
//
// Uses System.Net.WebSockets.ClientWebSocket -- available in .NET Framework
// 4.7.2 out of the box, no extra NuGet dependencies.
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
        public float Confidence { get; set; }
        public bool IsFinal { get; set; }
        public double DurationSeconds { get; set; }
    }

    public sealed class PythonClient : IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(PythonClient));

        public Uri ServerUri { get; set; } = new Uri("ws://127.0.0.1:8765/stt");
        public TimeSpan InitialRetryDelay { get; set; } = TimeSpan.FromSeconds(1);
        public TimeSpan MaxRetryDelay     { get; set; } = TimeSpan.FromSeconds(15);

        public event EventHandler<TranscriptEventArgs> TranscriptReceived;
        public event EventHandler<string> StatusReceived;
        public event EventHandler<Exception> ConnectionError;
        public event EventHandler Connected;
        public event EventHandler Disconnected;

        private ClientWebSocket _ws;
        private CancellationTokenSource _cts;
        private Task _rxTask;
        private readonly SemaphoreSlim _sendLock = new SemaphoreSlim(1, 1);
        private bool _disposed;

        public bool IsConnected =>
            _ws != null && _ws.State == WebSocketState.Open;

        // ---------------------------------------------------------------
        public void Start()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(PythonClient));
            if (_cts != null) return;

            _cts = new CancellationTokenSource();
            _rxTask = Task.Run(() => RunAsync(_cts.Token));
        }

        public async Task StopAsync()
        {
            if (_cts == null) return;
            _cts.Cancel();
            try { await _rxTask.ConfigureAwait(false); }
            catch { /* swallow */ }
            _cts.Dispose(); _cts = null; _rxTask = null;
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
                log.Warn("PythonClient: send failed: " + ex.Message);
            }
            finally { _sendLock.Release(); }
        }

        public Task SendFlushAsync() => SendControlAsync("{\"type\":\"flush\"}");
        public Task SendResetAsync() => SendControlAsync("{\"type\":\"reset\"}");

        private async Task SendControlAsync(string json)
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
            catch (Exception ex) { log.Warn("PythonClient: control send failed: " + ex.Message); }
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
                    log.Warn("PythonClient: connection error: " + ex.Message);
                    ConnectionError?.Invoke(this, ex);
                }
                finally
                {
                    try { _ws?.Dispose(); } catch { }
                    _ws = null;
                    Disconnected?.Invoke(this, EventArgs.Empty);
                }

                if (token.IsCancellationRequested) break;
                log.Info($"PythonClient: reconnecting in {delay.TotalSeconds:F0}s...");
                try { await Task.Delay(delay, token).ConfigureAwait(false); }
                catch { break; }
                delay = TimeSpan.FromSeconds(Math.Min(
                    MaxRetryDelay.TotalSeconds, delay.TotalSeconds * 2));
            }
        }

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
                catch (Exception ex) { log.Warn("PythonClient: bad msg: " + ex.Message); }
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
                        Text = (string)obj["text"] ?? "",
                        Confidence = (float?)obj["confidence"] ?? 0.0f,
                        IsFinal = (bool?)obj["final"] ?? true,
                        DurationSeconds = (double?)obj["duration"] ?? 0,
                    });
                    break;
                case "status":
                    StatusReceived?.Invoke(this, (string)obj["message"] ?? "");
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
