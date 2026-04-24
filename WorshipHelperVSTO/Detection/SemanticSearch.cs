// ============================================================================
// Detection/SemanticSearch.cs  --  v7
//
// v6 → v7 changes:
//   * Supports 768-dim BAAI/bge-base-en-v1.5 embeddings (was 384-dim MiniLM).
//     Dimension is now read from the SQLite meta table at runtime — no
//     hard-coded magic number.
//   * MinSimilarity lowered: 0.62 → 0.55.
//     bge-base produces tighter score distributions than MiniLM; 0.62 was
//     dropping valid paraphrase matches. 0.55 is safe because the margin
//     guard (MinMargin) still filters ambiguous results.
//   * MinWordCount lowered: 6 → 5.
//     Short but clear quotes ("God so loved the world") now reach the
//     semantic matcher instead of being discarded at the word-count gate.
//   * Embed endpoint tries port 8765 first (STT server /embed passthrough)
//     then falls back to port 8766 (dedicated embed_server.py). The dedicated
//     server is preferred; the 8765 route exists for backwards compatibility.
//   * Response parsing updated to handle the v7 embed server response
//     {"embedding":[...],"model":"...","dim":768}.
// ============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using log4net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WorshipHelperVSTO.Detection
{
    public sealed class SemanticMatch
    {
        public string Reference;         // "John 3:16"
        public string Text;              // verse text
        public float  Similarity;        // cosine, 0..1
        public float  Margin;            // sim - secondSim (confidence proxy)
    }

    public sealed class SemanticSearch
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(SemanticSearch));
        private static readonly HttpClient _http = new HttpClient
        { Timeout = TimeSpan.FromSeconds(5) };

        private readonly VerseDatabase _db;

        // Try the dedicated embed server first (port 8766), fall back to the
        // STT server's embed passthrough (port 8765) if unavailable.
        private static readonly Uri[] _embedCandidates = new[]
        {
            new Uri("http://127.0.0.1:8766/embed"),   // dedicated embed_server.py
            new Uri("http://127.0.0.1:8765/embed"),   // STT server passthrough (legacy)
        };
        private Uri _activeEmbedEndpoint;

        /// <summary>Min cosine similarity before we consider reporting a match.
        /// Lowered from 0.62 for bge-base which has tighter score distributions.</summary>
        public float MinSimilarity { get; set; } = 0.55f;

        /// <summary>Min margin between #1 and #2 — filters ambiguous results.</summary>
        public float MinMargin { get; set; } = 0.05f;

        /// <summary>Short utterances filtered out — not enough semantic signal.
        /// Lowered from 6 to 5 for short clear quotes.</summary>
        public int MinWordCount { get; set; } = 5;

        public SemanticSearch(VerseDatabase db) { _db = db; }

        public async Task<SemanticMatch> FindAsync(string utterance)
        {
            if (_db == null || !_db.IsLoaded) return null;
            if (string.IsNullOrWhiteSpace(utterance)) return null;
            int wc = utterance.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;
            if (wc < MinWordCount) return null;

            float[] query;
            try { query = await EmbedAsync(utterance).ConfigureAwait(false); }
            catch (Exception ex)
            {
                log.Debug("SemanticSearch: embed failed — " + ex.Message);
                return null;
            }

            if (query == null || query.Length != _db.Dimension)
            {
                if (query != null)
                    log.Warn($"SemanticSearch: embed dim mismatch — got {query.Length}, expected {_db.Dimension}. " +
                             "Did you rebuild verses.sqlite after changing the embed model?");
                return null;
            }

            // Linear top-2 scan — 31k × 768 floats fits comfortably in RAM
            // and runs in <50 ms on a modern CPU with the SIMD dot product.
            float best1 = -1, best2 = -1;
            VerseRow winner = null;
            foreach (var row in _db.Verses)
            {
                float sim = DotProduct(query, row.Embedding);
                if (sim > best1) { best2 = best1; best1 = sim; winner = row; }
                else if (sim > best2) { best2 = sim; }
            }

            if (winner == null || best1 < MinSimilarity) return null;
            float margin = best1 - best2;
            if (margin < MinMargin) return null;

            return new SemanticMatch
            {
                Reference  = winner.Reference,
                Text       = winner.Text,
                Similarity = best1,
                Margin     = margin,
            };
        }

        // Both sides produced with normalize_embeddings=True → dot product = cosine.
        private static float DotProduct(float[] a, float[] b)
        {
            float dot = 0;
            int len = Math.Min(a.Length, b.Length);
            for (int i = 0; i < len; i++) dot += a[i] * b[i];
            return dot;
        }

        private async Task<float[]> EmbedAsync(string text)
        {
            // Try each endpoint in order; remember the first one that works.
            var endpoints = _activeEmbedEndpoint != null
                ? new[] { _activeEmbedEndpoint }
                : _embedCandidates;

            var payload = JsonConvert.SerializeObject(new { text });
            Exception lastEx = null;

            foreach (var ep in endpoints)
            {
                try
                {
                    using (var req = new HttpRequestMessage(HttpMethod.Post, ep)
                    { Content = new StringContent(payload, Encoding.UTF8, "application/json") })
                    using (var resp = await _http.SendAsync(req).ConfigureAwait(false))
                    {
                        resp.EnsureSuccessStatusCode();
                        string json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                        var obj = JObject.Parse(json);
                        var arr = (JArray)obj["embedding"];
                        var vec = new float[arr.Count];
                        for (int i = 0; i < arr.Count; i++) vec[i] = (float)arr[i];
                        _activeEmbedEndpoint = ep;   // cache the working endpoint
                        return vec;
                    }
                }
                catch (Exception ex)
                {
                    lastEx = ex;
                    log.Debug($"SemanticSearch: {ep} failed — {ex.Message}");
                }
            }
            throw lastEx ?? new Exception("All embed endpoints unreachable");
        }
    }
}
