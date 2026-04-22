// ============================================================================
// Detection/SemanticSearch.cs
// Paraphrase/semantic lookup over the VerseDatabase.
//
// Philosophy: this is a SECOND-CHANCE matcher. If PatternMatcher already
// fired (high-confidence explicit reference), we don't touch it. If pattern
// matching fails AND the utterance is long enough to possibly be a quote,
// we embed it and ask for the top-k nearest verses. Only fire if the top
// match clears a high confidence threshold AND beats the runner-up by a
// meaningful margin.
//
// Embeddings are computed by calling the Python server via a dedicated
// HTTP POST /embed endpoint.
//
// If Python is unreachable, SemanticSearch silently no-ops (graceful
// degradation: pattern matching still works).
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
        { Timeout = TimeSpan.FromSeconds(3) };

        private readonly VerseDatabase _db;
        public Uri EmbedEndpoint { get; set; } = new Uri("http://127.0.0.1:8765/embed");

        /// <summary>Min cosine similarity before we consider reporting a match.</summary>
        public float MinSimilarity { get; set; } = 0.62f;

        /// <summary>Min margin between #1 and #2 so ambiguous matches are ignored.</summary>
        public float MinMargin { get; set; } = 0.05f;

        /// <summary>Short utterances are filtered out -- not enough signal.</summary>
        public int MinWordCount { get; set; } = 6;

        public SemanticSearch(VerseDatabase db) { _db = db; }

        public async Task<SemanticMatch> FindAsync(string utterance)
        {
            if (_db == null || !_db.IsLoaded) return null;
            if (string.IsNullOrWhiteSpace(utterance)) return null;
            if (utterance.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length < MinWordCount)
                return null;

            float[] query;
            try { query = await EmbedAsync(utterance).ConfigureAwait(false); }
            catch (Exception ex)
            {
                log.Debug("SemanticSearch: embed call failed -- " + ex.Message);
                return null;
            }
            if (query == null || query.Length != _db.Dimension) return null;

            // Top-2 scan -- tiny enough (~31k x 384) to run linearly on CPU.
            float best1 = -1, best2 = -1; VerseRow winner = null;
            foreach (var row in _db.Verses)
            {
                float sim = Cosine(query, row.Embedding);
                if (sim > best1) { best2 = best1; best1 = sim; winner = row; }
                else if (sim > best2) { best2 = sim; }
            }

            if (winner == null || best1 < MinSimilarity) return null;
            float margin = best1 - best2;
            if (margin < MinMargin) return null;

            return new SemanticMatch
            {
                Reference = winner.Reference,
                Text = winner.Text,
                Similarity = best1,
                Margin = margin,
            };
        }

        private static float Cosine(float[] a, float[] b)
        {
            // Both sides produced with normalize_embeddings=True, so plain dot product = cosine.
            float dot = 0;
            for (int i = 0; i < a.Length; i++) dot += a[i] * b[i];
            return dot;
        }

        private async Task<float[]> EmbedAsync(string text)
        {
            var payload = JsonConvert.SerializeObject(new { text });
            using (var req = new HttpRequestMessage(HttpMethod.Post, EmbedEndpoint)
            { Content = new StringContent(payload, Encoding.UTF8, "application/json") })
            using (var resp = await _http.SendAsync(req).ConfigureAwait(false))
            {
                resp.EnsureSuccessStatusCode();
                string json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                var arr = (JArray)JObject.Parse(json)["embedding"];
                var vec = new float[arr.Count];
                for (int i = 0; i < arr.Count; i++) vec[i] = (float)arr[i];
                return vec;
            }
        }
    }
}
