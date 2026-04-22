// ============================================================================
// Detection/VerseDatabase.cs
// Read-only wrapper over the SQLite DB produced by build_verse_embeddings.py.
// Loads all 31,102 verse embeddings into memory (approx 48 MB for MiniLM-L6-v2
// at 384 dims x 4 bytes = 1.5 KB / verse) for fast cosine search.
// ============================================================================
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using log4net;

namespace WorshipHelperVSTO.Detection
{
    public sealed class VerseRow
    {
        public int Id;
        public string Book;
        public int Chapter;
        public int Verse;
        public string Text;
        public float[] Embedding;

        public string Reference => $"{Book} {Chapter}:{Verse}";
    }

    public sealed class VerseDatabase
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(VerseDatabase));

        public IReadOnlyList<VerseRow> Verses { get; private set; } = Array.Empty<VerseRow>();
        public int Dimension { get; private set; }
        public string ModelName { get; private set; }
        public bool IsLoaded => Verses.Count > 0;

        public void Load(string sqlitePath)
        {
            if (!File.Exists(sqlitePath))
            {
                log.Warn($"VerseDatabase: file not found at {sqlitePath} -- semantic search disabled.");
                return;
            }

            var list = new List<VerseRow>(32_000);
            using (var conn = new SQLiteConnection("Data Source=" + sqlitePath + ";Version=3;Read Only=True;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT value FROM meta WHERE key='dim'", conn))
                using (var r = cmd.ExecuteReader())
                { if (r.Read()) Dimension = int.Parse(r.GetString(0)); }

                using (var cmd = new SQLiteCommand("SELECT value FROM meta WHERE key='model'", conn))
                using (var r = cmd.ExecuteReader())
                { if (r.Read()) ModelName = r.GetString(0); }

                using (var cmd = new SQLiteCommand(
                    "SELECT id, book, chapter, verse, text, embedding FROM verses", conn))
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        var blob = (byte[])r["embedding"];
                        var emb = new float[blob.Length / 4];
                        Buffer.BlockCopy(blob, 0, emb, 0, blob.Length);
                        list.Add(new VerseRow
                        {
                            Id = r.GetInt32(0),
                            Book = r.GetString(1),
                            Chapter = r.GetInt32(2),
                            Verse = r.GetInt32(3),
                            Text = r.GetString(4),
                            Embedding = emb,
                        });
                    }
                }
            }
            Verses = list;
            log.Info($"VerseDatabase: loaded {list.Count:N0} verses (dim={Dimension}, model={ModelName}).");
        }
    }
}
