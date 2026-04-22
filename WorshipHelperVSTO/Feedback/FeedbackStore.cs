// ============================================================================
// Feedback/FeedbackStore.cs
// Records every detection outcome (auto-insert, manual correction, miss) into
// a local SQLite file. Also stores user-defined paraphrase -> reference pairs
// that CorrectionEngine can promote into regex or seed embeddings from.
//
// Location: %APPDATA%\WorshipHelper\feedback.sqlite
// ============================================================================
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using log4net;

namespace WorshipHelperVSTO.Feedback
{
    public enum FeedbackKind
    {
        AutoInsert,      // detected + inserted
        UserCorrected,   // user typed the correct reference after a miss
        Missed,          // user said it but nothing fired
        Paraphrase,      // user-added custom mapping
    }

    public sealed class FeedbackRecord
    {
        public long   Id;
        public DateTime CreatedUtc;
        public FeedbackKind Kind;
        public string SpokenText;        // what STT heard
        public string DetectedReference; // what the detector produced (may be null)
        public string CorrectReference;  // ground truth (may be null until user corrects)
        public double Confidence;
    }

    public sealed class FeedbackStore
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(FeedbackStore));

        private readonly string _path;

        public FeedbackStore(string path = null)
        {
            _path = path ?? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "WorshipHelper", "feedback.sqlite");
            Directory.CreateDirectory(Path.GetDirectoryName(_path));
            EnsureSchema();
        }

        private SQLiteConnection Open()
        {
            var c = new SQLiteConnection("Data Source=" + _path + ";Version=3;");
            c.Open();
            return c;
        }

        private void EnsureSchema()
        {
            using (var c = Open())
            using (var cmd = c.CreateCommand())
            {
                cmd.CommandText = @"
                CREATE TABLE IF NOT EXISTS feedback (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_utc TEXT NOT NULL,
                    kind TEXT NOT NULL,
                    spoken_text TEXT,
                    detected_reference TEXT,
                    correct_reference TEXT,
                    confidence REAL
                );
                CREATE INDEX IF NOT EXISTS idx_kind ON feedback(kind);
                CREATE TABLE IF NOT EXISTS paraphrase (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    phrase TEXT NOT NULL,
                    reference TEXT NOT NULL,
                    created_utc TEXT NOT NULL
                );";
                cmd.ExecuteNonQuery();
            }
        }

        // ------------------------------------------------------------------
        public void Record(FeedbackRecord r)
        {
            try
            {
                using (var c = Open())
                using (var cmd = c.CreateCommand())
                {
                    cmd.CommandText = @"INSERT INTO feedback
                        (created_utc, kind, spoken_text, detected_reference, correct_reference, confidence)
                        VALUES (@t,@k,@s,@d,@cr,@cf)";
                    cmd.Parameters.AddWithValue("@t", DateTime.UtcNow.ToString("o"));
                    cmd.Parameters.AddWithValue("@k", r.Kind.ToString());
                    cmd.Parameters.AddWithValue("@s", (object)r.SpokenText ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@d", (object)r.DetectedReference ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@cr", (object)r.CorrectReference ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@cf", r.Confidence);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex) { log.Warn("FeedbackStore.Record failed: " + ex.Message); }
        }

        public void AddParaphrase(string phrase, string reference)
        {
            if (string.IsNullOrWhiteSpace(phrase) || string.IsNullOrWhiteSpace(reference)) return;
            try
            {
                using (var c = Open())
                using (var cmd = c.CreateCommand())
                {
                    cmd.CommandText = @"INSERT INTO paraphrase (phrase, reference, created_utc)
                                        VALUES (@p,@r,@t)";
                    cmd.Parameters.AddWithValue("@p", phrase.Trim().ToLowerInvariant());
                    cmd.Parameters.AddWithValue("@r", reference.Trim());
                    cmd.Parameters.AddWithValue("@t", DateTime.UtcNow.ToString("o"));
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex) { log.Warn("FeedbackStore.AddParaphrase failed: " + ex.Message); }
        }

        public IReadOnlyList<(string Phrase, string Reference)> GetParaphrases()
        {
            var list = new List<(string, string)>();
            try
            {
                using (var c = Open())
                using (var cmd = c.CreateCommand())
                {
                    cmd.CommandText = "SELECT phrase, reference FROM paraphrase";
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                            list.Add((r.GetString(0), r.GetString(1)));
                }
            }
            catch (Exception ex) { log.Warn("FeedbackStore.GetParaphrases failed: " + ex.Message); }
            return list;
        }

        public IReadOnlyDictionary<string, int> MishearingFrequency()
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var c = Open())
                using (var cmd = c.CreateCommand())
                {
                    cmd.CommandText = @"SELECT spoken_text, COUNT(*) FROM feedback
                                        WHERE kind IN ('UserCorrected','Missed')
                                          AND spoken_text IS NOT NULL
                                        GROUP BY spoken_text";
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                            map[r.GetString(0)] = r.GetInt32(1);
                }
            }
            catch (Exception ex) { log.Warn("MishearingFrequency failed: " + ex.Message); }
            return map;
        }
    }
}
