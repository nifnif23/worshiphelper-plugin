// ============================================================================
// AddIn/Pipeline.cs
// The end-to-end pipeline wired into ThisAddIn:
//
//   SpeechListener
//      --> SpeechRecognised
//              |
//              +-> CorrectionEngine.TryParaphraseLookup  (cheap, user rules)
//              +-> PatternMatcher.DetectBest             (explicit references)
//              +-> SemanticSearch.FindAsync              (paraphrase fallback)
//                     |
//                     v
//            ReferenceDetectedEventArgs (unified)
//                     |
//                     v
//            InsertScriptureFromSpeech (unchanged)
//
// All orchestration lives here; ThisAddIn.cs just calls InitialisePipeline()
// on startup and DisposePipeline() on shutdown.
// ============================================================================
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using log4net;
using Microsoft.Win32;
using WorshipHelperVSTO.Detection;
using WorshipHelperVSTO.Feedback;

namespace WorshipHelperVSTO
{
    public partial class ThisAddIn
    {
        // -- Speech service (kept for ribbon/toast compatibility) --------
        private SpeechToScriptureService _speechService;
        public SpeechToScriptureService SpeechService => _speechService;
        public AutoScriptureMode AutoScripture => AutoScriptureMode.Instance;

        // -- New pipeline pieces -----------------------------------------
        private VerseDatabase     _verseDb;
        private SemanticSearch    _semantic;
        private PatternMatcher    _pattern;
        private FeedbackStore     _feedback;
        private CorrectionEngine  _corrections;
        private RuntimeAdjustments _adjustments;

        // -----------------------------------------------------------------
        private void InitialiseSpeechService()
        {
            try
            {
                _speechService = new SpeechToScriptureService();
                _speechService.OnStatusChanged += SpeechService_OnStatusChanged;

                _pattern    = new PatternMatcher();
                _feedback   = new FeedbackStore();
                _corrections = new CorrectionEngine(_feedback);
                _adjustments = _corrections.Recompute();

                // Load semantic DB (optional -- graceful if missing).
                _verseDb = new VerseDatabase();
                string dbPath = System.IO.Path.Combine(appDataPath ?? "", "verses.sqlite");
                try { _verseDb.Load(dbPath); }
                catch (Exception ex) { log.Warn("Verse DB load failed: " + ex.Message); }
                _semantic = new SemanticSearch(_verseDb);

                _speechService.OnRawSpeech += async (s, e) =>
                    await HandleRawSpeechAsync(e.Text, e.Confidence).ConfigureAwait(false);

                log.Info("WorshipHelper speech pipeline initialised (Faster-Whisper + semantic).");
            }
            catch (Exception ex)
            {
                log.Error("Failed to initialise speech pipeline.", ex);
            }
        }

        private void ShutdownSpeechService()
        {
            try
            {
                AutoScriptureMode.Instance.Disable();
                _speechService?.Dispose();
                _speechService = null;
            }
            catch (Exception ex) { log.Warn("Error disposing speech service.", ex); }
        }

        // -----------------------------------------------------------------
        private async Task HandleRawSpeechAsync(string utterance, float speechConfidence)
        {
            if (string.IsNullOrWhiteSpace(utterance)) return;

            // 1. User-defined paraphrase shortcut.
            string custom = _corrections?.TryParaphraseLookup(utterance, _adjustments);
            if (custom != null)
            {
                FireReference(custom, utterance, speechConfidence, source: "paraphrase");
                return;
            }

            // 2. Let the existing SpeechToScriptureService handle explicit-reference
            //    detection via its own internal wiring (OnReferenceDetected).

            // 3. Semantic fallback -- only if pattern detector would not fire.
            var best = _pattern.DetectBest(utterance);
            if (best != null && best.Confidence >= _adjustments.PatternThreshold) return;

            if (_semantic != null)
            {
                try
                {
                    var match = await _semantic.FindAsync(utterance).ConfigureAwait(false);
                    if (match != null && match.Similarity >= _adjustments.SemanticThreshold)
                    {
                        log.Info($"Semantic match: \"{utterance}\" -> {match.Reference} " +
                                 $"(sim={match.Similarity:F2}, margin={match.Margin:F2})");
                        FireReference(match.Reference, utterance, match.Similarity, "semantic");
                    }
                }
                catch (Exception ex) { log.Debug("Semantic search failed: " + ex.Message); }
            }
        }

        private void FireReference(string reference, string spoken, double confidence, string source)
        {
            _feedback?.Record(new FeedbackRecord
            {
                Kind = FeedbackKind.AutoInsert,
                SpokenText = spoken,
                DetectedReference = reference,
                Confidence = confidence,
            });

            var args = new ReferenceDetectedEventArgs
            {
                NormalisedReference = reference,
                SpokenText = spoken,
                Confidence = confidence,
            };
            SpeechService_OnReferenceDetected(this, args);
        }

        // -----------------------------------------------------------------
        private void SpeechService_OnReferenceDetected(object sender, ReferenceDetectedEventArgs e)
        {
            log.Info($"Reference detected: \"{e.NormalisedReference}\" " +
                     $"(spoken: \"{e.SpokenText}\", conf={e.Confidence:F2})");

            try
            {
                if (AutoScriptureMode.Instance.IsEnabled)
                {
                    MarshalToUi(() => AutoScriptureMode.Instance.HandleDetectedReference(
                        e.NormalisedReference, e.SpokenText, InsertScriptureFromSpeech));
                    return;
                }
                MarshalToUi(() => InsertScriptureFromSpeech(e.NormalisedReference));
            }
            catch (Exception ex) { log.Error("Error inserting scripture from speech.", ex); }
        }

        private static void MarshalToUi(Action action)
        {
            var mainForm = System.Windows.Forms.Application.OpenForms.Count > 0
                ? System.Windows.Forms.Application.OpenForms[0] : null;
            if (mainForm != null && mainForm.InvokeRequired)
                mainForm.BeginInvoke(action);
            else
                action();
        }

        private void InsertScriptureFromSpeech(string normalisedReference)
        {
            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            var lastBible = registryKey.GetValue("LastBibleTranslation") as string ?? "NASB";
            var lastTemplate = registryKey.GetValue("LastScriptureTemplate") as string;
            var multiVerseSetting = registryKey.GetValue("MultiVerseProjection");
            bool multiVerse = multiVerseSetting != null && (int)multiVerseSetting == 1;

            var bible = OpenSongBibleReader.LoadTranslation(lastBible);

            ScriptureTemplate template = null;
            var installedTemplateFiles = System.IO.Directory.GetFiles(
                $@"{appDataPath}\Templates", "*.pptx");
            System.IO.Directory.CreateDirectory(
                $@"{userDataPath}\UserTemplates\Scripture");
            var userTemplateFiles = System.IO.Directory.GetFiles(
                $@"{userDataPath}\UserTemplates\Scripture", "*.pptx");

            foreach (var file in installedTemplateFiles.Concat(userTemplateFiles))
            {
                var t = new ScriptureTemplate(file);
                if (t.name == lastTemplate) { template = t; break; }
            }
            if (template == null)
            {
                var firstFile = installedTemplateFiles.Concat(userTemplateFiles).FirstOrDefault();
                if (firstFile == null) { log.Error("No scripture templates found."); return; }
                template = new ScriptureTemplate(firstFile);
            }

            var parsed = FullReferenceParser.ParseFullReference(bible, normalisedReference);
            if (parsed == null) { log.Warn("Could not parse: " + normalisedReference); return; }
            var (bookName, refPart) = parsed.Value;
            if (string.IsNullOrWhiteSpace(refPart)) { log.Warn("No chapter/verse: " + normalisedReference); return; }

            var parsedRef = ScriptureReferenceParser.Parse(refPart);
            var book = bible.books.First(b => b.name.Equals(bookName, StringComparison.OrdinalIgnoreCase));
            var chapter = book.chapters.FirstOrDefault(c => c.number == parsedRef.Chapter);
            if (chapter == null) { log.Warn($"Chapter {parsedRef.Chapter} not found in {bookName}."); return; }

            int maxVerse = chapter.verses.OrderBy(v => v.number).Last().number;
            var verseNumbers = new System.Collections.Generic.List<int>();
            foreach (var range in parsedRef.Ranges)
            {
                int s = Math.Max(1, range.Start);
                int end = range.End == int.MaxValue ? maxVerse : range.End;
                end = Math.Min(maxVerse, end);
                for (int v = s; v <= end; v++) verseNumbers.Add(v);
            }
            verseNumbers = verseNumbers.Distinct().OrderBy(v => v).ToList();
            if (!verseNumbers.Any()) { log.Warn("No valid verses: " + normalisedReference); return; }

            var sm = new ScriptureManager();
            int firstSlide = sm.addScripture(template, bible, bookName, parsedRef.Chapter, verseNumbers, multiVerse);
            log.Info($"Inserted scripture from speech: {normalisedReference} (slide {firstSlide})");

            try
            {
                if (firstSlide > 0) new SelectionManager().NavigateToInsertedSlide(firstSlide);
            }
            catch (Exception navEx) { log.Warn($"Navigation failed: {navEx.Message}"); }
        }

        private void SpeechService_OnStatusChanged(object sender, ServiceStatusEventArgs e)
        {
            log.Debug($"Speech status: {e.Message} (listening={e.IsListening}, err={e.IsError})");
        }
    }
}
