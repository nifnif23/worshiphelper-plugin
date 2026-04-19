// ============================================================================
// ThisAddIn_SpeechIntegration.cs
// 
// INTEGRATION EXAMPLE — shows how to wire the speech-to-scripture pipeline
// into your existing WorshipHelper VSTO add-in.
//
// This is NOT a separate file to drop in — it shows the CHANGES you need to
// make to your existing ThisAddIn.cs and TestRibbonItem.cs (or a new ribbon).
//
// ============================================================================

// ═══════════════════════════════════════════════════════════════════════════
// STEP 1:  Add reference to System.Speech
// ═══════════════════════════════════════════════════════════════════════════
//
// In Visual Studio:
//   1. Right-click "References" in the WorshipHelperVSTO project
//   2. Click "Add Reference…"
//   3. Go to "Assemblies" → "Framework"
//   4. Check "System.Speech"
//   5. Click OK
//
// Or add this to WorshipHelperVSTO.csproj inside an <ItemGroup>:
//
//   <Reference Include="System.Speech" />
//


// ═══════════════════════════════════════════════════════════════════════════
// STEP 2:  Add the four new .cs files to your project
// ═══════════════════════════════════════════════════════════════════════════
//
// Copy these files into the WorshipHelperVSTO folder:
//   - SpokenNumberConverter.cs
//   - BibleReferenceDetector.cs
//   - SpeechListener.cs
//   - SpeechToScriptureService.cs
//
// Visual Studio should automatically pick them up.  If not, include them
// via "Add → Existing Item…".


// ═══════════════════════════════════════════════════════════════════════════
// STEP 3:  Modify ThisAddIn.cs — add the service lifecycle
// ═══════════════════════════════════════════════════════════════════════════

using System;
using System.Linq;
using System.Windows.Forms;
using log4net;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;

namespace WorshipHelperVSTO
{
    // This partial class extends your existing ThisAddIn.
    // You can either merge these members into your existing ThisAddIn.cs
    // or keep this as a separate partial-class file (ThisAddIn.Speech.cs).

    public partial class ThisAddIn
    {
        // -----------------------------------------------------------------------
        // Speech service instance — lives for the lifetime of the add-in
        // -----------------------------------------------------------------------
        private SpeechToScriptureService _speechService;

        /// <summary>
        /// Public accessor so the ribbon can toggle listening.
        /// </summary>
        public SpeechToScriptureService SpeechService => _speechService;

        // -----------------------------------------------------------------------
        // Call this from ThisAddIn_Startup, AFTER existing initialisation
        // -----------------------------------------------------------------------
        private void InitialiseSpeechService()
        {
            try
            {
                _speechService = new SpeechToScriptureService();

                // ---------------------------------------------------------------
                // THE KEY WIRING: when a reference is detected, insert scripture.
                // ---------------------------------------------------------------
                _speechService.OnReferenceDetected += SpeechService_OnReferenceDetected;
                _speechService.OnStatusChanged += SpeechService_OnStatusChanged;

                log.Info("Speech-to-scripture service initialised (not yet listening).");
            }
            catch (Exception ex)
            {
                log.Error("Failed to initialise speech service.", ex);
                // Non-fatal: the rest of the add-in still works.
            }
        }

        // -----------------------------------------------------------------------
        // Call this from ThisAddIn_Shutdown
        // -----------------------------------------------------------------------
        private void ShutdownSpeechService()
        {
            try
            {
                _speechService?.Dispose();
                _speechService = null;
            }
            catch (Exception ex)
            {
                log.Warn("Error disposing speech service.", ex);
            }
        }

        // -----------------------------------------------------------------------
        // Reference detected handler — THIS IS WHERE SPEECH MEETS YOUR EXISTING CODE
        // -----------------------------------------------------------------------
        private void SpeechService_OnReferenceDetected(object sender, ReferenceDetectedEventArgs e)
        {
            // ⚠ This fires on a background thread!
            // We must marshal to the main (STA) thread to interact with PowerPoint.

            log.Info($"Speech detected reference: \"{e.NormalisedReference}\" " +
                     $"(spoken: \"{e.SpokenText}\", confidence: {e.Confidence:F2})");

            try
            {
                // Option A: Direct insertion (simplest path)
                // This calls your existing ScriptureManager directly.
                InsertScriptureFromSpeech(e.NormalisedReference);

                // Option B: If you prefer to show a confirmation dialog first,
                // uncomment the block below and comment out Option A above.
                /*
                System.Windows.Forms.Application.OpenForms[0]?.Invoke(
                    (Action)(() =>
                    {
                        var result = MessageBox.Show(
                            $"Detected scripture reference:\n\n{e.NormalisedReference}\n\n" +
                            $"(Heard: \"{e.SpokenText}\")\n\nInsert this scripture?",
                            "Speech Detection",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question,
                            MessageBoxDefaultButton.Button1);

                        if (result == DialogResult.Yes)
                        {
                            InsertScriptureFromSpeech(e.NormalisedReference);
                        }
                    }));
                */
            }
            catch (Exception ex)
            {
                log.Error($"Error inserting scripture from speech: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Takes a normalised reference string (e.g. "John 3:16") and
        /// inserts the scripture using your existing infrastructure.
        ///
        /// This bridges speech detection → your existing ScriptureManager.addScripture().
        /// </summary>
        private void InsertScriptureFromSpeech(string normalisedReference)
        {
            // Load Bible and template from saved preferences (same as InsertScriptureForm does)
            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            var lastBible = registryKey.GetValue("LastBibleTranslation") as string ?? "NASB";
            var lastTemplate = registryKey.GetValue("LastScriptureTemplate") as string;
            var multiVerseSetting = registryKey.GetValue("MultiVerseProjection");
            bool multiVerse = multiVerseSetting != null && (int)multiVerseSetting == 1;

            var bible = OpenSongBibleReader.LoadTranslation(lastBible);

            // Find the template
            ScriptureTemplate template = null;
            var installedTemplateFiles = System.IO.Directory.GetFiles(
                $@"{ThisAddIn.appDataPath}\Templates", "*.pptx");
            System.IO.Directory.CreateDirectory(
                $@"{ThisAddIn.userDataPath}\UserTemplates\Scripture");
            var userTemplateFiles = System.IO.Directory.GetFiles(
                $@"{ThisAddIn.userDataPath}\UserTemplates\Scripture", "*.pptx");

            foreach (var file in installedTemplateFiles.Concat(userTemplateFiles))
            {
                var t = new ScriptureTemplate(file);
                if (t.name == lastTemplate)
                {
                    template = t;
                    break;
                }
            }

            // Fallback to first available template
            if (template == null)
            {
                var firstFile = installedTemplateFiles.Concat(userTemplateFiles).FirstOrDefault();
                if (firstFile == null)
                {
                    log.Error("No scripture templates found.");
                    return;
                }
                template = new ScriptureTemplate(firstFile);
            }

            // Use your existing FullReferenceParser to split "John 3:16" → book + reference
            var parsed = FullReferenceParser.ParseFullReference(bible, normalisedReference);
            if (parsed == null)
            {
                log.Warn($"Could not parse full reference: \"{normalisedReference}\"");
                return;
            }

            var (bookName, refPart) = parsed.Value;

            if (string.IsNullOrWhiteSpace(refPart))
            {
                log.Warn($"No chapter/verse in reference: \"{normalisedReference}\"");
                return;
            }

            // Parse the chapter:verse portion
            var parsedRef = ScriptureReferenceParser.Parse(refPart);
            var book = bible.books.First(b => b.name.Equals(bookName, StringComparison.OrdinalIgnoreCase));
            var chapter = book.chapters.FirstOrDefault(c => c.number == parsedRef.Chapter);

            if (chapter == null)
            {
                log.Warn($"Chapter {parsedRef.Chapter} not found in {bookName}.");
                return;
            }

            int maxVerse = chapter.verses.OrderBy(v => v.number).Last().number;

            // Expand ranges to verse list
            var verseNumbers = new System.Collections.Generic.List<int>();
            foreach (var range in parsedRef.Ranges)
            {
                int s = Math.Max(1, range.Start);
                int end = range.End == int.MaxValue ? maxVerse : range.End;
                end = Math.Min(maxVerse, end);
                for (int v = s; v <= end; v++)
                    verseNumbers.Add(v);
            }
            verseNumbers = verseNumbers.Distinct().OrderBy(v => v).ToList();

            if (!verseNumbers.Any())
            {
                log.Warn($"No valid verses for reference: \"{normalisedReference}\"");
                return;
            }

            // INSERT THE SCRIPTURE — calls your existing ScriptureManager!
            new ScriptureManager().addScripture(
                template, bible, bookName, parsedRef.Chapter, verseNumbers, multiVerse);

            log.Info($"Successfully inserted scripture from speech: {normalisedReference}");
        }

        private void SpeechService_OnStatusChanged(object sender, ServiceStatusEventArgs e)
        {
            log.Debug($"Speech service status: {e.Message} (listening={e.IsListening}, error={e.IsError})");
            // TODO: Update a ribbon label or status bar if you have one
        }
    }
}


// ═══════════════════════════════════════════════════════════════════════════
// STEP 4:  Update ThisAddIn_Startup and ThisAddIn_Shutdown
// ═══════════════════════════════════════════════════════════════════════════
//
// In your existing ThisAddIn.cs, add these calls:
//
//   private void ThisAddIn_Startup(object sender, EventArgs e)
//   {
//       _keyboardProc = KeyboardHookCallback;
//       log = LogManager.GetLogger("WorshipHelperVSTO");
//       log.Info("Initialised logger");
//       SetWindowsHooks();
//
//       // ★ ADD THIS LINE:
//       InitialiseSpeechService();
//   }
//
//   private void ThisAddIn_Shutdown(object sender, EventArgs e)
//   {
//       UnhookWindowsHooks();
//
//       // ★ ADD THIS LINE:
//       ShutdownSpeechService();
//   }


// ═══════════════════════════════════════════════════════════════════════════
// STEP 5:  Add a ribbon button to toggle speech listening
// ═══════════════════════════════════════════════════════════════════════════
//
// In your TestRibbonItem.cs (or a new ribbon), add a button handler:

namespace WorshipHelperVSTO
{
    public partial class TestRibbonItem
    {
        // Add this button click handler (wire it to a new ribbon button).
        // In the ribbon designer, add a ToggleButton named "btnSpeechListen"
        // to the "grpScripture" group (or wherever you prefer).

        private void btnSpeechListen_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            try
            {
                var service = Globals.ThisAddIn.SpeechService;
                if (service == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Speech service is not available. Please check that a microphone " +
                        "is connected and Windows Speech Recognition is enabled.",
                        "Speech Service",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                bool nowListening = service.Toggle();

                // Update button appearance
                var btn = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;
                if (btn != null)
                {
                    btn.Label = nowListening ? "🎤 Listening…" : "🎤 Listen";
                    btn.Checked = nowListening;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error toggling speech listener:\n\n{ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}


// ═══════════════════════════════════════════════════════════════════════════
// STEP 6 (OPTIONAL):  Keyboard shortcut to toggle listening
// ═══════════════════════════════════════════════════════════════════════════
//
// You can also add speech toggle to your existing keyboard hook.
// In ThisAddIn.cs → KeyboardHookCallback, add a new hotkey:
//
//   if (keyPressed == 'L' && (Control.ModifierKeys & Keys.Shift) != 0)
//   {
//       // Shift+L toggles speech listening
//       _speechService?.Toggle();
//   }
//
// Or use any other key combination that doesn't conflict with PowerPoint.


// ═══════════════════════════════════════════════════════════════════════════
// STEP 7:  Ribbon designer changes (TestRibbonItem.Designer.cs)
// ═══════════════════════════════════════════════════════════════════════════
//
// If adding via the ribbon designer, add this XML to your ribbon:
//
//   <toggleButton id="btnSpeechListen"
//                 label="🎤 Listen"
//                 screentip="Toggle speech recognition"
//                 supertip="Listens for spoken Bible references and automatically inserts them."
//                 onAction="btnSpeechListen_Click"
//                 imageMso="SpeechMicrophone" />
//
// Or add it programmatically in the designer .cs file.
