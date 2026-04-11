using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using System.Linq;

namespace WorshipHelperVSTO
{
    public partial class TestRibbonItem
    {
        private SpeechToScriptureService _speechService;
        private System.Windows.Forms.Control _uiThreadMarshaller;
        private SpeechDebugPanel _debugPanel;

        private SpeechToScriptureService SpeechService
        {
            get
            {
                if (_speechService == null)
                {
                    _uiThreadMarshaller = new System.Windows.Forms.Control();
                    _uiThreadMarshaller.CreateControl();
                    _speechService = new SpeechToScriptureService();
                    _speechService.OnReferenceDetected += OnReferenceDetected;
                    _speechService.OnStatusChanged     += OnSpeechStatusChanged;
                    _speechService.OnRawSpeech         += OnRawSpeech;
                }
                return _speechService;
            }
        }

        private SpeechDebugPanel DebugPanel
        {
            get
            {
                if (_debugPanel == null || _debugPanel.IsDisposed)
                    _debugPanel = new SpeechDebugPanel();
                return _debugPanel;
            }
        }

        private void OnReferenceDetected(object sender, ReferenceDetectedEventArgs e)
        {
            if (_uiThreadMarshaller == null || !_uiThreadMarshaller.IsHandleCreated) return;

            _uiThreadMarshaller.BeginInvoke(new System.Action(() =>
            {
                try
                {
                    if (_debugPanel != null && !_debugPanel.IsDisposed)
                        _debugPanel.SetDetected(e.NormalisedReference);

                    using (var toast = new SpeechConfirmForm(e.NormalisedReference))
                    {
                        toast.ShowDialog();
                        if (!toast.Confirmed) return;
                        InsertReference(toast.Reference);
                    }
                }
                catch (System.Exception ex)
                {
                    log4net.LogManager.GetLogger(typeof(TestRibbonItem)).Error("Speech toast failed", ex);
                }
            }));
        }

        private void InsertReference(string normalisedReference)
        {
            try
            {
                var lastTranslation = Registry.CurrentUser
                    .OpenSubKey(@"SOFTWARE\WorshipHelper")
                    ?.GetValue("LastBibleTranslation") as string ?? "ESV";

                var bible = OpenSongBibleReader.LoadTranslation(lastTranslation);

                var templateFiles = System.IO.Directory.GetFiles(
                    ThisAddIn.appDataPath + @"\Templates", "*.pptx");
                if (templateFiles.Length == 0) return;
                var template = new ScriptureTemplate(templateFiles[0]);

                var parsed = FullReferenceParser.ParseFullReference(bible, normalisedReference);
                if (parsed == null) return;

                var scriptureRef = ScriptureReference.parse(bible, parsed.Value.BookName, parsed.Value.Reference);
                if (scriptureRef == null) return;

                var verseNums = Enumerable.Range(scriptureRef.verseNumStart,
                    scriptureRef.verseNumEnd - scriptureRef.verseNumStart + 1).ToList();

                new ScriptureManager().addScripture(template, bible,
                    scriptureRef.bookName, scriptureRef.chapterNum, verseNums, verseNums.Count > 1);
            }
            catch (System.Exception ex)
            {
                log4net.LogManager.GetLogger(typeof(TestRibbonItem)).Error("Speech insert failed", ex);
            }
        }

        private void OnRawSpeech(object sender, SpeechRecognisedEventArgs e)
        {
            if (_debugPanel == null || _debugPanel.IsDisposed) return;
            _debugPanel.SetHeard($"{e.Text}  (conf {e.Confidence:F2})");
        }

        private void OnSpeechStatusChanged(object sender, ServiceStatusEventArgs e)
        {
            if (_debugPanel == null || _debugPanel.IsDisposed) return;
            _debugPanel.SetStatus(e.Message, e.IsError);
        }

        private void btnMonitorSpeech_Click(object sender, RibbonControlEventArgs e)
        {
            DebugPanel.Show();
            DebugPanel.BringToFront();
        }

        private void btnToggleSpeech_Click(object sender, RibbonControlEventArgs e)
        {
            bool isNowListening = SpeechService.Toggle();

            if (isNowListening)
            {
                btnToggleSpeech.Image = global::WorshipHelperVSTO.Properties.Resources.mic_active;
                btnToggleSpeech.Label = "Listening...";
            }
            else
            {
                btnToggleSpeech.Image = global::WorshipHelperVSTO.Properties.Resources.mic;
                btnToggleSpeech.Label = "Listen";
            }

            if (_debugPanel != null && !_debugPanel.IsDisposed)
                _debugPanel.SetListening(isNowListening);
        }
        private void TestRibbonItem_Load(object sender, RibbonUIEventArgs e)
        {
            var favRegistryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper\Favourites");
            var favouriteCount = favRegistryKey.GetValueNames().Length;
            for (int i=0; i < favouriteCount; i++)
            {
                var file = favRegistryKey.GetValueNames()[i];
                var slideButton = favouriteButtons[i];
                
                var pathParts = file.Split(new char[] { '\\' });
                slideButton.Label = pathParts[pathParts.Length - 1].Replace(".pptx", "").Replace(".ppt", "");
                slideButton.Tag = file;
                slideButton.Visible = true;
            }

            // Hide the unused buttons
            for (int i = favouriteCount; i < favouriteButtons.Count; i++)
            {
                var slideButton = favouriteButtons[i];
                slideButton.Visible = false;
            }
            btnAddFavourite.Enabled = favouriteCount < favouriteButtons.Count;

            #if !DEBUG
            grpDebug.Visible = false;
            #endif
        }

        private void btnInsertSong_Click(object sender, RibbonControlEventArgs e)
        {
            new SongManager().InsertSong();
        }

        private void btnInsertScripture_Click(object sender, RibbonControlEventArgs e)
        {
            // FIX: Use ShowDialog() so the form is modal and properly disposed after closing.
            // Previously used Show() which could cause issues with multiple forms open.
            using (var form = new InsertScriptureForm())
            {
                form.ShowDialog();
            }
        }

        private void btnInsertOneClick_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;

            string fileName;
            if(sender is RibbonSplitButton)
            {
                // This IS the parent SplitButton
                fileName = (sender as RibbonControl).Tag as string;
            } else
            {
                // Get it from the tag of the parent SplitButton
                fileName = (sender as RibbonControl).Parent.Tag as string;
            }

            new SongManager().InsertSongFromFile(fileName);
        }

        private void btnRemoveOneClick_Click(object sender, RibbonControlEventArgs e)
        {
            var favRegistryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper\Favourites");
            var fileName = (sender as RibbonControl).Parent.Tag as string;
            try
            {
                favRegistryKey.DeleteValue(fileName);
            } catch (System.ArgumentException)
            {
                System.Windows.Forms.MessageBox.Show("This item appears to have already been deleted - try restarting PowerPoint.");
            }

            // Force a refresh
            TestRibbonItem_Load(null, null);
        }

        private void btnAddFavourite_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;

            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            var favRegistryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper\Favourites");

            if (favRegistryKey.GetValueNames().Length >= 5) {
                System.Windows.Forms.MessageBox.Show("No more favourites can be added");
            }
                        
            var lastSongLocation = registryKey.GetValue("LastSongLocation") as string;

            FileDialog dialog = app.FileDialog[MsoFileDialogType.msoFileDialogOpen];
            dialog.Title = "Select Song or Presentation";
            if (lastSongLocation != null) dialog.InitialFileName = lastSongLocation;
            if (dialog.Show() == -1) // If user selected a file
            {
                foreach (string item in dialog.SelectedItems)
                {
                    favRegistryKey.SetValue(item, item);
                }
            }

            // Force a refresh
            TestRibbonItem_Load(null, null);
        }

        private void btnSelfTest_Click(object sender, RibbonControlEventArgs e)
        {
            new SelfTestManager().SelfTest();
        }
    }
}
