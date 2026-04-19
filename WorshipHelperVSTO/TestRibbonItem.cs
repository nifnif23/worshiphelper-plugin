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
        private System.Threading.SynchronizationContext _uiContext;
        private SpeechDebugPanel _debugPanel;

        private SpeechToScriptureService SpeechService
        {
            get
            {
                if (_speechService == null)
                {
                    // Capture the UI SynchronizationContext on this (UI) thread.
                    // This is more reliable than Control.BeginInvoke in a VSTO host
                    // where the hidden-control approach can fail to pump messages
                    // when no WinForms message loop is actively running.
                    _uiContext = System.Threading.SynchronizationContext.Current
                                 ?? new System.Windows.Forms.WindowsFormsSynchronizationContext();

                    // Reuse the add-in-wide speech service if it's been created
                    // by ThisAddIn_Startup → InitialiseSpeechService(). Otherwise
                    // create our own (e.g. in unit-test or standalone contexts).
                    // Either way, only ONE SpeechToScriptureService instance
                    // exists in the process — avoiding double mic capture.
                    _speechService = Globals.ThisAddIn?.SpeechService
                                     ?? new SpeechToScriptureService();

                    _speechService.OnReferenceDetected += OnReferenceDetected;
                    _speechService.OnStatusChanged     += OnSpeechStatusChanged;
                    _speechService.OnRawSpeech         += OnRawSpeech;
                    if (_speechService.Listener != null)
                        _speechService.Listener.PhaseChanged += OnPhaseChanged;
                }
                return _speechService;
            }
        }

        private void OnPhaseChanged(object sender, SpeechPhaseChangedEventArgs e)
        {
            if (_debugPanel == null || _debugPanel.IsDisposed) return;
            _debugPanel.SetPhase(e.Phase, e.BookName);
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
            // Marshal to the UI thread via the SynchronizationContext captured when
            // the service was first created. This works whether or not the debug
            // panel (Monitor) is open — unlike the old Control.BeginInvoke approach,
            // which required a visible WinForms message pump to reliably fire.
            _uiContext?.Post(_ =>
            {
                try
                {
                    if (_debugPanel != null && !_debugPanel.IsDisposed)
                        _debugPanel.SetDetected(e.NormalisedReference);

                    // Auto Scripture Mode (Shift in presenter view) → insert immediately,
                    // no confirmation dialog.
                    if (AutoScriptureMode.Instance.IsEnabled)
                    {
                        InsertReference(e.NormalisedReference);
                        return;
                    }

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
            }, null);
        }

        private void InsertReference(string normalisedReference)
        {
            try
            {
                var regKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\WorshipHelper");
                var lastTranslation = regKey?.GetValue("LastBibleTranslation") as string ?? "ESV";

                var bible = OpenSongBibleReader.LoadTranslation(lastTranslation);

                // Template resolution order:
                //   1. SpeechScriptureTemplate — set specifically for speech via "Set Template" button
                //   2. LastScriptureTemplate   — the last template picked in InsertScriptureForm
                //   3. First available template — fallback
                var speechTemplateName = regKey?.GetValue("SpeechScriptureTemplate") as string;
                var fallbackTemplateName = regKey?.GetValue("LastScriptureTemplate") as string;

                var installedFiles = System.IO.Directory.GetFiles(
                    ThisAddIn.appDataPath + @"\Templates", "*.pptx");
                System.IO.Directory.CreateDirectory(
                    ThisAddIn.userDataPath + @"\UserTemplates\Scripture");
                var userFiles = System.IO.Directory.GetFiles(
                    ThisAddIn.userDataPath + @"\UserTemplates\Scripture", "*.pptx");
                var allTemplateFiles = installedFiles.Concat(userFiles).ToArray();

                if (allTemplateFiles.Length == 0) return;

                ScriptureTemplate template = null;
                foreach (var file in allTemplateFiles)
                {
                    var t = new ScriptureTemplate(file);
                    if (speechTemplateName != null && t.name == speechTemplateName)
                        { template = t; break; }
                    if (template == null && fallbackTemplateName != null && t.name == fallbackTemplateName)
                        template = t;
                }
                if (template == null)
                    template = new ScriptureTemplate(allTemplateFiles[0]);

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

        /// <summary>
        /// Opens a simple picker so the user can choose which slide template the
        /// speech listener inserts scripture into. The choice is saved under
        /// SpeechScriptureTemplate in the registry and is independent of the
        /// LastScriptureTemplate setting used by the manual Insert Scripture form.
        /// </summary>
        private void btnSpeechTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var installedFiles = System.IO.Directory.GetFiles(
                    ThisAddIn.appDataPath + @"\Templates", "*.pptx");
                System.IO.Directory.CreateDirectory(
                    ThisAddIn.userDataPath + @"\UserTemplates\Scripture");
                var userFiles = System.IO.Directory.GetFiles(
                    ThisAddIn.userDataPath + @"\UserTemplates\Scripture", "*.pptx");
                var allFiles = installedFiles.Concat(userFiles).ToArray();

                if (allFiles.Length == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "No scripture templates found in:\n" + ThisAddIn.appDataPath + @"\Templates",
                        "No Templates", System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // Read current speech template selection
                var regKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\WorshipHelper");
                var currentName = regKey?.GetValue("SpeechScriptureTemplate") as string;

                // Build a simple picker form
                using (var picker = new System.Windows.Forms.Form())
                {
                    picker.Text = "Speech: Choose Template";
                    picker.Size = new System.Drawing.Size(340, 220);
                    picker.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                    picker.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                    picker.MaximizeBox = false; picker.MinimizeBox = false;

                    var lbl = new System.Windows.Forms.Label
                    {
                        Text = "Template to use when speech inserts scripture:",
                        Location = new System.Drawing.Point(12, 12),
                        Size = new System.Drawing.Size(300, 32),
                        Font = new System.Drawing.Font("Segoe UI", 9f)
                    };
                    picker.Controls.Add(lbl);

                    var list = new System.Windows.Forms.ListBox
                    {
                        Location = new System.Drawing.Point(12, 48),
                        Size = new System.Drawing.Size(300, 100),
                        Font = new System.Drawing.Font("Segoe UI", 9.5f)
                    };
                    foreach (var f in allFiles)
                        list.Items.Add(new ScriptureTemplate(f));
                    // Pre-select current
                    for (int i = 0; i < list.Items.Count; i++)
                        if (((ScriptureTemplate)list.Items[i]).name == currentName)
                            { list.SelectedIndex = i; break; }
                    if (list.SelectedIndex < 0 && list.Items.Count > 0)
                        list.SelectedIndex = 0;
                    picker.Controls.Add(list);

                    var btnOk = new System.Windows.Forms.Button
                    {
                        Text = "OK", DialogResult = System.Windows.Forms.DialogResult.OK,
                        Location = new System.Drawing.Point(152, 158), Size = new System.Drawing.Size(75, 26)
                    };
                    var btnCancel = new System.Windows.Forms.Button
                    {
                        Text = "Cancel", DialogResult = System.Windows.Forms.DialogResult.Cancel,
                        Location = new System.Drawing.Point(237, 158), Size = new System.Drawing.Size(75, 26)
                    };
                    picker.Controls.Add(btnOk);
                    picker.Controls.Add(btnCancel);
                    picker.AcceptButton = btnOk;
                    picker.CancelButton = btnCancel;

                    if (picker.ShowDialog() == System.Windows.Forms.DialogResult.OK
                        && list.SelectedItem is ScriptureTemplate chosen)
                    {
                        Registry.CurrentUser
                            .CreateSubKey(@"SOFTWARE\WorshipHelper")
                            .SetValue("SpeechScriptureTemplate", chosen.name);

                        // Update the ribbon button label to confirm
                        btnSpeechTemplate.Label = "Template: " + chosen.name;
                    }
                }
            }
            catch (System.Exception ex)
            {
                log4net.LogManager.GetLogger(typeof(TestRibbonItem)).Error("Template picker failed", ex);
            }
        }

        private void OnRawSpeech(object sender, SpeechRecognisedEventArgs e)
        {
            if (_debugPanel == null || _debugPanel.IsDisposed) return;
            // Pass confidence as a separate parameter so the confidence bar
            // can render it independently of the heard-text label.
            _debugPanel.SetHeard(e.Text, e.Confidence);
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
            try
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
                    // Turning off listening also kills auto scripture mode
                    AutoScriptureMode.Instance.Disable();
                }

                if (_debugPanel != null && !_debugPanel.IsDisposed)
                    _debugPanel.SetListening(isNowListening);
            }
            catch (System.IO.DirectoryNotFoundException ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    ex.Message + "\n\nDownload a Vosk model from https://alphacephei.com/vosk/models " +
                    "and extract it to the path shown above.",
                    "Vosk Model Not Found",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Failed to start speech recognition:\n\n{ex.Message}",
                    "Speech Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        private void TestRibbonItem_Load(object sender, RibbonUIEventArgs e)
        {
            // Restore saved speech template label so user can see what's selected
            var speechTemplate = Registry.CurrentUser
                .OpenSubKey(@"SOFTWARE\WorshipHelper")
                ?.GetValue("SpeechScriptureTemplate") as string;
            if (!string.IsNullOrEmpty(speechTemplate))
                btnSpeechTemplate.Label = "Template: " + speechTemplate;

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
