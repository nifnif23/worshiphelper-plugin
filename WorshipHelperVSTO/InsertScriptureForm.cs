using log4net;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WorshipHelperVSTO
{
    public partial class InsertScriptureForm : Form
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        Bible bible;

        // Debounce: prevent rapid double-click / Enter from triggering multiple inserts
        private DateTime _lastInsertTime = DateTime.MinValue;
        private const int DEBOUNCE_MS = 1000;

        // Custom suggestion dropdown (replaces built-in autocomplete for popularity ranking)
        private ListBox _suggestionBox;
        private bool _suppressSuggestions = false;

        // Color scheme constants
        private static readonly Color AccentGreen = Color.FromArgb(46, 125, 50);
        private static readonly Color DarkGreen = Color.FromArgb(27, 94, 32);
        private static readonly Color LightGreen = Color.FromArgb(232, 245, 233);
        private static readonly Color SuccessGreen = Color.FromArgb(56, 142, 60);
        private static readonly Color Gold = Color.FromArgb(184, 134, 11);

        public InsertScriptureForm()
        {
            log.Info("Loading InsertScriptureForm");
            InitializeComponent();

            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            var lastTemplate = registryKey.GetValue("LastScriptureTemplate") as string;
            var lastBible = registryKey.GetValue("LastBibleTranslation") as string;
            var multiVerseSetting = registryKey.GetValue("MultiVerseProjection");

            // Restore multi-verse checkbox state (defaults to false / unchecked)
            chkMultiVerse.Checked = multiVerseSetting != null && (int)multiVerseSetting == 1;

            // Get a list of available templates, populate list and set initial selection
            log.Debug("Loading scripture templates");
            var installedTemplateFiles = Directory.GetFiles($@"{ThisAddIn.appDataPath}\Templates", "*.pptx");
            Directory.CreateDirectory($@"{ThisAddIn.userDataPath}\UserTemplates\Scripture");
            var userTemplateFiles = Directory.GetFiles($@"{ThisAddIn.userDataPath}\UserTemplates\Scripture", "*.pptx");
            foreach (var file in installedTemplateFiles.Concat(userTemplateFiles))
            {
                var template = new ScriptureTemplate(file);
                cmbTemplate.Items.Add(template);
                if (template.name == lastTemplate)
                {
                    cmbTemplate.SelectedItem = template;
                }
            }
            if (cmbTemplate.SelectedItem == null && cmbTemplate.Items.Count > 0)
            {
                cmbTemplate.SelectedIndex = 0;
            }

            // Get a list of installed bibles, populate list and set initial selection
            log.Debug("Loading bibles");
            var installedBibleFiles = Directory.GetFiles($@"{ThisAddIn.appDataPath}\Bibles", "*.xmm");
            foreach (var file in installedBibleFiles)
            {
                var bibleName = file.Split(new char[] { '\\' }).Last().Replace(".xmm", "");
                cmbTranslation.Items.Add(bibleName);
                if (bibleName == lastBible)
                {
                    cmbTranslation.SelectedItem = bibleName;
                }
            }
            if (cmbTranslation.SelectedItem == null && cmbTranslation.Items.Count > 0)
            {
                cmbTranslation.SelectedIndex = 0;
            }

            // Initialise so that we can populate the books
            log.Debug($"Loading default bible ({cmbTranslation.SelectedItem})");
            bible = OpenSongBibleReader.LoadTranslation(cmbTranslation.SelectedItem as string);

            // ------------------------------------------------------------------
            // Custom suggestion dropdown (popularity-ranked)
            // Replaces the built-in AutoComplete which always sorts alphabetically.
            // ------------------------------------------------------------------
            txtBook.AutoCompleteMode = AutoCompleteMode.None;  // disable built-in
            SetupSuggestionBox();

            btnInsert.Enabled = false;

            // Start in single-reference mode
            SetMode(false);
        }

        // -----------------------------------------------------------------------
        // Custom suggestion dropdown
        // -----------------------------------------------------------------------
        private void SetupSuggestionBox()
        {
            _suggestionBox = new ListBox();
            _suggestionBox.Font = new Font("Segoe UI", 9.5F);
            _suggestionBox.Visible = false;
            _suggestionBox.IntegralHeight = false; // allow partial-item height
            _suggestionBox.BorderStyle = BorderStyle.FixedSingle;
            _suggestionBox.Cursor = Cursors.Hand;
            _suggestionBox.BackColor = Color.White;
            _suggestionBox.ForeColor = Color.FromArgb(33, 33, 33);

            // Position it directly below txtBook
            PositionSuggestionBox();

            // Selection events
            _suggestionBox.Click += SuggestionBox_Click;
            _suggestionBox.KeyDown += SuggestionBox_KeyDown;

            // Add it to the form and bring to front
            this.Controls.Add(_suggestionBox);
            _suggestionBox.BringToFront();
        }

        private void PositionSuggestionBox()
        {
            if (_suggestionBox == null || txtBook == null) return;
            _suggestionBox.Location = new Point(txtBook.Left, txtBook.Bottom + 2);
            _suggestionBox.Width = txtBook.Width;
        }

        private void UpdateSuggestions()
        {
            if (_suppressSuggestions || bible == null) return;

            string input = txtBook.Text;
            if (string.IsNullOrWhiteSpace(input))
            {
                _suggestionBox.Visible = false;
                return;
            }

            var suggestions = FullReferenceParser.GetSuggestions(bible, input);

            if (suggestions.Count == 0)
            {
                _suggestionBox.Visible = false;
                return;
            }

            // If the only suggestion IS the current text (exact match), hide
            if (suggestions.Count == 1 && suggestions[0].Equals(input, StringComparison.OrdinalIgnoreCase))
            {
                _suggestionBox.Visible = false;
                return;
            }

            _suggestionBox.BeginUpdate();
            _suggestionBox.Items.Clear();
            foreach (var s in suggestions)
                _suggestionBox.Items.Add(s);
            _suggestionBox.EndUpdate();

            // Size: show up to 8 items
            int visibleItems = Math.Min(suggestions.Count, 8);
            _suggestionBox.Height = visibleItems * _suggestionBox.ItemHeight + 4;

            PositionSuggestionBox();
            _suggestionBox.Visible = true;
            _suggestionBox.BringToFront();
        }

        private void AcceptSuggestion()
        {
            if (_suggestionBox.SelectedItem == null) return;

            _suppressSuggestions = true;
            txtBook.Text = _suggestionBox.SelectedItem.ToString();
            txtBook.SelectionStart = txtBook.Text.Length;
            _suggestionBox.Visible = false;
            _suppressSuggestions = false;

            // Move focus to reference field
            txtReference.Focus();
        }

        private void SuggestionBox_Click(object sender, EventArgs e)
        {
            AcceptSuggestion();
        }

        private void SuggestionBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                AcceptSuggestion();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                _suggestionBox.Visible = false;
                txtBook.Focus();
            }
        }

        // -----------------------------------------------------------------------
        // Mode switching
        // -----------------------------------------------------------------------
        private bool isBulkMode = false;

        private void SetMode(bool bulk)
        {
            isBulkMode = bulk;

            // Single-reference controls
            txtBook.Visible = !bulk;
            lblBook.Visible = !bulk;
            txtReference.Visible = !bulk;
            lblReference.Visible = !bulk;

            // Bulk controls
            txtBulk.Visible = bulk;
            lblBulk.Visible = bulk;
            lblBulkHint.Visible = bulk;

            btnModeSingle.Visible = bulk;
            btnModeBulk.Visible = !bulk;

            // Hide suggestion dropdown when switching modes
            if (_suggestionBox != null)
                _suggestionBox.Visible = false;

            if (bulk)
            {
                btnInsert.Enabled = !string.IsNullOrWhiteSpace(txtBulk.Text);
                lblStatus.Text = "Paste full references, one per line.";
                lblStatus.ForeColor = Color.FromArgb(117, 117, 117);
            }
            else
            {
                btnInsert.Enabled = isValidReference();
                lblStatus.Text = "";
            }
        }

        private void btnModeBulk_Click(object sender, EventArgs e)
        {
            SetMode(true);
        }

        private void btnModeSingle_Click(object sender, EventArgs e)
        {
            SetMode(false);
        }

        // -----------------------------------------------------------------------
        // Validation (single mode)
        // -----------------------------------------------------------------------
        private void txtSearchBox_TextChanged(object sender, EventArgs e)
        {
            if (!isBulkMode)
            {
                btnInsert.Enabled = isValidReference();
                UpdateSuggestions();
            }
        }

        private void txtSearchBox_KeyPress(object sender, KeyPressEventArgs e)
        {
        }

        /// <summary>
        /// Handle special keys in txtBook for suggestion navigation.
        /// </summary>
        private void txtBook_KeyDown(object sender, KeyEventArgs e)
        {
            if (_suggestionBox == null || !_suggestionBox.Visible) return;

            if (e.KeyCode == Keys.Down)
            {
                // Move focus into the suggestion box
                _suggestionBox.Focus();
                if (_suggestionBox.Items.Count > 0 && _suggestionBox.SelectedIndex < 0)
                    _suggestionBox.SelectedIndex = 0;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                _suggestionBox.Visible = false;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter && _suggestionBox.Visible && _suggestionBox.Items.Count > 0)
            {
                // Accept the first suggestion if none is selected
                if (_suggestionBox.SelectedIndex < 0)
                    _suggestionBox.SelectedIndex = 0;
                AcceptSuggestion();
                e.Handled = true;
                e.SuppressKeyPress = true; // prevent the Enter "ding"
            }
        }

        private bool isValidReference()
        {
            log.Debug($"Checking reference validity (book: {txtBook.Text}, reference: {txtReference.Text})");

            if (string.IsNullOrWhiteSpace(txtBook.Text) || bible == null)
                return false;

            // Use the new resolver that handles abbreviations
            var resolvedBook = FullReferenceParser.ResolveBookName(bible, txtBook.Text);
            if (resolvedBook == null)
                return false;

            if (string.IsNullOrWhiteSpace(txtReference.Text))
            {
                // Whole book? At least need chapter. Mark as invalid.
                return false;
            }

            try
            {
                ScriptureReferenceParser.Parse(txtReference.Text);
                return true;
            }
            catch (Exception ex)
            {
                log.Debug($"Reference parse failed: {ex.Message}");
                return false;
            }
        }

        // -----------------------------------------------------------------------
        // Bulk text changed
        // -----------------------------------------------------------------------
        private void txtBulk_TextChanged(object sender, EventArgs e)
        {
            if (isBulkMode)
            {
                btnInsert.Enabled = !string.IsNullOrWhiteSpace(txtBulk.Text);

                // Count parseable lines
                var lines = SplitBulkInput(txtBulk.Text);
                int valid = 0;
                foreach (var line in lines)
                {
                    var parsed = FullReferenceParser.ParseFullReference(bible, line);
                    if (parsed != null)
                    {
                        try
                        {
                            ScriptureReferenceParser.Parse(parsed.Value.Reference);
                            valid++;
                        }
                        catch { /* skip invalid */ }
                    }
                }
                lblStatus.Text = $"{valid} of {lines.Count} reference(s) recognised.";
                lblStatus.ForeColor = valid == lines.Count ? SuccessGreen : Gold;
            }
        }

        /// <summary>
        /// Splits bulk input text into individual reference strings.
        /// Supports newlines, semicolons, and pipe as delimiters.
        /// Filters out blanks.
        /// </summary>
        private List<string> SplitBulkInput(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return new List<string>();

            // Split on newlines, semicolons, or pipe
            var parts = Regex.Split(input, @"[\r\n;|]+");
            return parts
                .Select(p => p.Trim())
                .Where(p => p.Length > 0)
                .ToList();
        }

        // -----------------------------------------------------------------------
        // Insert
        // -----------------------------------------------------------------------
        private void btnInsert_Click(object sender, EventArgs e)
        {
            // Debounce: ignore rapid double-click / Enter presses
            if ((DateTime.Now - _lastInsertTime).TotalMilliseconds < DEBOUNCE_MS)
            {
                log.Debug("Ignoring rapid repeat insert (debounce)");
                return;
            }
            _lastInsertTime = DateTime.Now;

            log.Info("About to insert scripture");

            try
            {
                if (isBulkMode)
                {
                    InsertBulk();
                }
                else
                {
                    InsertSingle();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error inserting scripture", ex);
                MessageBox.Show($"Error inserting scripture:\n\n{ex.Message}",
                    "Insert Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // FIX: Close the form after successful insert instead of leaving it open.
            // The previous behaviour left it open which confused users.
            log.Debug("Scripture inserted successfully; closing form.");
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void InsertSingle()
        {
            // Resolve book name (handles abbreviations)
            var resolvedName = FullReferenceParser.ResolveBookName(bible, txtBook.Text);
            if (resolvedName == null)
                throw new Exception($"Could not resolve book name: {txtBook.Text}");

            var book = bible.books.First(b => b.name.Equals(resolvedName, StringComparison.OrdinalIgnoreCase));

            var parsed = ScriptureReferenceParser.Parse(txtReference.Text);

            log.Debug($"Parsed reference: chapter={parsed.Chapter}, ranges={string.Join(";", parsed.Ranges.Select(r => $"{r.Start}-{r.End}"))}");

            var chapter = book.chapters.First(c => c.number == parsed.Chapter);
            var verses = chapter.verses.OrderBy(v => v.number).ToList();
            int maxVerse = verses.Last().number;

            // Expand all ranges into a flat list of verse numbers
            var verseNumbers = ExpandRanges(parsed.Ranges, maxVerse);

            if (!verseNumbers.Any())
            {
                log.Warn("No valid verses resolved from reference.");
                MessageBox.Show("No valid verses found for this reference.", "Invalid Reference", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            log.Debug($"Final verse list: {string.Join(",", verseNumbers)}");

            new ScriptureManager().addScripture(
                cmbTemplate.SelectedItem as ScriptureTemplate,
                bible,
                book.name,
                parsed.Chapter,
                verseNumbers,
                chkMultiVerse.Checked);
        }

        private void InsertBulk()
        {
            var lines = SplitBulkInput(txtBulk.Text);
            int inserted = 0;
            var errors = new List<string>();

            foreach (var line in lines)
            {
                try
                {
                    var fullParsed = FullReferenceParser.ParseFullReference(bible, line);
                    if (fullParsed == null)
                    {
                        errors.Add($"Cannot parse: \"{line}\"");
                        continue;
                    }

                    var (bookName, refPart) = fullParsed.Value;
                    var book = bible.books.First(b => b.name.Equals(bookName, StringComparison.OrdinalIgnoreCase));

                    List<int> verseNumbers;
                    int chapterNum;

                    if (string.IsNullOrWhiteSpace(refPart))
                    {
                        // Whole book? We'll just do chapter 1 (user needs to be more specific)
                        errors.Add($"No chapter specified for: \"{line}\"");
                        continue;
                    }

                    var parsed = ScriptureReferenceParser.Parse(refPart);
                    chapterNum = parsed.Chapter;

                    var chapter = book.chapters.FirstOrDefault(c => c.number == chapterNum);
                    if (chapter == null)
                    {
                        errors.Add($"Chapter {chapterNum} not found in {bookName}: \"{line}\"");
                        continue;
                    }

                    int maxVerse = chapter.verses.OrderBy(v => v.number).Last().number;
                    verseNumbers = ExpandRanges(parsed.Ranges, maxVerse);

                    if (!verseNumbers.Any())
                    {
                        errors.Add($"No valid verses for: \"{line}\"");
                        continue;
                    }

                    new ScriptureManager().addScripture(
                        cmbTemplate.SelectedItem as ScriptureTemplate,
                        bible,
                        book.name,
                        chapterNum,
                        verseNumbers,
                        chkMultiVerse.Checked);

                    inserted++;
                }
                catch (Exception ex)
                {
                    errors.Add($"\"{line}\": {ex.Message}");
                }
            }

            if (errors.Any())
            {
                string msg = $"Inserted {inserted} reference(s).\n\n" +
                             $"The following could not be processed:\n" +
                             string.Join("\n", errors.Select(e => "\u2022 " + e));
                MessageBox.Show(msg, "Bulk Insert Results", MessageBoxButtons.OK,
                    errors.Count == lines.Count ? MessageBoxIcon.Error : MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Expands parsed ranges into a flat, sorted, deduplicated list of verse numbers.
        /// </summary>
        private List<int> ExpandRanges(List<(int Start, int End)> ranges, int maxVerse)
        {
            var verseNumbers = new List<int>();
            foreach (var range in ranges)
            {
                int s = Math.Max(1, range.Start);
                int end = range.End == int.MaxValue ? maxVerse : range.End;
                end = Math.Min(maxVerse, end);

                for (int v = s; v <= end; v++)
                    verseNumbers.Add(v);
            }
            return verseNumbers.Distinct().OrderBy(v => v).ToList();
        }

        // -----------------------------------------------------------------------
        // Other event handlers
        // -----------------------------------------------------------------------
        private void txtReference_TextChanged(object sender, EventArgs e)
        {
            if (!isBulkMode)
                btnInsert.Enabled = isValidReference();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void cmbTranslation_SelectionChangeCommitted(object sender, EventArgs e)
        {
            var box = (sender as ComboBox);
            var translationName = box.SelectedItem as string;
            log.Info($"Selecting translation: {translationName}");

            bible = OpenSongBibleReader.LoadTranslation(translationName);

            // Re-validate current book text against new translation
            if (!isBulkMode)
            {
                btnInsert.Enabled = isValidReference();
            }

            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            registryKey.SetValue("LastBibleTranslation", translationName);
        }

        private void cmbTemplate_SelectionChangeCommitted(object sender, EventArgs e)
        {
            var box = (sender as ComboBox);
            var template = box.SelectedItem as ScriptureTemplate;
            log.Info($"Selected template: {template.name}");
            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            registryKey.SetValue("LastScriptureTemplate", template.name);
        }

        private void chkMultiVerse_CheckedChanged(object sender, EventArgs e)
        {
            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            registryKey.SetValue("MultiVerseProjection", chkMultiVerse.Checked ? 1 : 0, RegistryValueKind.DWord);
            log.Debug($"MultiVerseProjection preference saved: {chkMultiVerse.Checked}");
        }

        /// <summary>
        /// Hide suggestion box when the form loses focus or is deactivated.
        /// </summary>
        protected override void OnDeactivate(EventArgs e)
        {
            base.OnDeactivate(e);
            if (_suggestionBox != null)
                _suggestionBox.Visible = false;
        }
    }

    public class ScriptureTemplate
    {
        public string name { get; }
        public string path { get; }

        public ScriptureTemplate(string path)
        {
            this.path = path;
            this.name = path.Split(new char[] { '\\' }).Last().Replace(".pptx", "");
        }

        public override string ToString()
        {
            return name;
        }
    }

    // -----------------------------------------------------------------------
    //  ParsedReference + ScriptureReferenceParser
    //  (chapter : verse reference parser — book name handled separately)
    // -----------------------------------------------------------------------
    public class ParsedReference
    {
        public int Chapter { get; set; }
        public List<(int Start, int End)> Ranges { get; set; } = new List<(int, int)>();
    }

    public static class ScriptureReferenceParser
    {
        /// <summary>
        /// Parses a chapter:verse reference string (without the book name).
        /// Supports many common formats:
        ///   "3"              → whole chapter 3
        ///   "3:"             → whole chapter 3
        ///   "3:16"           → chapter 3, verse 16
        ///   "3:16-18"        → chapter 3, verses 16–18
        ///   "3:16,17,18"     → chapter 3, verses 16, 17, 18
        ///   "3:16-18,20"     → chapter 3, verses 16–18 and 20
        ///   "3v16"  / "3V16" → treated as "3:16"
        ///   "3.16"           → treated as "3:16"
        ///   En-dash / em-dash are normalised to hyphen.
        ///   Spaces are stripped.
        /// </summary>
        public static ParsedReference Parse(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                throw new ArgumentException("Reference is empty.");

            // Normalize input
            input = input
                .Replace("\u2013", "-")   // en-dash
                .Replace("\u2014", "-")   // em-dash
                .Replace(" ", "")         // remove all spaces
                .Replace("v", ":")        // support "3v16"
                .Replace("V", ":")
                .Replace(".", ":");       // support "3.16"

            // Collapse repeated colons that may result from normalisation (e.g. "3..16" → "3::16")
            while (input.Contains("::"))
                input = input.Replace("::", ":");

            var result = new ParsedReference();

            var parts = input.Split(':');

            if (!int.TryParse(parts[0], out int chapter) || chapter <= 0)
                throw new FormatException("Invalid chapter in reference.");

            result.Chapter = chapter;

            // Whole chapter
            if (parts.Length == 1)
            {
                result.Ranges.Add((1, int.MaxValue));
                return result;
            }

            // Join everything after the first colon (handles pathological "3:16:17" → "16:17" → "16,17")
            var versePart = string.Join(",", parts.Skip(1));

            if (string.IsNullOrWhiteSpace(versePart))
            {
                // "3:" → whole chapter
                result.Ranges.Add((1, int.MaxValue));
                return result;
            }

            var segments = versePart.Split(',');

            foreach (var seg in segments)
            {
                if (string.IsNullOrWhiteSpace(seg))
                    continue;

                if (seg.Contains("-"))
                {
                    var r = seg.Split('-');
                    if (r.Length != 2)
                        throw new FormatException("Invalid verse range segment.");

                    if (!int.TryParse(r[0], out int start) || !int.TryParse(r[1], out int end))
                        throw new FormatException("Invalid verse numbers in range.");

                    if (end < start)
                        throw new FormatException("Verse range end is before start.");

                    result.Ranges.Add((start, end));
                }
                else
                {
                    if (!int.TryParse(seg, out int v))
                        throw new FormatException("Invalid verse number.");

                    result.Ranges.Add((v, v));
                }
            }

            if (!result.Ranges.Any())
                throw new FormatException("No valid verse ranges found.");

            return result;
        }
    }
}
