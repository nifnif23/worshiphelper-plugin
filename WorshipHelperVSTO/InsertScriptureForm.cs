using log4net;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
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

            var source = new AutoCompleteStringCollection();
            log.Debug("Adding books");
            source.AddRange(bible.books.Select(book => book.name).ToArray());
            txtBook.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtBook.AutoCompleteSource = AutoCompleteSource.CustomSource;
            txtBook.AutoCompleteCustomSource = source;

            btnInsert.Enabled = false;
        }

        private void txtSearchBox_TextChanged(object sender, EventArgs e)
        {
            btnInsert.Enabled = isValidReference();
        }

        private void txtSearchBox_KeyPress(object sender, KeyPressEventArgs e)
        {
        }

        private bool isValidReference()
        {
            log.Debug($"Checking reference validity (book: {txtBook.Text}, reference: {txtReference.Text})");

            if (string.IsNullOrWhiteSpace(txtBook.Text) || bible == null)
                return false;

            var bookNames = bible.books.Select(book => book.name.ToLower()).ToList();
            var validBook = bookNames.Contains(txtBook.Text.ToLower());

            if (!validBook)
                return false;

            try
            {
                // Use the same parser as insertion
                ScriptureReferenceParser.Parse(txtReference.Text);
                return true;
            }
            catch (Exception ex)
            {
                log.Debug($"Reference parse failed: {ex.Message}");
                return false;
            }
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            log.Info("About to insert scripture");

            var book = bible.books
                .First(b => b.name.Equals(txtBook.Text, StringComparison.OrdinalIgnoreCase));

            // Parse reference using the universal parser
            var parsed = ScriptureReferenceParser.Parse(txtReference.Text);

            log.Debug($"Parsed reference: chapter={parsed.Chapter}, ranges={string.Join(";", parsed.Ranges.Select(r => $"{r.Start}-{r.End}"))}");

            var chapter = book.chapters
                .First(c => c.number == parsed.Chapter);

            // Ensure verses are sorted numerically
            var verses = chapter.verses
                .OrderBy(v => v.number)
                .ToList();

            int maxVerse = verses.Last().number;

            // Expand all ranges into a flat list of verse numbers
            var verseNumbers = new List<int>();

            foreach (var range in parsed.Ranges)
            {
                int s = Math.Max(1, range.Start);
                int e = range.End == int.MaxValue ? maxVerse : range.End;
                e = Math.Min(maxVerse, e);

                for (int v = s; v <= e; v++)
                    verseNumbers.Add(v);
            }

            // Remove duplicates and sort
            verseNumbers = verseNumbers.Distinct().OrderBy(v => v).ToList();

            if (!verseNumbers.Any())
            {
                log.Warn("No valid verses resolved from reference.");
                MessageBox.Show("No valid verses found for this reference.", "Invalid Reference", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            log.Debug($"Final verse list: {string.Join(",", verseNumbers)}");

            try
            {
                new ScriptureManager().addScripture(
                    cmbTemplate.SelectedItem as ScriptureTemplate,
                    bible,
                    book.name,
                    parsed.Chapter,
                    verseNumbers.First(),
                    verseNumbers.Last(),
                    chkMultiVerse.Checked);

                log.Debug("Insert complete");
            }
            finally
            {
                log.Debug("Closing scripture window");
                this.Close();
            }
        }

        private void txtReference_TextChanged(object sender, EventArgs e)
        {
            btnInsert.Enabled = isValidReference();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmbTranslation_SelectionChangeCommitted(object sender, EventArgs e)
        {
            var box = (sender as ComboBox);
            var translationName = box.SelectedItem as string;
            log.Info($"Selecting translation: {translationName}");

            bible = OpenSongBibleReader.LoadTranslation(translationName);

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

    public class ParsedReference
    {
        public int Chapter { get; set; }
        public List<(int Start, int End)> Ranges { get; set; } = new();
    }

    public static class ScriptureReferenceParser
    {
        public static ParsedReference Parse(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                throw new ArgumentException("Reference is empty.");

            // Normalize input
            input = input
                .Replace("–", "-")   // en-dash
                .Replace("—", "-")   // em-dash
                .Replace(" ", "")    // remove all spaces
                .Replace("v", ":")   // support "3v16"
                .Replace("V", ":");

            // Now formats like "3v16-18" become "3:16-18"

            var result = new ParsedReference();

            var parts = input.Split(':');

            if (!int.TryParse(parts[0], out int chapter))
                throw new FormatException("Invalid chapter in reference.");

            result.Chapter = chapter;

            // Whole chapter
            if (parts.Length == 1)
            {
                result.Ranges.Add((1, int.MaxValue));
                return result;
            }

            if (parts.Length > 2)
                throw new FormatException("Too many ':' or 'v' separators in reference.");

            var versePart = parts[1];

            if (string.IsNullOrWhiteSpace(versePart))
            {
                // Treat as whole chapter
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