using log4net;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using static Microsoft.Office.Core.MsoTriState;

namespace WorshipHelperVSTO
{
    public class ScriptureManager
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        // -----------------------------------------------------------------------
        // Tuning constants
        // -----------------------------------------------------------------------
        int maxHeight = 340;

        // The minimum font size we will shrink to before giving up.
        const float MIN_FONT_SIZE = 8f;

        // -----------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------

        /// <summary>
        /// Inserts scripture slides for a simple contiguous range (legacy signature).
        /// </summary>
        public void addScripture(ScriptureTemplate template, Bible bible, string bookName,
                                 int chapterNum, int verseNumStart, int verseNumEnd,
                                 bool multiVerse = false)
        {
            // Build a flat, contiguous verse list and forward to the list-based overload.
            var verseNumbers = Enumerable.Range(verseNumStart, verseNumEnd - verseNumStart + 1).ToList();
            addScripture(template, bible, bookName, chapterNum, verseNumbers, multiVerse);
        }

        /// <summary>
        /// Inserts scripture slides for an arbitrary (possibly non-contiguous) list of verse numbers.
        /// This is the preferred entry-point used by the new UI.
        /// </summary>
        public void addScripture(ScriptureTemplate template, Bible bible, string bookName,
                                 int chapterNum, List<int> verseNumbers,
                                 bool multiVerse = false)
        {
            if (verseNumbers == null || !verseNumbers.Any())
                throw new ArgumentException("No verse numbers supplied.");

            verseNumbers = verseNumbers.Distinct().OrderBy(v => v).ToList();

            log.Debug($"Inserting scripture: {bookName} {chapterNum}:{string.Join(",", verseNumbers)} " +
                      $"({bible.name}) template={template.name}, multiVerse={multiVerse}");

            var chapter = bible.books
                               .First(b => b.name.Equals(bookName, StringComparison.OrdinalIgnoreCase))
                               .chapters.First(c => c.number == chapterNum);

            // Resolve to actual Verse objects (skip any that don't exist in data)
            var verseList = verseNumbers
                .Select(n => chapter.verses.FirstOrDefault(v => v.number == n))
                .Where(v => v != null)
                .ToList();

            if (!verseList.Any())
                throw new ArgumentException("None of the requested verses exist in this chapter.");

            string translation = bible.name;

            // Build a human-readable reference label
            string referenceLabel = BuildReferenceLabel(bookName, chapterNum, verseNumbers, chapter.verses.Count, translation);

            if (multiVerse)
                addScriptureMultiVerse(template, chapter, verseList, referenceLabel, translation);
            else
                addScriptureOneVersePerSlide(template, chapter, verseList, bookName, chapterNum, translation);
        }

        // -----------------------------------------------------------------------
        // MODE A: One verse per slide
        // -----------------------------------------------------------------------
        private void addScriptureOneVersePerSlide(ScriptureTemplate template, Chapter chapter,
                                                  List<Verse> verseList, string bookName,
                                                  int chapterNum, string translation)
        {
            Application app = Globals.ThisAddIn.Application;

            Presentation templatePresentation = app.Presentations.Open(template.path, msoTrue, msoFalse, msoFalse);
            var templateSlide = templatePresentation.Slides[1];

            var color1 = templateSlide.Shapes[2].TextFrame.TextRange.Font.Color.RGB;
            var color2 = templateSlide.Shapes[3].TextFrame.TextRange.Font.Color.RGB;
            var originalFontSize = templateSlide.Shapes[2].TextFrame.TextRange.Font.Size;

            // Insert in descending order at a fixed index so the final deck is in ascending order.
            var descending = verseList.OrderByDescending(v => v.number).ToList();
            int insertAt = new SelectionManager().GetNextSlideIndex();

            foreach (var verse in descending)
            {
                log.Debug($"Adding slide for verse {verse.number}");

                var reference = $"{bookName} {chapterNum}:{verse.number} ({translation})";
                var currentSlide = newSlideFromTemplate(templatePresentation, insertAt);

                currentSlide.Shapes[2].TextFrame.TextRange.Font.Color.RGB = color1;
                currentSlide.Shapes[3].TextFrame.TextRange.Font.Color.RGB = color2;

                var objBodyTextBox = currentSlide.Shapes[2];
                var objDescTextBox = currentSlide.Shapes[3];

                objBodyTextBox.TextFrame.TextRange.Font.Size = originalFontSize;
                objDescTextBox.TextFrame.TextRange.Text = reference;

                string verseText = "\u00AB" + verse.number + "\u00BB " + verse.text;
                objBodyTextBox.TextFrame.TextRange.Text = verseText;

                // Force layout refresh before checking height
                ForceLayoutRefresh(objBodyTextBox);

                while (objBodyTextBox.Height > maxHeight && objBodyTextBox.TextFrame.TextRange.Font.Size > MIN_FONT_SIZE)
                {
                    objBodyTextBox.TextFrame.TextRange.Font.Size -= 1;
                    ForceLayoutRefresh(objBodyTextBox);
                }

                ApplySuperscriptMarkers(objBodyTextBox, new List<Verse> { verse });
            }

            templatePresentation.Close();

            // Select all inserted slides
            int slideCount = verseList.Count;
            if (slideCount > 0)
            {
                try
                {
                    int[] slideIdxs = Enumerable.Range(insertAt, slideCount).ToArray();
                    app.ActivePresentation.Slides.Range(slideIdxs).Select();
                }
                catch (Exception ex)
                {
                    // FIX: During live presentation, slide selection may fail with
                    // "Slide (unknown)" errors because the selection model differs
                    // in slideshow mode. The slides were still inserted successfully.
                    log.Warn($"Could not select inserted slides (this is normal during live presentation): {ex.Message}");
                }
            }
        }

        // -----------------------------------------------------------------------
        // MODE B: Multi-verse projection — GREEDY PACKING (simple & reliable)
        // -----------------------------------------------------------------------
        // FIX: Removed the broken "even distribution" pass that was causing
        // slide duplication errors and index corruption. Now uses a single
        // greedy pass: pack as many verses as will fit on each slide.
        // -----------------------------------------------------------------------
        private void addScriptureMultiVerse(ScriptureTemplate template, Chapter chapter,
                                            List<Verse> verseList, string referenceLabel,
                                            string translation)
        {
            Application app = Globals.ThisAddIn.Application;

            Presentation templatePresentation = app.Presentations.Open(template.path, msoTrue, msoFalse, msoFalse);
            var templateSlide = templatePresentation.Slides[1];
            var color1 = templateSlide.Shapes[2].TextFrame.TextRange.Font.Color.RGB;
            var color2 = templateSlide.Shapes[3].TextFrame.TextRange.Font.Color.RGB;
            var originalFontSize = templateSlide.Shapes[2].TextFrame.TextRange.Font.Size;

            int insertAt = new SelectionManager().GetNextSlideIndex();

            // ---- Single greedy pass: pack verses into slides ----
            var slideBatches = new List<List<Verse>>();
            var currentBatch = new List<Verse>();

            // Create one measuring slide to probe text heights
            var measuringSlide = newSlideFromTemplate(templatePresentation, insertAt);
            measuringSlide.Shapes[2].TextFrame.TextRange.Font.Color.RGB = color1;
            measuringSlide.Shapes[3].TextFrame.TextRange.Font.Color.RGB = color2;
            var measureBox = measuringSlide.Shapes[2];
            measureBox.TextFrame.TextRange.Font.Size = originalFontSize;
            measureBox.TextFrame.TextRange.Text = "";

            string runningText = "";

            foreach (var verse in verseList)
            {
                string verseText = "\u00AB" + verse.number + "\u00BB " + verse.text + " ";
                string candidateText = runningText + verseText;
                
                measureBox.TextFrame.TextRange.Font.Size = originalFontSize;
                measureBox.TextFrame.TextRange.Text = candidateText;
                ForceLayoutRefresh(measureBox);

                if (measureBox.Height > maxHeight && runningText.Length > 0)
                {
                    // Overflow — commit current batch and start fresh
                    slideBatches.Add(new List<Verse>(currentBatch));
                    currentBatch.Clear();
                    runningText = "";

                    // Retry this verse on a fresh slide
                    runningText = verseText;
                    currentBatch.Add(verse);

                    measureBox.TextFrame.TextRange.Font.Size = originalFontSize;
                    measureBox.TextFrame.TextRange.Text = runningText;
                    ForceLayoutRefresh(measureBox);

                    // If single verse still overflows, shrink font
                    while (measureBox.Height > maxHeight && measureBox.TextFrame.TextRange.Font.Size > MIN_FONT_SIZE)
                    {
                        measureBox.TextFrame.TextRange.Font.Size -= 1;
                        ForceLayoutRefresh(measureBox);
                    }
                }
                else
                {
                    runningText = candidateText;
                    currentBatch.Add(verse);
                }
            }
            if (currentBatch.Any())
                slideBatches.Add(currentBatch);

            // Delete the temporary measuring slide
            measuringSlide.Delete();

            log.Debug($"Greedy packing: {slideBatches.Count} slide(s) for {verseList.Count} verse(s). " +
                      $"Distribution: {string.Join(", ", slideBatches.Select(g => g.Count))}");

            // ---- Create actual slides from greedy batches ----
            var startSlideIndex = insertAt;

            for (int s = 0; s < slideBatches.Count; s++)
            {
                // Always create from template to avoid Duplicate() index issues
                var slide = newSlideFromTemplate(templatePresentation, insertAt + s);

                slide.Shapes[2].TextFrame.TextRange.Font.Color.RGB = color1;
                slide.Shapes[3].TextFrame.TextRange.Font.Color.RGB = color2;

                var objBodyTextBox = slide.Shapes[2];
                var objDescTextBox = slide.Shapes[3];

                objBodyTextBox.TextFrame.TextRange.Font.Size = originalFontSize;
                objDescTextBox.TextFrame.TextRange.Text = referenceLabel;

                // Build the text for this slide
                string slideText = "";
                foreach (var verse in slideBatches[s])
                {
                    slideText += "\u00AB" + verse.number + "\u00BB " + verse.text + " ";
                }
                objBodyTextBox.TextFrame.TextRange.Text = slideText.TrimEnd();
                ForceLayoutRefresh(objBodyTextBox);

                // Shrink font if necessary
                while (objBodyTextBox.Height > maxHeight && objBodyTextBox.TextFrame.TextRange.Font.Size > MIN_FONT_SIZE)
                {
                    objBodyTextBox.TextFrame.TextRange.Font.Size -= 1;
                    ForceLayoutRefresh(objBodyTextBox);
                }

                // Apply superscript formatting to verse number markers
                ApplySuperscriptMarkers(objBodyTextBox, slideBatches[s]);
            }

            templatePresentation.Close();

            // Select all newly inserted slides
            int totalSlides = slideBatches.Count;
            if (totalSlides > 0)
            {
                try
                {
                    int[] slideIndexes = Enumerable.Range(startSlideIndex, totalSlides).ToArray();
                    log.Debug($"Selecting slides from {startSlideIndex} to {startSlideIndex + totalSlides - 1}");
                    app.ActivePresentation.Slides.Range(slideIndexes).Select();
                }
                catch (Exception ex)
                {
                    // FIX: During live presentation, slide selection may fail with
                    // "Slide (unknown)" errors. The slides were still inserted successfully.
                    log.Warn($"Could not select inserted slides (this is normal during live presentation): {ex.Message}");
                }
            }
        }

        // -----------------------------------------------------------------------
        // Helpers
        // -----------------------------------------------------------------------

        /// <summary>
        /// Uses guillemet markers «N» to format verse numbers as superscript,
        /// then removes the marker characters. Processes markers from RIGHT to LEFT
        /// so that index shifts from deletion don't affect earlier markers.
        /// </summary>
        private void ApplySuperscriptMarkers(Shape bodyTextBox, List<Verse> verses)
        {
            string text = bodyTextBox.TextFrame.TextRange.Text;

            // Collect all marker positions first (verse number + position)
            var markers = new List<(int Index, int Length, int VerseNum)>();

            foreach (var verse in verses)
            {
                string toFind = "\u00AB" + verse.number + "\u00BB";
                int pos = text.IndexOf(toFind);
                if (pos >= 0)
                {
                    markers.Add((pos, toFind.Length, verse.number));
                }
            }

            // Process from right to left so earlier indices are not affected
            foreach (var marker in markers.OrderByDescending(m => m.Index))
            {
                int oneBasedStart = marker.Index + 1;
                int len = marker.Length;
                string numStr = marker.VerseNum.ToString();

                // Make the whole marker superscript
                bodyTextBox.TextFrame.TextRange.Characters(oneBasedStart, len).Font.Superscript = msoTrue;

                // Delete trailing » (after making superscript, the character at the end)
                bodyTextBox.TextFrame.TextRange.Characters(oneBasedStart + len - 1, 1).Delete();

                // Delete leading «
                bodyTextBox.TextFrame.TextRange.Characters(oneBasedStart, 1).Delete();
            }
        }

        /// <summary>
        /// Forces PowerPoint to recalculate the text box layout so that
        /// Height is accurate. Without this, Height may return stale values.
        /// </summary>
        private void ForceLayoutRefresh(Shape textBox)
        {
            try
            {
                // Reading Width after setting Text forces PPT to re-layout
                var _ = textBox.TextFrame.TextRange.BoundWidth;
            }
            catch
            {
                // Fallback: just read Height (some versions don't expose BoundWidth)
                try { var _ = textBox.Height; } catch { }
            }
        }

        /// <summary>
        /// Builds a human-readable reference label like "John 3:16-18,20 (ESV)".
        /// Collapses contiguous runs into ranges.
        /// </summary>
        private string BuildReferenceLabel(string bookName, int chapterNum,
                                           List<int> verseNumbers, int totalVersesInChapter,
                                           string translation)
        {
            if (verseNumbers.Count == totalVersesInChapter &&
                verseNumbers.First() == 1 && verseNumbers.Last() == totalVersesInChapter)
            {
                // Whole chapter
                return $"{bookName} {chapterNum} ({translation})";
            }

            // Collapse contiguous runs
            var ranges = new List<string>();
            int start = verseNumbers[0];
            int prev = start;
            for (int i = 1; i < verseNumbers.Count; i++)
            {
                if (verseNumbers[i] == prev + 1)
                {
                    prev = verseNumbers[i];
                }
                else
                {
                    ranges.Add(start == prev ? $"{start}" : $"{start}-{prev}");
                    start = verseNumbers[i];
                    prev = start;
                }
            }
            ranges.Add(start == prev ? $"{start}" : $"{start}-{prev}");

            string verseRef = string.Join(",", ranges);
            return $"{bookName} {chapterNum}:{verseRef} ({translation})";
        }

        private Slide newSlideFromTemplate(Presentation templatePresentation, int insertAt)
        {
            Application app = Globals.ThisAddIn.Application;
            templatePresentation.Slides[1].Copy();
            return app.ActivePresentation.Slides.Paste(insertAt)[1];
        }

        public static DocumentWindow getMainWindow()
        {
            Application app = Globals.ThisAddIn.Application;
            foreach (DocumentWindow win in app.ActivePresentation.Windows)
            {
                if (!win.Caption.Contains("Presenter View"))
                    return win;
            }
            return null;
        }
    }
}
