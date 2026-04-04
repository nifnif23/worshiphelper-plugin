using log4net;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using static Microsoft.Office.Core.MsoTriState;

namespace WorshipHelperVSTO
{
    class ScriptureManager
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        int maxHeight = 400;

        /// <summary>
        /// Inserts scripture slides into the active PowerPoint presentation.
        /// </summary>
        /// <param name="multiVerse">
        /// When true, verses are packed together onto as few slides as possible (multi-verse
        /// projection mode). When false, each verse gets its own dedicated slide.
        /// </param>
        public void addScripture(ScriptureTemplate template, Bible bible, string bookName,
                                 int chapterNum, int verseNumStart, int verseNumEnd,
                                 bool multiVerse = false)
        {
            log.Debug($"Inserting scripture from {bookName} {chapterNum}:{verseNumStart}-{verseNumEnd} " +
                      $"({bible.name}) using template {template.name}, multiVerse={multiVerse}");

            if (multiVerse)
                addScriptureMultiVerse(template, bible, bookName, chapterNum, verseNumStart, verseNumEnd);
            else
                addScriptureOneVersePerSlide(template, bible, bookName, chapterNum, verseNumStart, verseNumEnd);
        }

        // -----------------------------------------------------------------------
        // MODE A: One verse per slide
        // Inserts in descending verse order at a fixed index so the final deck
        // ends up in ascending order (1, 2, 3 ...).
        // -----------------------------------------------------------------------
        private void addScriptureOneVersePerSlide(ScriptureTemplate template, Bible bible,
                                                  string bookName, int chapterNum,
                                                  int verseNumStart, int verseNumEnd)
        {
            Application app = Globals.ThisAddIn.Application;

            Presentation templatePresentation = app.Presentations.Open(template.path, msoTrue, msoFalse, msoFalse);
            var templateSlide = templatePresentation.Slides[1];

            var color1 = templateSlide.Shapes[2].TextFrame.TextRange.Font.Color.RGB;
            var color2 = templateSlide.Shapes[3].TextFrame.TextRange.Font.Color.RGB;
            var originalFontSize = templateSlide.Shapes[2].TextFrame.TextRange.Font.Size;

            var translation = bible.name;
            var chapter = bible.books
                               .Where(item => item.name == bookName).First()
                               .chapters.Where(item => item.number == chapterNum).First();

            // Order descending so repeated inserts at the same index end up in correct final order
            var verseList = chapter.verses
                                   .Where(verse => verse.number >= verseNumStart && verse.number <= verseNumEnd)
                                   .OrderByDescending(verse => verse.number)
                                   .ToList();

            // Calculate insertAt ONCE before the loop so it stays fixed.
            int insertAt = new SelectionManager().GetNextSlideIndex();

            foreach (var verse in verseList)
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

                string verseText = "$" + verse.number + "$ " + verse.text;
                objBodyTextBox.TextFrame.TextRange.Text = verseText;

                while (objBodyTextBox.Height > maxHeight && objBodyTextBox.TextFrame.TextRange.Font.Size > 8)
                {
                    objBodyTextBox.TextFrame.TextRange.Font.Size -= 1;
                }

                string toFind    = "$" + verse.number + "$";
                int    markerIdx = objBodyTextBox.TextFrame.TextRange.Text.IndexOf(toFind);
                if (markerIdx > -1)
                {
                    objBodyTextBox.TextFrame.TextRange.Characters(markerIdx + 1, toFind.Length).Font.Superscript = msoTrue;
                    objBodyTextBox.TextFrame.TextRange.Characters(markerIdx + 1, 1).Delete();
                    objBodyTextBox.TextFrame.TextRange.Characters(markerIdx + toFind.Length - 1, 1).Delete();
                }
            }

            templatePresentation.Close();

            // Select all inserted slides
            int slideCount = verseList.Count;
            if (slideCount > 0)
            {
                int[] slideIdxs = Enumerable.Range(insertAt, slideCount).ToArray();
                app.ActivePresentation.Slides.Range(slideIdxs).Select();
            }
        }

        // -----------------------------------------------------------------------
        // MODE B: Multi-verse projection
        // Packs as many verses as will fit onto a single slide, then overflows
        // to a new duplicate slide when the text box exceeds maxHeight.
        // The reference label uses a compact format (e.g. "John 3:16-18 (ESV)").
        // -----------------------------------------------------------------------
        private void addScriptureMultiVerse(ScriptureTemplate template, Bible bible,
                                            string bookName, int chapterNum,
                                            int verseNumStart, int verseNumEnd)
        {
            int verseCount = verseNumEnd - verseNumStart + 1;

            Application app = Globals.ThisAddIn.Application;

            // Copy the template slide and close the template presentation
            Presentation templatePresentation = app.Presentations.Open(template.path, msoTrue, msoFalse, msoFalse);
            var currentSlide = newSlideFromTemplate(templatePresentation, new SelectionManager().GetNextSlideIndex());

            // Explicitly restore text colours (PasteSourceFormatting is async)
            var templateSlide = templatePresentation.Slides[1];
            var color1 = templateSlide.Shapes[2].TextFrame.TextRange.Font.Color.RGB;
            var color2 = templateSlide.Shapes[3].TextFrame.TextRange.Font.Color.RGB;
            currentSlide.Shapes[2].TextFrame.TextRange.Font.Color.RGB = color1;
            currentSlide.Shapes[3].TextFrame.TextRange.Font.Color.RGB = color2;

            templatePresentation.Close();

            var objBodyTextBox = currentSlide.Shapes[2];
            var objDescTextBox = currentSlide.Shapes[3];
            var originalFontSize = objBodyTextBox.TextFrame.TextRange.Font.Size;

            var translation = bible.name;
            var chapter = bible.books.Where(item => item.name == bookName).First()
                                     .chapters.Where(item => item.number == chapterNum).First();

            var verseList = chapter.verses
                                   .Where(verse => verse.number >= verseNumStart && verse.number <= verseNumEnd)
                                   .OrderBy(verse => verse.number)
                                   .ToList();

            // Build a compact reference label for the whole passage
            string verseReference;
            if (verseNumStart == 1 && verseNumEnd == chapter.verses.Count)
                verseReference = "";
            else if (verseNumStart == verseNumEnd)
                verseReference = $":{verseNumStart}";
            else
                verseReference = $":{verseNumStart}-{verseNumEnd}";
            var reference = $"{bookName} {chapterNum}{verseReference} ({translation})";

            objBodyTextBox.TextFrame.TextRange.Text = "";
            objDescTextBox.TextFrame.TextRange.Text = reference;

            var startSlideIndex = currentSlide.SlideIndex;
            var numSlidesAdded = 0;

            for (int i = 0; i < verseCount; i++)
            {
                log.Debug($"Adding verse {verseList[i].number}");
                var originalText = objBodyTextBox.TextFrame.TextRange.Text;
                var verseText = "$" + verseList[i].number + "$ " + verseList[i].text + " ";
                objBodyTextBox.TextFrame.TextRange.Text = originalText + verseText;

                if (objBodyTextBox.Height > maxHeight)
                {
                    if (originalText == "")
                    {
                        // Single verse is too long for the slide — shrink font until it fits
                        while (objBodyTextBox.Height > maxHeight && objBodyTextBox.TextFrame.TextRange.Font.Size > 8)
                        {
                            objBodyTextBox.TextFrame.TextRange.Font.Size -= 1;
                        }
                    }
                    else
                    {
                        log.Debug("Adding new slide");

                        // Undo the extra text and move to a fresh duplicate slide
                        objBodyTextBox.TextFrame.TextRange.Text = originalText;

                        currentSlide = currentSlide.Duplicate()[1];
                        numSlidesAdded++;
                        objBodyTextBox = currentSlide.Shapes[2];
                        objDescTextBox = currentSlide.Shapes[3];

                        objBodyTextBox.TextFrame.TextRange.Font.Size = originalFontSize;
                        objBodyTextBox.TextFrame.TextRange.Text = "";
                        objDescTextBox.TextFrame.TextRange.Text = reference;

                        i--; // retry this verse on the new slide
                    }
                }
            }

            var endSlideIndex = startSlideIndex + numSlidesAdded;

            // Superscript all verse number markers ($N$) across every produced slide
            for (int slideIndex = startSlideIndex; slideIndex <= endSlideIndex; slideIndex++)
            {
                currentSlide = app.ActivePresentation.Slides[slideIndex];
                objBodyTextBox = currentSlide.Shapes[2];
                foreach (Verse verse in verseList)
                {
                    string toFind = "$" + verse.number + "$";
                    int verseIndex = objBodyTextBox.TextFrame.TextRange.Text.IndexOf(toFind);
                    if (verseIndex > -1)
                    {
                        objBodyTextBox.TextFrame.TextRange.Characters(verseIndex + 1, toFind.Length).Font.Superscript = msoTrue;
                        objBodyTextBox.TextFrame.TextRange.Characters(verseIndex + 1, 1).Delete();
                        objBodyTextBox.TextFrame.TextRange.Characters(verseIndex + toFind.Length - 1, 1).Delete();
                    }
                }
            }

            // Select the newly inserted slides
            int[] slideIndexes = new int[numSlidesAdded + 1];
            for (int i = 0; i < numSlidesAdded + 1; i++)
            {
                slideIndexes[i] = i + startSlideIndex;
            }
            log.Debug($"Selecting slides from {startSlideIndex} to {endSlideIndex}");
            app.ActivePresentation.Slides.Range(slideIndexes).Select();
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
