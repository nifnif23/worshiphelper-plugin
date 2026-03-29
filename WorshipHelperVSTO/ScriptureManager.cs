using log4net;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Linq;
using static Microsoft.Office.Core.MsoTriState;

namespace WorshipHelperVSTO
{
    class ScriptureManager
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        int maxHeight = 400;

        public void addScripture(ScriptureTemplate template, Bible bible, string bookName, int chapterNum, int verseNumStart, int verseNumEnd)
        {
            log.Debug($"Inserting scripture from {bookName} {chapterNum}:{verseNumStart}-{verseNumEnd} ({bible.name}) using template {template.name}");

            Application app = Globals.ThisAddIn.Application;

            // Open the template presentation once; we copy from it for every verse
            Presentation templatePresentation = app.Presentations.Open(template.path, msoTrue, msoFalse, msoFalse);
            var templateSlide = templatePresentation.Slides[1];

            // Capture template colours so they survive the paste operation
            var color1 = templateSlide.Shapes[2].TextFrame.TextRange.Font.Color.RGB;
            var color2 = templateSlide.Shapes[3].TextFrame.TextRange.Font.Color.RGB;
            var originalFontSize = templateSlide.Shapes[2].TextFrame.TextRange.Font.Size;

            var translation = bible.name;
            var chapter = bible.books
                               .Where(item => item.name == bookName).First()
                               .chapters.Where(item => item.number == chapterNum).First();

            var verseList = chapter.verses
                                   .Where(verse => verse.number >= verseNumStart && verse.number <= verseNumEnd)
                                   .OrderBy(verse => verse.number)
                                   .ToList();

            int startSlideIndex = -1;
            int lastSlideIndex  = -1;

            // ── One slide per verse ──────────────────────────────────────────────
            foreach (var verse in verseList)
            {
                log.Debug($"Adding slide for verse {verse.number}");

                // Individual verse reference label (e.g. "John 3:16 (ESV)")
                var reference = $"{bookName} {chapterNum}:{verse.number} ({translation})";

                // Paste a fresh copy of the template slide at the current insert position
                var currentSlide = newSlideFromTemplate(templatePresentation);

                if (startSlideIndex == -1) startSlideIndex = currentSlide.SlideIndex;
                lastSlideIndex = currentSlide.SlideIndex;

                // Restore template colours (may be overridden by destination theme)
                currentSlide.Shapes[2].TextFrame.TextRange.Font.Color.RGB = color1;
                currentSlide.Shapes[3].TextFrame.TextRange.Font.Color.RGB = color2;

                var objBodyTextBox = currentSlide.Shapes[2];
                var objDescTextBox = currentSlide.Shapes[3];

                // Reset font size in case a previous verse shrank it
                objBodyTextBox.TextFrame.TextRange.Font.Size = originalFontSize;
                objDescTextBox.TextFrame.TextRange.Text = reference;

                // Write verse text; wrap verse number in $ markers for superscripting below
                string verseText = "$" + verse.number + "$ " + verse.text;
                objBodyTextBox.TextFrame.TextRange.Text = verseText;

                // If even a single verse is too long, shrink the font to fit
                while (objBodyTextBox.Height > maxHeight && objBodyTextBox.TextFrame.TextRange.Font.Size > 8)
                {
                    objBodyTextBox.TextFrame.TextRange.Font.Size -= 1;
                }

                // Superscript the verse number and strip the $ delimiters
                string toFind    = "$" + verse.number + "$";
                int    markerIdx = objBodyTextBox.TextFrame.TextRange.Text.IndexOf(toFind);
                if (markerIdx > -1)
                {
                    // Superscript the entire "$N$" token first
                    objBodyTextBox.TextFrame.TextRange.Characters(markerIdx + 1, toFind.Length).Font.Superscript = msoTrue;
                    // Delete leading $  (indices shift by -1 after this)
                    objBodyTextBox.TextFrame.TextRange.Characters(markerIdx + 1, 1).Delete();
                    // Delete trailing $ (now sits at markerIdx + toFind.Length - 1)
                    objBodyTextBox.TextFrame.TextRange.Characters(markerIdx + toFind.Length - 1, 1).Delete();
                }

                // Select this slide so the next newSlideFromTemplate inserts after it
                app.ActivePresentation.Slides.Range(new int[] { lastSlideIndex }).Select();
            }

            templatePresentation.Close();

            // Re-select all newly added slides so a future addition goes after them
            if (startSlideIndex != -1)
            {
                int numSlides   = lastSlideIndex - startSlideIndex + 1;
                int[] slideIdxs = new int[numSlides];
                for (int i = 0; i < numSlides; i++)
                    slideIdxs[i] = i + startSlideIndex;

                log.Debug($"Selecting slides {startSlideIndex} to {lastSlideIndex}");
                app.ActivePresentation.Slides.Range(slideIdxs).Select();
            }
        }

        private Slide newSlideFromTemplate(Presentation templatePresentation)
        {
            Application app = Globals.ThisAddIn.Application;

            var insertAt = new SelectionManager().GetNextSlideIndex();
            log.Debug($"Pasting template slide at position {insertAt}");
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
