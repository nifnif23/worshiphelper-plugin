using log4net;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic; // Added for list manipulation
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

            Presentation templatePresentation = app.Presentations.Open(template.path, msoTrue, msoFalse, msoFalse);
            var templateSlide = templatePresentation.Slides[1];

            var color1 = templateSlide.Shapes[2].TextFrame.TextRange.Font.Color.RGB;
            var color2 = templateSlide.Shapes[3].TextFrame.TextRange.Font.Color.RGB;
            var originalFontSize = templateSlide.Shapes[2].TextFrame.TextRange.Font.Size;

            var translation = bible.name;
            var chapter = bible.books
                               .Where(item => item.name == bookName).First()
                               .chapters.Where(item => item.number == chapterNum).First();

            // --- CHANGE: Reverse the order here ---
            // We order descending so that the last verse is processed first.
            // This way, when PowerPoint pastes at the selection, the final result is 1, 2, 3...
            var verseList = chapter.verses
                                   .Where(verse => verse.number >= verseNumStart && verse.number <= verseNumEnd)
                                   .OrderByDescending(verse => verse.number) 
                                   .ToList();

            int startSlideIndex = -1;
            int lastSlideIndex  = -1;

            foreach (var verse in verseList)
            {
                log.Debug($"Adding slide for verse {verse.number}");

                var reference = $"{bookName} {chapterNum}:{verse.number} ({translation})";
                var currentSlide = newSlideFromTemplate(templatePresentation);

                // Track indices for the final selection highlight
                if (startSlideIndex == -1) startSlideIndex = currentSlide.SlideIndex;
                lastSlideIndex = currentSlide.SlideIndex;

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

                // We don't necessarily need to select every time in a loop if the paste 
                // location is handled by newSlideFromTemplate, but keeping your logic:
                app.ActivePresentation.Slides.Range(new int[] { lastSlideIndex }).Select();
            }

            templatePresentation.Close();

            // Re-select all added slides (Small logic tweak: ensure order is handled correctly)
            if (startSlideIndex != -1)
            {
                int min = Math.Min(startSlideIndex, lastSlideIndex);
                int max = Math.Max(startSlideIndex, lastSlideIndex);
                int numSlides = max - min + 1;
                int[] slideIdxs = new int[numSlides];
                for (int i = 0; i < numSlides; i++)
                    slideIdxs[i] = i + min;

                app.ActivePresentation.Slides.Range(slideIdxs).Select();
            }
        }

        private Slide newSlideFromTemplate(Presentation templatePresentation)
        {
            Application app = Globals.ThisAddIn.Application;
            var insertAt = new SelectionManager().GetNextSlideIndex();
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
