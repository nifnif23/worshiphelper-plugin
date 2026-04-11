using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;

namespace WorshipHelperVSTO
{
    public class SelectionManager
    {
        Application app = Globals.ThisAddIn.Application;
        WindowManager windowManager = new WindowManager();

        public int GetNextSlideIndex()
        {
            // If there are no slides, it is 1
            if (app.ActivePresentation.Slides.Count == 0)
            {
                Debug.WriteLine("There are no slides, so insert index is 1");
                return 1;
            }

            // If in presentation mode, it is the index of the slide after the one currently shown
            if (app.SlideShowWindows.Count > 0)
            {
                var index = app.ActivePresentation.SlideShowWindow.View.Slide.SlideIndex + 1;
                Debug.WriteLine($"We are presenting; insert index is {index}");
                return index;
            }

            var window = windowManager.GetMainWindow(); // i.e. not the presenter view

            // Guard: if no main window could be found, append at end
            if (window == null)
            {
                Debug.WriteLine("No main window found; appending at end");
                return app.ActivePresentation.Slides.Count + 1;
            }

            // If in edit mode, and there is a selection, it is the end of the selection
            if (window.Selection.Type == PpSelectionType.ppSelectionSlides)
            {
                var index = getLastSelectedIndex(window.Selection.SlideRange) + 1;
                Debug.WriteLine($"There is an active selection; insert index is {index}");
                return index;
            }

            // If there is no selection, toggle the view to force a selection
            Debug.WriteLine("There is no active selection; toggling view mode");
            toggleViewMode();

            // After toggling, re-check that we have a slide selection
            if (window.Selection.Type == PpSelectionType.ppSelectionSlides)
            {
                var index = window.Selection.SlideRange.SlideIndex + 1;
                Debug.WriteLine($"Insert index is {index}");
                return index;
            }

            // Fallback: append at end
            Debug.WriteLine("Still no selection after toggle; appending at end");
            return app.ActivePresentation.Slides.Count + 1;
        }

        public void GoToSlide(int index)
        {
            var window = windowManager.GetMainWindow();
            if (window != null)
            {
                window.View.GotoSlide(index);
            }
        }

        private int getLastSelectedIndex(SlideRange range)
        {
            int index = -1;
            foreach(Slide slide in range)
            {
                if (slide.SlideIndex > index) index = slide.SlideIndex;
            }
            return index;
        }

        private void toggleViewMode()
        {
            var activeWindow = app.ActiveWindow;
            activeWindow.ViewType = PpViewType.ppViewSlide;
            activeWindow.ViewType = PpViewType.ppViewNormal;
        }
    }
}
