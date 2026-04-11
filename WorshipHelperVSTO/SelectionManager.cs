using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Diagnostics;

namespace WorshipHelperVSTO
{
    public class SelectionManager
    {
        Application app = Globals.ThisAddIn.Application;
        WindowManager windowManager = new WindowManager();

        public int GetNextSlideIndex()
        {
            try
            {
                // If there are no slides, it is 1
                if (app.ActivePresentation.Slides.Count == 0)
                {
                    Debug.WriteLine("There are no slides, so insert index is 1");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error checking slide count: {ex.Message}");
                return 1;
            }

            // If in presentation mode, it is the index of the slide after the one currently shown
            try
            {
                if (app.SlideShowWindows.Count > 0)
                {
                    try
                    {
                        var slideShowWindow = app.ActivePresentation.SlideShowWindow;
                        if (slideShowWindow != null && slideShowWindow.View != null)
                        {
                            var currentSlide = slideShowWindow.View.Slide;
                            if (currentSlide != null)
                            {
                                var index = currentSlide.SlideIndex + 1;
                                Debug.WriteLine($"We are presenting; insert index is {index}");
                                return index;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // FIX: The SlideShowWindow.View.Slide can throw when:
                        // - Between slide transitions
                        // - At the "End of slide show" black screen
                        // - When the slideshow is in an intermediate state
                        // Fall through to use the slide count as fallback
                        Debug.WriteLine($"Error accessing slide show view (slide unknown): {ex.Message}");

                        // Best fallback during presentation: append at end
                        int fallbackIndex = app.ActivePresentation.Slides.Count + 1;
                        Debug.WriteLine($"Falling back to end of presentation; insert index is {fallbackIndex}");
                        return fallbackIndex;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error checking SlideShowWindows: {ex.Message}");
                // Continue to try edit-mode logic
            }

            DocumentWindow window = null;
            try
            {
                window = windowManager.GetMainWindow(); // i.e. not the presenter view
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting main window: {ex.Message}");
            }

            // Guard: if no main window could be found, append at end
            if (window == null)
            {
                Debug.WriteLine("No main window found; appending at end");
                return app.ActivePresentation.Slides.Count + 1;
            }

            // If in edit mode, and there is a selection, it is the end of the selection
            try
            {
                var selectionType = window.Selection.Type;
                if (selectionType == PpSelectionType.ppSelectionSlides)
                {
                    var index = getLastSelectedIndex(window.Selection.SlideRange) + 1;
                    Debug.WriteLine($"There is an active selection; insert index is {index}");
                    return index;
                }
            }
            catch (Exception ex)
            {
                // FIX: Selection.Type can throw if the window is in an unexpected state
                // (e.g., during presentation mode, or if no view is active)
                Debug.WriteLine($"Error checking selection type: {ex.Message}");
            }

            // If there is no selection, toggle the view to force a selection
            Debug.WriteLine("There is no active selection; toggling view mode");
            try
            {
                toggleViewMode();

                // After toggling, re-check that we have a slide selection
                if (window.Selection.Type == PpSelectionType.ppSelectionSlides)
                {
                    var index = window.Selection.SlideRange.SlideIndex + 1;
                    Debug.WriteLine($"Insert index is {index}");
                    return index;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error after toggling view mode: {ex.Message}");
            }

            // Fallback: append at end
            Debug.WriteLine("Still no selection after toggle; appending at end");
            try
            {
                return app.ActivePresentation.Slides.Count + 1;
            }
            catch
            {
                return 1;
            }
        }

        public void GoToSlide(int index)
        {
            try
            {
                var window = windowManager.GetMainWindow();
                if (window != null)
                {
                    window.View.GotoSlide(index);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error navigating to slide {index}: {ex.Message}");
            }
        }

        private int getLastSelectedIndex(SlideRange range)
        {
            int index = -1;
            try
            {
                foreach (Slide slide in range)
                {
                    if (slide.SlideIndex > index) index = slide.SlideIndex;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error iterating SlideRange: {ex.Message}");
            }
            return index;
        }

        private void toggleViewMode()
        {
            try
            {
                var activeWindow = app.ActiveWindow;
                if (activeWindow != null)
                {
                    activeWindow.ViewType = PpViewType.ppViewSlide;
                    activeWindow.ViewType = PpViewType.ppViewNormal;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error toggling view mode: {ex.Message}");
            }
        }
    }
}
