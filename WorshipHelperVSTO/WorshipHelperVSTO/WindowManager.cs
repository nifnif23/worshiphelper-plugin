using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Diagnostics;

namespace WorshipHelperVSTO
{
    public class WindowManager
    {
        Application app = Globals.ThisAddIn.Application;

        public DocumentWindow GetMainWindow()
        {
            try
            {
                if (app.Presentations.Count == 0)
                {
                    return null;
                }

                foreach (DocumentWindow win in app.ActivePresentation.Windows)
                {
                    try
                    {
                        // There is probably a better way...
                        if (!win.Caption.Contains("Presenter View"))
                        {
                            return win;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error reading window caption: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error enumerating windows in GetMainWindow: {ex.Message}");
            }
            return null;
        }

        public DocumentWindow GetPresenterView()
        {
            try
            {
                if (app.Presentations.Count == 0)
                {
                    return null;
                }

                foreach (DocumentWindow win in app.ActivePresentation.Windows)
                {
                    try
                    {
                        // There is probably a better way...
                        if (win.Caption.Contains("Presenter View"))
                        {
                            return win;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error reading window caption: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error enumerating windows in GetPresenterView: {ex.Message}");
            }
            return null;
        }
    }
}
