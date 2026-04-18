using log4net;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WorshipHelperVSTO
{
    public partial class ThisAddIn
    {
        private static ILog log;
        public static String appDataPath;
        public static String userDataPath = $@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\WorshipHelper";

        // NOTE: We need a backing field to prevent the delegate being garbage collected
        private SafeNativeMethods.HookProc _keyboardProc;

        private IntPtr _hookIdKeyboard;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _keyboardProc = KeyboardHookCallback;
            log = LogManager.GetLogger("WorshipHelperVSTO");
            log.Info("Initalised logger");
            SetWindowsHooks();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            UnhookWindowsHooks();
        }

        private void SetWindowsHooks()
        {
            // FIX: WH_KEYBOARD is thread-local — it only fires on the thread that
            // installed it, which is the VSTO main thread. When PowerPoint's Presenter
            // View is active, keystrokes are dispatched on a DIFFERENT thread (the
            // presenter view window's message loop), so the old hook never fired.
            //
            // WH_KEYBOARD_LL (low-level keyboard hook) is system-wide and fires
            // regardless of which window or thread has focus. It requires a valid
            // module handle and threadId = 0 (global). The callback receives a
            // KBDLLHOOKSTRUCT* via lParam instead of the old packed lParam value.
            IntPtr hMod = SafeNativeMethods.GetModuleHandle(
                System.IO.Path.GetFileName(
                    System.Reflection.Assembly.GetExecutingAssembly().Location));

            _hookIdKeyboard =
                SafeNativeMethods.SetWindowsHookEx(
                    (int)SafeNativeMethods.HookType.WH_KEYBOARD_LL,
                    _keyboardProc,
                    hMod,
                    0); // 0 = global (all threads)

            if (_hookIdKeyboard == IntPtr.Zero)
            {
                int err = Marshal.GetLastWin32Error();
                log.Warn($"SetWindowsHookEx (WH_KEYBOARD_LL) failed with error {err}. " +
                         "Ctrl shortcut may not work in Presenter View.");
            }
        }

        private void UnhookWindowsHooks()
        {
            SafeNativeMethods.UnhookWindowsHookEx(_hookIdKeyboard);
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0)
                {
                    // WH_KEYBOARD_LL: wParam is the message type (WM_KEYDOWN/WM_KEYUP/etc.)
                    // lParam is a pointer to a KBDLLHOOKSTRUCT.
                    uint msg = (uint)wParam.ToInt64();
                    bool isKeyUp = (msg == (uint)SafeNativeMethods.WindowMessages.WM_KEYUP ||
                                   msg == (uint)SafeNativeMethods.WindowMessages.WM_SYSKEYUP);

                    if (!isKeyUp)
                    {
                        return SafeNativeMethods.CallNextHookEx(_hookIdKeyboard, nCode, wParam, lParam);
                    }

                    // Read the virtual key code from the KBDLLHOOKSTRUCT
                    var hookStruct = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                    int vkCode = hookStruct.vkCode;

                    bool formOpen = System.Windows.Forms.Application.OpenForms.Count > 0;
                    if (formOpen)
                    {
                        log.Debug("Ignoring key press while form open");
                        return SafeNativeMethods.CallNextHookEx(_hookIdKeyboard, nCode, wParam, lParam);
                    }

                    DocumentWindow presenterView = new WindowManager().GetPresenterView();
                    bool presenting = presenterView != null &&
                                      presenterView.Active == Microsoft.Office.Core.MsoTriState.msoTrue;

                    if (presenting)
                    {
                        log.Debug($"Key pressed while presenting: vkCode={vkCode}");

                        // VK_CONTROL = 0x11. Left (0xA2) and right (0xA3) also checked.
                        bool isCtrl = vkCode == 0x11 || vkCode == 0xA2 || vkCode == 0xA3;

                        // VK_SHIFT = 0x10. Left (0xA0) and right (0xA1) also checked.
                        // Shift toggles Auto Scripture Mode — say a reference and it inserts.
                        bool isShift = vkCode == 0x10 || vkCode == 0xA0 || vkCode == 0xA1;

                        if (isCtrl)
                        {
                            log.Debug("Opening Add Content Live form");
                            var mainForm = System.Windows.Forms.Application.OpenForms.Count > 0
                                ? System.Windows.Forms.Application.OpenForms[0]
                                : null;

                            if (mainForm != null && mainForm.InvokeRequired)
                            {
                                mainForm.BeginInvoke((Action)ShowAddContentLiveForm);
                            }
                            else
                            {
                                ShowAddContentLiveForm();
                            }
                        }
                        else if (isShift)
                        {
                            // Toggle Auto Scripture Mode
                            var mainForm = System.Windows.Forms.Application.OpenForms.Count > 0
                                ? System.Windows.Forms.Application.OpenForms[0]
                                : null;

                            Action toggleAction = () =>
                            {
                                try
                                {
                                    bool nowOn = AutoScriptureMode.Instance.Toggle(
                                        Globals.ThisAddIn.SpeechService);
                                    log.Info($"Auto Scripture Mode toggled via Shift key: {(nowOn ? "ON" : "OFF")}");
                                }
                                catch (Exception ex)
                                {
                                    log.Error("Failed to toggle Auto Scripture Mode.", ex);
                                }
                            };

                            if (mainForm != null && mainForm.InvokeRequired)
                                mainForm.BeginInvoke(toggleAction);
                            else
                                toggleAction();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                log.Error("Unexpected error while handling keypress", e);
            }

            return SafeNativeMethods.CallNextHookEx(_hookIdKeyboard, nCode, wParam, lParam);
        }

        private void ShowAddContentLiveForm()
        {
            try
            {
                using (var addContentLiveForm = new AddContentLiveForm())
                {
                    addContentLiveForm.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error showing AddContentLiveForm", ex);
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private class KBDLLHOOKSTRUCT
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        public static int FLAG_RELEASED = 0x80;

        internal static class SafeNativeMethods
        {
            public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

            public enum HookType
            {
                WH_KEYBOARD    = 2,
                WH_KEYBOARD_LL = 13,  // Low-level: system-wide, fires on all threads/windows
            }

            public enum WindowMessages : uint
            {
                WM_KEYDOWN    = 0x0100,
                WM_KEYFIRST   = 0x0100,
                WM_KEYLAST    = 0x0108,
                WM_KEYUP      = 0x0101,
                WM_SYSDEADCHAR = 0x0107,
                WM_SYSKEYDOWN = 0x0104,
                WM_SYSKEYUP   = 0x0105
            }

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern IntPtr GetModuleHandle(string lpModuleName);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool UnhookWindowsHookEx(IntPtr hhk);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern IntPtr SetWindowsHookEx(
                int idHook,
                HookProc lpfn,
                IntPtr hMod,
                uint dwThreadId);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern IntPtr CallNextHookEx(
                IntPtr hhk,
                int nCode,
                IntPtr wParam,
                IntPtr lParam);

            [DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern int GetCurrentThreadId();
        }

    // Called from ribbon constructor
    public static void PreInitialize()
        {
            // Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            // CodeBase is the location of the DLL
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            appDataPath = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString()) + "\\Data";
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
