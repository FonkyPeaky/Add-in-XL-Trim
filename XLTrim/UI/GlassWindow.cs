using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace XLTrim.UI
{
    /// <summary>
    /// Base window for XL Trim dialogs.
    /// On Windows 11 the OS rounds the window corners via DWM (DWMWA_WINDOW_CORNER_PREFERENCE).
    /// On Windows 10 the window stays rectangular (system default).
    /// Transparency is handled purely in XAML (AllowsTransparency + Background).
    /// </summary>
    public class GlassWindow : Window
    {
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(
            IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

        private const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
        private const int DWMWCP_ROUND                   = 2;   // Standard rounded corners

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            try
            {
                var hwnd = new WindowInteropHelper(this).Handle;
                int pref = DWMWCP_ROUND;
                DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, ref pref, sizeof(int));
            }
            catch { /* Windows 10 — API not available, silently ignored */ }
        }
    }
}
