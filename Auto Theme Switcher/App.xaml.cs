using AutoThemeSwitcher;
using Microsoft.UI.Xaml;
using System;
using System.Threading;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace Auto_Theme_Switcher
{
    public partial class App : Application
    {
        private const string MutexName = "AutoThemeSwitcherSingleInstance";
        private static Mutex? _mutex;
        private MainWindow? _window;

        public App()
        {
            this.InitializeComponent();
        }
        protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            bool createdNew;
            _mutex = new Mutex(true, MutexName, out createdNew);

            if (!createdNew)
            {
                // Another instance exists, activate it
                ActivateExistingInstance();
                Exit();
                return;
            }

            _window = new MainWindow();
            _window.Activate();
        }
        private void ActivateExistingInstance()
        {
            try
            {
                // Use Windows API to find and activate the existing window
                var processes = System.Diagnostics.Process.GetProcessesByName(
                    System.Diagnostics.Process.GetCurrentProcess().ProcessName);

                foreach (var process in processes)
                {
                    if (process.Id != System.Diagnostics.Process.GetCurrentProcess().Id)
                    {
                        NativeMethods.SetForegroundWindow(process.MainWindowHandle);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error activating existing instance: {ex.Message}");
            }
        }
    }
    internal static class NativeMethods
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        internal const int SW_RESTORE = 9;
    }
}
