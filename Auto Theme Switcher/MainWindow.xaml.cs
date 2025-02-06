using Microsoft.UI;
using Microsoft.UI.Composition.SystemBackdrops;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Devices.Geolocation;
using Windows.UI.ViewManagement;
using WinRT;

namespace AutoThemeSwitcher
{
    public sealed partial class MainWindow : Window
    {
        private const IconShowOptions showIconAndSystemMenu = IconShowOptions.ShowIconAndSystemMenu;
        private WindowsSystemDispatcherQueueHelper? wsdqHelper;
        // Store the backdrop controller to reapply the configuration as needed.
        private DesktopAcrylicController? backdropController;
        private SystemBackdropConfiguration? backdropConfiguration;
        private NotifyIcon trayIcon;
        private Microsoft.UI.Dispatching.DispatcherQueue? dispatcherQueue;
        private DispatcherTimer timer;
        private DateTime sunrise;
        private DateTime sunset;
        private string location = "";

        // Cache geolocation values to avoid duplicate lookups
        private double? cachedLatitude;
        private double? cachedLongitude;

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        private bool isThemeChanging = false;

        public MainWindow()
        {
            this.InitializeComponent();

            InitializeTrayIcon();

            SetWindowPositionAndSize();

            // Try to set the system backdrop and save the controller instance.
            TrySetMicaBackdrop();
            InitializeWindowStyle();

            // Initialize the timer for periodic theme checks.
            timer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(1) };
            timer.Tick += Timer_Tick;

            _ = InitializeThemeAutomationAsync();

            this.Closed += MainWindow_Closed;
        }

        /// <summary>
        /// Initialize the tray icon and its context menu.
        /// </summary>
        private void InitializeTrayIcon()
        {
            trayIcon = new NotifyIcon();
            string iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico");
            trayIcon.Icon = File.Exists(iconPath) ? new Icon(iconPath) : SystemIcons.Application;
            trayIcon.Visible = true;
            trayIcon.Click += TrayIcon_Click;
            InitializeTrayIconMenu();
        }

        /// <summary>
        /// Sets the fixed window position and size. Disables resizing and maximizing.
        /// </summary>
        private void SetWindowPositionAndSize()
        {
            var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
            var appWindow = AppWindow.GetFromWindowId(windowId);
            const int windowDimensions = 520;
            appWindow.Resize(new Windows.Graphics.SizeInt32(windowDimensions, windowDimensions));

            // Position the window near the bottom-right of the primary screen.
            var screenWidth = GetSystemMetrics(0);
            var screenHeight = GetSystemMetrics(1);
            int taskbarHeight = 40; // Approximation
            appWindow.Move(new Windows.Graphics.PointInt32(
                screenWidth - windowDimensions - 10,
                screenHeight - windowDimensions - taskbarHeight - 30
            ));

            // Disable resizing and remove the maximize button.
            if (appWindow.Presenter is OverlappedPresenter presenter)
            {
                presenter.IsResizable = false;
                presenter.IsMaximizable = false;
            }
        }

        /// <summary>
        /// Customize the window title bar.
        /// </summary>
        private void InitializeWindowStyle()
        {
            if (AppWindowTitleBar.IsCustomizationSupported())
            {
                var titleBar = AppWindow.TitleBar;
                titleBar.ExtendsContentIntoTitleBar = true;
                titleBar.PreferredHeightOption = TitleBarHeightOption.Standard;
                titleBar.IconShowOptions = showIconAndSystemMenu;
                titleBar.ButtonBackgroundColor = Colors.Transparent;
                titleBar.ButtonInactiveBackgroundColor = Colors.Transparent;
                titleBar.ButtonHoverBackgroundColor = Colors.Transparent;
                titleBar.ButtonPressedBackgroundColor = Colors.Transparent;
                UpdateTitleBarButtonColors();
            }
        }

        private async Task InitializeThemeAutomationAsync()
        {
            await GetAndCacheLocationAsync();
            await UpdateSunriseSunsetTimesAsync();
            UpdateUI();
            timer.Start();
        }

        private void InitializeTrayIconMenu()
        {
            var flyout = new MenuFlyout();

            var openItem = new MenuFlyoutItem
            {
                Text = "Open",
                Icon = new FontIcon { Glyph = "\uE8A5" }
            };
            openItem.Click += (s, e) => TrayIcon_Click(null, EventArgs.Empty);

            var toggleThemeItem = new MenuFlyoutItem
            {
                Text = "Toggle Theme",
                Icon = new FontIcon { Glyph = "\uE771" }
            };
            toggleThemeItem.Click += (s, e) => ToggleButton_Click(null, new RoutedEventArgs());

            var exitItem = new MenuFlyoutItem
            {
                Text = "Exit",
                Icon = new FontIcon { Glyph = "\uE8BB" }
            };
            exitItem.Click += (s, e) => QuitButton_Click(null, new RoutedEventArgs());

            flyout.Items.Add(openItem);
            flyout.Items.Add(toggleThemeItem);
            flyout.Items.Add(new MenuFlyoutSeparator());
            flyout.Items.Add(exitItem);

            trayIcon.ContextMenuStrip = new ContextMenuStrip();
            trayIcon.ContextMenuStrip.Opening += (s, e) =>
            {
                flyout.ShowAt((FrameworkElement)RootGrid);
                e.Cancel = true;
            };

            trayIcon.MouseUp += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    flyout.ShowAt((FrameworkElement)RootGrid);
                }
            };
        }

        private void MainWindow_Closed(object sender, Microsoft.UI.Xaml.WindowEventArgs e)
        {
            // Instead of closing, hide the window.
            e.Handled = true;
            var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
            var appWindow = AppWindow.GetFromWindowId(windowId);
            appWindow.Hide();
        }

        private void TrayIcon_Click(object? sender, EventArgs e)
        {
            var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
            var appWindow = AppWindow.GetFromWindowId(windowId);
            appWindow.Show();
            appWindow.SetPresenter(AppWindowPresenterKind.Default);
        }

        private async Task GetAndCacheLocationAsync()
        {
            try
            {
                var geolocator = new Geolocator();
                var position = await geolocator.GetGeopositionAsync();
                cachedLatitude = position.Coordinate.Point.Position.Latitude;
                cachedLongitude = position.Coordinate.Point.Position.Longitude;

                var basicPosition = new BasicGeoposition
                {
                    Latitude = cachedLatitude.Value,
                    Longitude = cachedLongitude.Value
                };
                var geopoint = new Geopoint(basicPosition);
                var civicAddress = await Windows.Services.Maps.MapLocationFinder.FindLocationsAtAsync(geopoint);

                if (civicAddress?.Locations.Count > 0)
                {
                    var address = civicAddress.Locations[0].Address;
                    location = $"{address.Town}, {address.Country}";
                }
                else
                {
                    location = $"({cachedLatitude:F2}, {cachedLongitude:F2})";
                }
            }
            catch (Exception ex)
            {
                location = "Unknown location";
                Debug.WriteLine($"Error fetching location: {ex.Message}");
            }
        }

        private async Task UpdateSunriseSunsetTimesAsync()
        {
            try
            {
                if (cachedLatitude.HasValue && cachedLongitude.HasValue)
                {
                    (sunrise, sunset) = CalculateSunriseSunset(DateTime.Today, cachedLatitude.Value, cachedLongitude.Value);
                }
                else
                {
                    await GetAndCacheLocationAsync();
                    if (cachedLatitude.HasValue && cachedLongitude.HasValue)
                    {
                        (sunrise, sunset) = CalculateSunriseSunset(DateTime.Today, cachedLatitude.Value, cachedLongitude.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error updating sunrise/sunset: {ex.Message}");
            }
        }

        private (DateTime Sunrise, DateTime Sunset) CalculateSunriseSunset(DateTime date, double latitude, double longitude)
        {
            const double DEG_TO_RAD = Math.PI / 180.0;
            int dayOfYear = date.DayOfYear;
            double latRad = latitude * DEG_TO_RAD;
            double declination = 23.45 * DEG_TO_RAD * Math.Sin(DEG_TO_RAD * (360.0 / 365.0) * (dayOfYear - 81));
            double hourAngle = Math.Acos(-Math.Tan(latRad) * Math.Tan(declination));
            double solarNoon = 12.0 - (longitude / 15.0);
            double sunriseOffset = hourAngle * 180.0 / (15.0 * Math.PI);
            double sunriseHour = solarNoon - sunriseOffset;
            double sunsetHour = solarNoon + sunriseOffset;
            var utcOffset = TimeZoneInfo.Local.GetUtcOffset(date);

            // Adjust 30 minutes for civil twilight.
            DateTime sunriseTime = date.Date.AddHours(sunriseHour).Add(utcOffset).AddMinutes(-30);
            DateTime sunsetTime = date.Date.AddHours(sunsetHour).Add(utcOffset).AddMinutes(30);
            return (sunriseTime, sunsetTime);
        }

        private void UpdateUI()
        {
            int hour = DateTime.Now.Hour;
            string greeting = hour switch
            {
                >= 5 and < 12 => "👋 Good morning!",
                >= 12 and < 17 => "👋 Good afternoon!",
                >= 17 and < 22 => "👋 Good evening!",
                _ => "🌙 Good night!"
            };

            LocationTextBlock.Text = greeting;
            SunriseTextBlock.Text = $"{sunrise:t}";
            SunsetTextBlock.Text = $"{sunset:t}";

            DateTime now = DateTime.Now;
            DateTime nextSwitch = now < sunrise ? sunrise : now < sunset ? sunset : sunrise.AddDays(1);
            NextSwitchTextBlock.Text = $"{nextSwitch:t}";

            bool isDark = now < sunrise || now >= sunset;
            ThemeStatusTextBlock.Text = $"{(isDark ? "Dark" : "Light")}";
        }

        private async void Timer_Tick(object? sender, object? e)
        {
            // Update sunrise/sunset times if a new day has started.
            if (DateTime.Today != sunrise.Date)
            {
                await UpdateSunriseSunsetTimesAsync();
            }
            UpdateUI();
            UpdateTheme();
        }

        private void UpdateTheme()
        {
            DateTime now = DateTime.Now;
            bool shouldBeDark = now < sunrise || now >= sunset;
            ApplicationTheme currentTheme = GetCurrentTheme();

            if ((shouldBeDark && currentTheme != ApplicationTheme.Dark) ||
                (!shouldBeDark && currentTheme != ApplicationTheme.Light))
            {
                SetTheme(shouldBeDark ? ApplicationTheme.Dark : ApplicationTheme.Light);
            }
        }

        private ApplicationTheme GetCurrentTheme()
        {
            var settings = new UISettings();
            var bgColor = settings.GetColorValue(UIColorType.Background);
            // Assume dark theme if background color is black.
            return bgColor.R == 0 ? ApplicationTheme.Dark : ApplicationTheme.Light;
        }

        /// <summary>
        /// Sets the theme by updating registry keys via PowerShell and reapplying the backdrop.
        /// </summary>
        private async void SetTheme(ApplicationTheme theme)
        {
            var ps = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-Command " +
                            $"Set-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name SystemUsesLightTheme -Value {(theme == ApplicationTheme.Light ? 1 : 0)}; " +
                            $"Set-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name AppsUseLightTheme -Value {(theme == ApplicationTheme.Light ? 1 : 0)}",
                CreateNoWindow = true,
                UseShellExecute = false
            };
            Process.Start(ps)?.WaitForExit();

            // Broadcast the setting change.
            const int HWND_BROADCAST = 0xffff;
            const int WM_SETTINGCHANGE = 0x001A;
            SendMessageTimeout(new IntPtr(HWND_BROADCAST), WM_SETTINGCHANGE, IntPtr.Zero, "ImmersiveColorSet",
                SendMessageTimeoutFlags.SMTO_NORMAL, 1000, out IntPtr _);

            await Task.Delay(500);

            try
            {
                dispatcherQueue?.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Normal, () =>
                {
                    if (Microsoft.UI.Xaml.Application.Current != null &&
                        Microsoft.UI.Xaml.Application.Current.RequestedTheme != theme)
                    {
                        Microsoft.UI.Xaml.Application.Current.RequestedTheme = theme;
                        UpdateTitleBarButtonColors();
                        ReapplyBackdrop();
                    }
                });
            }
            catch (COMException)
            {
                // In case of COM exceptions, try again after a short delay.
                Task.Delay(100).Wait();
                dispatcherQueue?.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Normal, () =>
                {
                    if (Microsoft.UI.Xaml.Application.Current != null &&
                        Microsoft.UI.Xaml.Application.Current.RequestedTheme != theme)
                    {
                        Microsoft.UI.Xaml.Application.Current.RequestedTheme = theme;
                        UpdateTitleBarButtonColors();
                        ReapplyBackdrop();
                    }
                });
            }
        }

        /// <summary>
        /// Reapply the backdrop configuration so that the Mica (or Acrylic) effect remains active.
        /// </summary>
        private void ReapplyBackdrop()
        {
            if (backdropController != null && backdropConfiguration != null)
            {
                backdropController.SetSystemBackdropConfiguration(backdropConfiguration);
            }
        }

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, int Msg, IntPtr wParam, string lParam,
            SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [Flags]
        private enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0,
        }

        private void UpdateTitleBarButtonColors()
        {
            var titleBar = AppWindow.TitleBar;
            if (Microsoft.UI.Xaml.Application.Current.RequestedTheme == ApplicationTheme.Dark)
            {
                titleBar.ButtonForegroundColor = Colors.White;
                titleBar.ButtonHoverForegroundColor = Colors.White;
            }
            else
            {
                titleBar.ButtonForegroundColor = Colors.Black;
                titleBar.ButtonHoverForegroundColor = Colors.Black;
            }
        }

        /// <summary>
        /// Sets up the system backdrop using DesktopAcrylicController (which provides a Mica/Acrylic effect).
        /// </summary>
        /// <returns>True if successfully applied; otherwise false.</returns>
        private bool TrySetMicaBackdrop()
        {
            if (DesktopAcrylicController.IsSupported())
            {
                wsdqHelper = new WindowsSystemDispatcherQueueHelper();
                wsdqHelper.EnsureWindowsSystemDispatcherQueueController();

                backdropConfiguration = new SystemBackdropConfiguration();
                dispatcherQueue = Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread();

                backdropController = new DesktopAcrylicController();
                backdropController.AddSystemBackdropTarget(this.As<Microsoft.UI.Composition.ICompositionSupportsSystemBackdrop>());
                backdropController.SetSystemBackdropConfiguration(backdropConfiguration);
                return true;
            }
            return false;
        }

        private async void ToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (isThemeChanging)
                return; // Prevent duplicate toggles

            isThemeChanging = true;
            try
            {
                var currentTheme = GetCurrentTheme();
                // Toggle between dark and light themes.
                await Task.Run(async () =>
                {
                    SetTheme(currentTheme == ApplicationTheme.Dark ? ApplicationTheme.Light : ApplicationTheme.Dark);
                    await Task.Delay(1000); // Short cooldown period
                });
                UpdateUI();
            }
            finally
            {
                isThemeChanging = false;
            }
        }

        private void QuitButton_Click(object sender, RoutedEventArgs e)
        {
            timer?.Stop();
            trayIcon.Visible = false;
            trayIcon.Dispose();
            Microsoft.UI.Xaml.Application.Current.Exit();
        }
    }

    // Helper class to ensure a Windows system dispatcher queue exists.
    class WindowsSystemDispatcherQueueHelper
    {
        [StructLayout(LayoutKind.Sequential)]
        struct DispatcherQueueOptions
        {
            internal int dwSize;
            internal int threadType;
            internal int apartmentType;
        }

        [DllImport("CoreMessaging.dll")]
        private static extern int CreateDispatcherQueueController(
            [In] DispatcherQueueOptions options,
            [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object dispatcherQueueController);

        private object? m_dispatcherQueueController;

        public void EnsureWindowsSystemDispatcherQueueController()
        {
            if (Windows.System.DispatcherQueue.GetForCurrentThread() != null)
                return;

            if (m_dispatcherQueueController == null)
            {
                DispatcherQueueOptions options;
                options.dwSize = Marshal.SizeOf(typeof(DispatcherQueueOptions));
                options.threadType = 2;    // Current thread
                options.apartmentType = 2; // COM STA

                object dispatcherQueueController = new();
                CreateDispatcherQueueController(options, ref dispatcherQueueController);
                m_dispatcherQueueController = dispatcherQueueController;
            }
        }
    }
}
