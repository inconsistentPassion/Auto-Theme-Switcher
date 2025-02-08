using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.UI;
using Microsoft.UI.Composition.SystemBackdrops;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using WinRT;
using Windows.Devices.Geolocation;
using Windows.UI.ViewManagement;

namespace AutoThemeSwitcher
{
    public sealed partial class MainWindow : Window
    {
        #region Fields

        private const IconShowOptions ShowIconAndSystemMenu = IconShowOptions.ShowIconAndSystemMenu;

        private WindowsSystemDispatcherQueueHelper? _wsdqHelper;
        private DesktopAcrylicController? _backdropController;
        private SystemBackdropConfiguration? _backdropConfiguration;
        private DispatcherQueue? _dispatcherQueue;
        private readonly DispatcherTimer _timer;

        private NotifyIcon _trayIcon;

        private DateTime _sunrise;
        private DateTime _sunset;
        private string _location = string.Empty;

        // Cached geolocation values
        private double? _cachedLatitude;
        private double? _cachedLongitude;

        private bool _isThemeChanging;
        private bool _isAutomationEnabled;

        #endregion

        #region Win32 Interop

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, int Msg, IntPtr wParam, string lParam,
            SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [Flags]
        private enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0,
        }

        #endregion

        public MainWindow()
        {
            this.InitializeComponent();

            InitializeTrayIcon();
            SetWindowPositionAndSize();
            TrySetMicaBackdrop();
            InitializeWindowStyle();

            // Initialize automation timer (ticks every minute)
            _timer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(1) };
            _timer.Tick += Timer_Tick;

            // Initialize theme automation (asynchronously)
            _ = InitializeThemeAutomationAsync();

            this.Closed += MainWindow_Closed;
        }

        #region Initialization Methods

        private void InitializeTrayIcon()
        {
            _trayIcon = new NotifyIcon();
            string iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico");
            _trayIcon.Icon = File.Exists(iconPath) ? new Icon(iconPath) : SystemIcons.Application;
            _trayIcon.Visible = true;
            _trayIcon.Click += TrayIcon_Click;
            InitializeTrayIconMenu();
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
            toggleThemeItem.Click += (s, e) => ToggleButton_Click(this, new RoutedEventArgs());

            var exitItem = new MenuFlyoutItem
            {
                Text = "Exit",
                Icon = new FontIcon { Glyph = "\uE8BB" }
            };
            exitItem.Click += (s, e) => QuitButton_Click(this, new RoutedEventArgs());

            flyout.Items.Add(openItem);
            flyout.Items.Add(toggleThemeItem);
            flyout.Items.Add(new MenuFlyoutSeparator());
            flyout.Items.Add(exitItem);

            // Use the tray icon's ContextMenuStrip to show the flyout.
            _trayIcon.ContextMenuStrip = new ContextMenuStrip();
            _trayIcon.ContextMenuStrip.Opening += (s, e) =>
            {
                flyout.ShowAt(RootGrid);
                e.Cancel = true;
            };

            _trayIcon.MouseUp += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    flyout.ShowAt(RootGrid);
                }
            };
        }

        private void SetWindowPositionAndSize()
        {
            var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
            var appWindow = AppWindow.GetFromWindowId(windowId);

            const int windowDimensions = 520;
            appWindow.Resize(new Windows.Graphics.SizeInt32(windowDimensions, windowDimensions));

            // Position the window near the bottom-right of the primary screen.
            int screenWidth = GetSystemMetrics(0);
            int screenHeight = GetSystemMetrics(1);
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

        private void InitializeWindowStyle()
        {
            if (AppWindowTitleBar.IsCustomizationSupported())
            {
                var titleBar = AppWindow.TitleBar;
                titleBar.ExtendsContentIntoTitleBar = true;
                titleBar.PreferredHeightOption = TitleBarHeightOption.Standard;
                titleBar.IconShowOptions = ShowIconAndSystemMenu;
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
            _timer.Start();
        }

        private bool TrySetMicaBackdrop()
        {
            if (DesktopAcrylicController.IsSupported())
            {
                _wsdqHelper = new WindowsSystemDispatcherQueueHelper();
                _wsdqHelper.EnsureWindowsSystemDispatcherQueueController();

                _backdropConfiguration = new SystemBackdropConfiguration();
                _dispatcherQueue = DispatcherQueue.GetForCurrentThread();

                _backdropController = new DesktopAcrylicController();
                _backdropController.AddSystemBackdropTarget(this.As<Microsoft.UI.Composition.ICompositionSupportsSystemBackdrop>());
                _backdropController.SetSystemBackdropConfiguration(_backdropConfiguration);
                return true;
            }
            return false;
        }

        #endregion

        #region Event Handlers

        private void MainWindow_Closed(object sender, WindowEventArgs e)
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

        private async void Timer_Tick(object? sender, object? e)
        {
            if (_isAutomationEnabled)
            {
                // If a new day has started, update sunrise/sunset times.
                if (DateTime.Today != _sunrise.Date)
                {
                    await UpdateSunriseSunsetTimesAsync();
                }
                UpdateUI();
                UpdateTheme();
            }
        }

        private async void ToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isThemeChanging)
            {
                return; // Prevent duplicate toggles
            }

            _isThemeChanging = true;
            try
            {
                var currentTheme = GetCurrentTheme();
                // Toggle between dark and light themes.
                await Task.Run(async () =>
                {
                    SetTheme(currentTheme == ApplicationTheme.Dark ? ApplicationTheme.Light : ApplicationTheme.Dark);
                    await Task.Delay(1000); // Cooldown period
                });
                UpdateUI();
            }
            finally
            {
                _isThemeChanging = false;
            }
        }

        private void QuitButton_Click(object sender, RoutedEventArgs e)
        {
            _timer.Stop();
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            Microsoft.UI.Xaml.Application.Current.Exit();
        }

        private void AutomationToggleSwitch_Toggled(object sender, RoutedEventArgs e)
        {
            _isAutomationEnabled = ((ToggleSwitch)sender).IsOn;
            if (_isAutomationEnabled)
            {
                _timer.Start();
                UpdateUI();
                UpdateTheme();
            }
            else
            {
                _timer.Stop();
            }
        }

        #endregion

        #region Theme and UI Update Methods

        private async Task GetAndCacheLocationAsync()
        {
            try
            {
                var geolocator = new Geolocator();
                var position = await geolocator.GetGeopositionAsync();

                _cachedLatitude = position.Coordinate.Point.Position.Latitude;
                _cachedLongitude = position.Coordinate.Point.Position.Longitude;

                var basicPosition = new BasicGeoposition
                {
                    Latitude = _cachedLatitude.Value,
                    Longitude = _cachedLongitude.Value
                };

                var geopoint = new Geopoint(basicPosition);
                var civicAddress = await Windows.Services.Maps.MapLocationFinder.FindLocationsAtAsync(geopoint);

                if (civicAddress?.Locations.Count > 0)
                {
                    var address = civicAddress.Locations[0].Address;
                    _location = $"{address.Town}, {address.Country}";
                }
                else
                {
                    _location = $"({_cachedLatitude:F2}, {_cachedLongitude:F2})";
                }
            }
            catch (Exception ex)
            {
                _location = "Unknown location";
                Debug.WriteLine($"Error fetching location: {ex.Message}");
            }
        }

        private async Task UpdateSunriseSunsetTimesAsync()
        {
            try
            {
                if (_cachedLatitude.HasValue && _cachedLongitude.HasValue)
                {
                    (_sunrise, _sunset) = CalculateSunriseSunset(DateTime.Today, _cachedLatitude.Value, _cachedLongitude.Value);
                }
                else
                {
                    await GetAndCacheLocationAsync();
                    if (_cachedLatitude.HasValue && _cachedLongitude.HasValue)
                    {
                        (_sunrise, _sunset) = CalculateSunriseSunset(DateTime.Today, _cachedLatitude.Value, _cachedLongitude.Value);
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
            try
            {
                DateTime localNoon = new DateTime(date.Year, date.Month, date.Day, 12, 0, 0, DateTimeKind.Local);
                DateTime utcNoon = localNoon.ToUniversalTime();
                int dayOfYear = utcNoon.DayOfYear;

                double fractionalYear = (2 * Math.PI / 365) * (dayOfYear - 1 + (utcNoon.Hour - 12) / 24.0);

                double eqTime = 229.18 * (
                     0.000075
                   + 0.001868 * Math.Cos(fractionalYear)
                   - 0.032077 * Math.Sin(fractionalYear)
                   - 0.014615 * Math.Cos(2 * fractionalYear)
                   - 0.040849 * Math.Sin(2 * fractionalYear));

                double declination = 0.006918
                   - 0.399912 * Math.Cos(fractionalYear)
                   + 0.070257 * Math.Sin(fractionalYear)
                   - 0.006758 * Math.Cos(2 * fractionalYear)
                   + 0.000907 * Math.Sin(2 * fractionalYear)
                   - 0.002697 * Math.Cos(3 * fractionalYear)
                   + 0.00148 * Math.Sin(3 * fractionalYear);

                double latRad = latitude * Math.PI / 180.0;

                double zenithRad = 90.833 * Math.PI / 180.0;

                double cosHA = (Math.Cos(zenithRad) - Math.Sin(latRad) * Math.Sin(declination))
                             / (Math.Cos(latRad) * Math.Cos(declination));

                if (cosHA < -1 || cosHA > 1)
                    return (DateTime.MinValue, DateTime.MinValue);

                double hourAngleDeg = Math.Acos(cosHA) * 180.0 / Math.PI;

                double sunriseMinutesUTC = 720 - 4 * (longitude + hourAngleDeg) - eqTime;
                double sunsetMinutesUTC = 720 - 4 * (longitude - hourAngleDeg) - eqTime;

                DateTime sunriseUtc = utcNoon.Date.AddMinutes(sunriseMinutesUTC);
                DateTime sunsetUtc = utcNoon.Date.AddMinutes(sunsetMinutesUTC);

                DateTime sunrise = TimeZoneInfo.ConvertTimeFromUtc(sunriseUtc, TimeZoneInfo.Local);
                DateTime sunset = TimeZoneInfo.ConvertTimeFromUtc(sunsetUtc, TimeZoneInfo.Local);

                return (sunrise, sunset);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error calculating sunrise/sunset: {ex.Message}");
                return (DateTime.MinValue, DateTime.MinValue);
            }
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

            // Update text blocks (assumed to be defined in XAML)
            LocationTextBlock.Text = greeting;
            SunriseTextBlock.Text = $"{_sunrise:t}";
            SunsetTextBlock.Text = $"{_sunset:t}";

            DateTime now = DateTime.Now;
            DateTime nextSwitch = now < _sunrise ? _sunrise : now < _sunset ? _sunset : _sunrise.AddDays(1);
            NextSwitchTextBlock.Text = $"{nextSwitch:t}";

            bool isDark = now < _sunrise || now >= _sunset;
            ThemeStatusTextBlock.Text = isDark ? "Dark" : "Light";
        }

        private void UpdateTheme()
        {
            DateTime now = DateTime.Now;
            bool shouldBeDark = now < _sunrise || now >= _sunset;
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

        private async void SetTheme(ApplicationTheme theme)
        {
            // Use PowerShell to update registry settings for theme.
            var psInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments =
                    $"-Command Set-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name SystemUsesLightTheme -Value {(theme == ApplicationTheme.Light ? 1 : 0)}; " +
                    $"Set-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name AppsUseLightTheme -Value {(theme == ApplicationTheme.Light ? 1 : 0)}",
                CreateNoWindow = true,
                UseShellExecute = false
            };
            Process.Start(psInfo)?.WaitForExit();

            // Broadcast the setting change.
            const int HWND_BROADCAST = 0xffff;
            const int WM_SETTINGCHANGE = 0x001A;
            SendMessageTimeout(new IntPtr(HWND_BROADCAST), WM_SETTINGCHANGE, IntPtr.Zero, "ImmersiveColorSet",
                SendMessageTimeoutFlags.SMTO_NORMAL, 1000, out _);

            await Task.Delay(500);

            // Update the application's requested theme.
            try
            {
                _dispatcherQueue?.TryEnqueue(DispatcherQueuePriority.Normal, () =>
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
                // Try again after a short delay if a COM exception occurs.
                await Task.Delay(100);
                _dispatcherQueue?.TryEnqueue(DispatcherQueuePriority.Normal, () =>
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

        private void ReapplyBackdrop()
        {
            if (_backdropController != null && _backdropConfiguration != null)
            {
                _backdropController.SetSystemBackdropConfiguration(_backdropConfiguration);
            }
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

        #endregion
    }

    /// <summary>
    /// Helper class to ensure a Windows System Dispatcher Queue Controller exists for the current thread.
    /// </summary>
    class WindowsSystemDispatcherQueueHelper
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct DispatcherQueueOptions
        {
            internal int dwSize;
            internal int threadType;
            internal int apartmentType;
        }

        [DllImport("CoreMessaging.dll")]
        private static extern int CreateDispatcherQueueController(
            [In] DispatcherQueueOptions options,
            [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object dispatcherQueueController);

        private object? _dispatcherQueueController;

        public void EnsureWindowsSystemDispatcherQueueController()
        {
            if (Windows.System.DispatcherQueue.GetForCurrentThread() != null)
            {
                return;
            }

            if (_dispatcherQueueController == null)
            {
                var options = new DispatcherQueueOptions
                {
                    dwSize = Marshal.SizeOf(typeof(DispatcherQueueOptions)),
                    threadType = 2,    // Current thread
                    apartmentType = 2  // COM STA
                };

                object dispatcherQueueController = new();
                CreateDispatcherQueueController(options, ref dispatcherQueueController);
                _dispatcherQueueController = dispatcherQueueController;
            }
        }
    }
}
