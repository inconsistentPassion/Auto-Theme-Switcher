﻿using Microsoft.UI;
using Microsoft.UI.Composition.SystemBackdrops;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.ApplicationModel;
using Windows.Devices.Geolocation;
using Windows.UI.ViewManagement;
using WinRT;


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
        private bool _isClosing;

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

        public MainWindow()
        {
            this.InitializeComponent();
            Title = $"AutoThemes";
            InitializeTrayIcon();
            SetWindowPositionAndSize();
            TrySetMicaBackdrop();
            InitializeWindowStyle();
            _timer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(1) };
            _timer.Tick += Timer_Tick;

            // Initialize theme automation (asynchronously)
            _ = InitializeThemeAutomationAsync();

            _isAutomationEnabled = LoadAutomationSetting();
            AutomationToggleSwitch.IsOn = _isAutomationEnabled;

            this.Closed += MainWindow_Closed;

            // Initialize Startup Task Toggle
            _ = InitializeStartupTaskAsync();
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

            // Get DPI scaling factor
            var dpi = GetDpiForWindow(windowHandle);
            var scalingFactor = (float)dpi / 96.0f;

            // Adjust window dimensions for DPI scaling
            const int baseWindowDimensions = 520;
            var scaledDimensions = (int)(baseWindowDimensions * scalingFactor);

            appWindow.Resize(new Windows.Graphics.SizeInt32(scaledDimensions, scaledDimensions));

            // Get scaled screen metrics
            var screenWidth = GetSystemMetrics(0);
            var screenHeight = GetSystemMetrics(1);
            var taskbarHeight = (int)(40 * scalingFactor); // Scale taskbar height

            // Calculate position accounting for scaling
            var posX = screenWidth - scaledDimensions - (int)(10 * scalingFactor);
            var posY = screenHeight - scaledDimensions - taskbarHeight - (int)(30 * scalingFactor);

            appWindow.Move(new Windows.Graphics.PointInt32(posX, posY));

            if (appWindow.Presenter is OverlappedPresenter presenter)
            {
                presenter.IsResizable = false;
                presenter.IsMaximizable = false;
            }
        }

        // Add this DPI helper method
        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

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

        private async Task ConfigureStartupTaskAsync(bool enable)
        {
            StartupTask startupTask = await StartupTask.GetAsync("MyStartupTask");
            if (enable)
            {
                StartupTaskState newState = await startupTask.RequestEnableAsync();
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
        public class LocationData
        {
            public double Latitude { get; set; }
            public double Longitude { get; set; }
            public string Location { get; set; } = string.Empty;
        }

        // Returns the file path where location data will be stored.
        private string GetLocationDataFilePath()
        {
            // For example, store the file in %APPDATA%\AutoThemeSwitcher\location.json
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AutoThemeSwitcher");
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return Path.Combine(folder, "location.json");
        }

        // Saves the location data to disk.
        private void SaveLocationData(LocationData data)
        {
            try
            {
                string filePath = GetLocationDataFilePath();
                string json = JsonSerializer.Serialize(data);
                File.WriteAllText(filePath, json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error saving location data: " + ex.Message);
            }
        }

        // Loads the location data from disk, if it exists.
        private LocationData? LoadLocationData()
        {
            try
            {
                string filePath = GetLocationDataFilePath();
                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    return JsonSerializer.Deserialize<LocationData>(json);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error loading location data: " + ex.Message);
            }
            return null;
        }

        #region Event Handlers
        private void MainWindow_Closed(object sender, WindowEventArgs e)
        {
            if (_isClosing)
            {
                // Allow the window to close if we're actually exiting
                return;
            }

            // Instead of closing, hide the window
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
            _isClosing = true;
            _timer.Stop();
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            Microsoft.UI.Xaml.Application.Current.Exit();
        }

        private void AutomationToggleSwitch_Toggled(object sender, RoutedEventArgs e)
        {
            _isAutomationEnabled = ((ToggleSwitch)sender).IsOn;
            SaveAutomationSetting(_isAutomationEnabled);
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
        private void SaveAutomationSetting(bool isEnabled)
        {
            Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"SOFTWARE\AutoThemeSwitcher")
                .SetValue("AutomationEnabled", isEnabled ? 1 : 0);
        }

        private bool LoadAutomationSetting()
        {
            var value = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\AutoThemeSwitcher")
                ?.GetValue("AutomationEnabled");
            return value != null && (int)value == 1;
        }
        #endregion

        #region Theme and UI Update Methods

        private async Task GetAndCacheLocationAsync()
        {
            double? newLatitude = null;
            double? newLongitude = null;
            string newLocation = string.Empty;
            bool fetched = false;

            try
            {
                var geolocator = new Geolocator();
                var position = await geolocator.GetGeopositionAsync();

                newLatitude = position.Coordinate.Point.Position.Latitude;
                newLongitude = position.Coordinate.Point.Position.Longitude;

                var basicPosition = new BasicGeoposition
                {
                    Latitude = newLatitude.Value,
                    Longitude = newLongitude.Value
                };

                var geopoint = new Geopoint(basicPosition);
                var civicAddress = await Windows.Services.Maps.MapLocationFinder.FindLocationsAtAsync(geopoint);

                if (civicAddress?.Locations.Count > 0)
                {
                    var address = civicAddress.Locations[0].Address;
                    newLocation = $"{address.Town}, {address.Country}";
                }
                else
                {
                    newLocation = $"({newLatitude:F2}, {newLongitude:F2})";
                }
                fetched = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error fetching location: {ex.Message}");
            }

            // If we couldn't fetch a new location, attempt to load stored location data.
            if (!fetched)
            {
                var storedData = LoadLocationData();
                if (storedData != null)
                {
                    _cachedLatitude = storedData.Latitude;
                    _cachedLongitude = storedData.Longitude;
                    _location = storedData.Location;
                    Debug.WriteLine($"Using stored location data from previous instance: {_location}");
                }
                else
                {
                    _location = "Unknown location";
                }
                return;
            }

            // Define a threshold for detecting significant location change.
            const double threshold = 0.001; // roughly 111 meters at the equator

            // If a cached location exists, check if the new location differs significantly.
            if (_cachedLatitude.HasValue && _cachedLongitude.HasValue)
            {
                if (Math.Abs(newLatitude.Value - _cachedLatitude.Value) < threshold &&
                    Math.Abs(newLongitude.Value - _cachedLongitude.Value) < threshold)
                {
                    Debug.WriteLine("Location unchanged; using cached location.");
                    return;
                }
            }

            // Update cached values.
            _cachedLatitude = newLatitude;
            _cachedLongitude = newLongitude;
            _location = newLocation;
            Debug.WriteLine($"New location detected: {_location}");

            // Persist the new location.
            SaveLocationData(new LocationData
            {
                Latitude = newLatitude.Value,
                Longitude = newLongitude.Value,
                Location = newLocation
            });
        }
        private async void StartupToggleSwitch_Toggled(object sender, RoutedEventArgs e)
        {
            bool enable = StartupToggleSwitch.IsOn;
            try
            {
                StartupTask startupTask = await StartupTask.GetAsync("MyStartupTask");

                switch (startupTask.State)
                {
                    case StartupTaskState.Disabled:
                        if (enable)
                        {
                            StartupTaskState newState = await startupTask.RequestEnableAsync();
                            StartupToggleSwitch.IsOn = (newState == StartupTaskState.Enabled);
                            Debug.WriteLine($"StartupTask state after request: {newState}");
                        }
                        break;

                    case StartupTaskState.Enabled:
                        if (!enable)
                        {
                            startupTask.Disable();
                            StartupToggleSwitch.IsOn = false;
                            Debug.WriteLine("StartupTask has been disabled.");
                        }
                        break;

                    case StartupTaskState.EnabledByPolicy:
                    case StartupTaskState.DisabledByPolicy:
                        // Handle policies if necessary
                        StartupToggleSwitch.IsEnabled = false;
                        Debug.WriteLine("StartupTask state is managed by policy and cannot be changed.");
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception occurred: {ex.Message}");
                // Optionally, inform the user about the error
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
                DateTime localNoon = new(date.Year, date.Month, date.Day, 12, 0, 0, DateTimeKind.Local);
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
        private async Task InitializeStartupTaskAsync()
        {
            try
            {
                StartupTask startupTask = await StartupTask.GetAsync("MyStartupTask");
                StartupToggleSwitch.IsOn = (startupTask.State == StartupTaskState.Enabled);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to initialize startup task: {ex.Message}");
            }
        }

        #endregion
    }

    public static class WindowManagement
    {
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        private const int SW_RESTORE = 9;

        public static void BringWindowToFront(IntPtr hWnd)
        {
            if (IsIconic(hWnd))
            {
                ShowWindow(hWnd, SW_RESTORE);
            }

            SetForegroundWindow(hWnd);
        }
    }

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

        public void BringToFront()
        {
            var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
            var appWindow = AppWindow.GetFromWindowId(windowId);

            appWindow.Show();
            WindowManagement.SetForegroundWindow(windowHandle);

            if (appWindow.Presenter is OverlappedPresenter presenter)
            {
                presenter.Restore();
            }
        }
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
#endregion