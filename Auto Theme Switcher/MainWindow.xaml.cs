using Microsoft.UI;
using Microsoft.UI.Composition.SystemBackdrops;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
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

        // New fields for improved functionality
        private DateTime _lastLocationUpdate;
        private const int LocationUpdateIntervalHours = 6;
        private ProgressBar? _switchProgressBar;

        #endregion

        #region Win32 Interop

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, int Msg, IntPtr wParam, string lParam,
            SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        [Flags]
        private enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0,
        }

        #endregion

        public MainWindow()
        {
            this.InitializeComponent();
            Title = "Auto Theme Switcher";
            
            // Initialize components
            InitializeTrayIcon();
            SetWindowPositionAndSize();
            TrySetMicaBackdrop();
            InitializeWindowStyle();
            InitializeProgressBar();
            
            // Setup timer with improved interval
            _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) }; // More frequent updates
            _timer.Tick += Timer_Tick;

            // Initialize theme automation
            _ = InitializeThemeAutomationAsync();

            _isAutomationEnabled = LoadAutomationSetting();
            AutomationToggleSwitch.IsOn = _isAutomationEnabled;

            this.Closed += MainWindow_Closed;

            // Initialize Startup Task
            _ = InitializeStartupTaskAsync();
            
            // Add keyboard shortcuts
            this.KeyDown += MainWindow_KeyDown;
        }

        #region Initialization Methods

        private void InitializeTrayIcon()
        {
            _trayIcon = new NotifyIcon();
            string iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico");
            
            // Fallback to system icon if custom icon not found
            if (File.Exists(iconPath))
            {
                _trayIcon.Icon = new Icon(iconPath);
            }
            else
            {
                // Create a simple themed icon
                _trayIcon.Icon = CreateThemedIcon();
            }
            
            _trayIcon.Visible = true;
            _trayIcon.Click += TrayIcon_Click;
            _trayIcon.DoubleClick += TrayIcon_DoubleClick;
            InitializeTrayIconMenu();
        }

        private Icon CreateThemedIcon()
        {
            // Create a simple bitmap icon with theme-aware colors
            var bitmap = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Transparent);
                
                // Draw a simple moon/sun icon
                var rect = new Rectangle(4, 4, 24, 24);
                g.FillEllipse(Brushes.Gold, rect);
                
                // Add some stars for dark theme
                g.FillEllipse(Brushes.White, new Rectangle(8, 8, 4, 4));
                g.FillEllipse(Brushes.White, new Rectangle(20, 12, 3, 3));
            }
            
            return Icon.FromHandle(bitmap.GetHicon());
        }

        private void InitializeTrayIconMenu()
        {
            var contextMenu = new ContextMenuStrip();
            
            // Open item
            var openItem = new ToolStripMenuItem("Open Auto Theme Switcher");
            openItem.Click += (s, e) => TrayIcon_Click(null, EventArgs.Empty);
            contextMenu.Items.Add(openItem);
            
            // Toggle theme item
            var toggleItem = new ToolStripMenuItem("Toggle Theme");
            toggleItem.Click += (s, e) => ToggleTheme();
            contextMenu.Items.Add(toggleItem);
            
            // Separator
            contextMenu.Items.Add(new ToolStripSeparator());
            
            // Settings item
            var settingsItem = new ToolStripMenuItem("Settings");
            settingsItem.Click += (s, e) => ShowSettingsDialog();
            contextMenu.Items.Add(settingsItem);
            
            // Separator
            contextMenu.Items.Add(new ToolStripSeparator());
            
            // Exit item
            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (s, e) => QuitApplication();
            contextMenu.Items.Add(exitItem);
            
            _trayIcon.ContextMenuStrip = contextMenu;
        }

        private void InitializeProgressBar()
        {
            // Initialize the progress bar for next switch countdown
            if (SwitchProgressBar != null)
            {
                SwitchProgressBar.Minimum = 0;
                SwitchProgressBar.Maximum = 100;
                SwitchProgressBar.Value = 0;
                SwitchProgressBar.CornerRadius = new CornerRadius(2);
            }
        }

        private void SetWindowPositionAndSize()
        {
            try
            {
                var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
                var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
                var appWindow = AppWindow.GetFromWindowId(windowId);

                // Get DPI scaling factor
                var dpi = GetDpiForWindow(windowHandle);
                var scalingFactor = (float)dpi / 96.0f;

                // Improved window dimensions
                const int baseWidth = 480;
                const int baseHeight = 520;
                var scaledWidth = (int)(baseWidth * scalingFactor);
                var scaledHeight = (int)(baseHeight * scalingFactor);

                appWindow.Resize(new Windows.Graphics.SizeInt32(scaledWidth, scaledHeight));

                // Get scaled screen metrics
                var screenWidth = GetSystemMetrics(0);
                var screenHeight = GetSystemMetrics(1);
                var taskbarHeight = (int)(40 * scalingFactor);

                // Calculate position accounting for scaling
                var posX = screenWidth - scaledWidth - (int)(20 * scalingFactor);
                var posY = screenHeight - scaledHeight - taskbarHeight - (int)(40 * scalingFactor);

                appWindow.Move(new Windows.Graphics.PointInt32(posX, posY));

                if (appWindow.Presenter is OverlappedPresenter presenter)
                {
                    presenter.IsResizable = false;
                    presenter.IsMaximizable = false;
                    presenter.IsMinimizable = true;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error setting window position: {ex.Message}");
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
                
                // Make title bar buttons transparent to match acrylic background
                titleBar.ButtonBackgroundColor = Colors.Transparent;
                titleBar.ButtonInactiveBackgroundColor = Colors.Transparent;
                titleBar.ButtonHoverBackgroundColor = Color.FromArgb(0x20, 0xFF, 0xFF, 0xFF);
                titleBar.ButtonPressedBackgroundColor = Color.FromArgb(0x40, 0xFF, 0xFF, 0xFF);
                
                UpdateTitleBarButtonColors();
            }
        }

        private async Task ConfigureStartupTaskAsync(bool enable)
        {
            try
            {
                StartupTask startupTask = await StartupTask.GetAsync("MyStartupTask");
                if (enable)
                {
                    StartupTaskState newState = await startupTask.RequestEnableAsync();
                }
                else
                {
                    startupTask.Disable();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error configuring startup task: {ex.Message}");
            }
        }

        private async Task InitializeThemeAutomationAsync()
        {
            try
            {
                await GetAndCacheLocationAsync();
                await UpdateSunriseSunsetTimesAsync();
                UpdateUI();
                UpdateTheme();
                _timer.Start();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error initializing theme automation: {ex.Message}");
                // Show error to user
                LocationTextBlock.Text = "Location unavailable";
            }
        }

        private bool TrySetMicaBackdrop()
        {
            if (DesktopAcrylicController.IsSupported())
            {
                try
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
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error setting backdrop: {ex.Message}");
                    return false;
                }
            }
            return false;
        }

        #endregion

        #region Data Management

        public class LocationData
        {
            public double Latitude { get; set; }
            public double Longitude { get; set; }
            public string Location { get; set; } = string.Empty;
            public DateTime LastUpdated { get; set; }
        }

        private string GetLocationDataFilePath()
        {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AutoThemeSwitcher");
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return Path.Combine(folder, "location.json");
        }

        private void SaveLocationData(LocationData data)
        {
            try
            {
                string filePath = GetLocationDataFilePath();
                string json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(filePath, json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error saving location data: " + ex.Message);
            }
        }

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

        #endregion

        #region Event Handlers

        private void MainWindow_Closed(object sender, WindowEventArgs e)
        {
            if (_isClosing)
            {
                return;
            }

            // Hide to system tray instead of closing
            e.Handled = true;
            var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
            var appWindow = AppWindow.GetFromWindowId(windowId);
            appWindow.Hide();
            
            // Show notification
            _trayIcon.ShowBalloonTip(2000, "Auto Theme Switcher", "Application minimized to system tray", ToolTipIcon.Info);
        }

        private void MainWindow_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
        {
            // Keyboard shortcuts
            if (e.Key == Windows.System.VirtualKey.T && 
                Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down))
            {
                e.Handled = true;
                ToggleTheme();
            }
            else if (e.Key == Windows.System.VirtualKey.R && 
                     Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down))
            {
                e.Handled = true;
                _ = RefreshLocationAndTimesAsync();
            }
        }

        private void TrayIcon_Click(object? sender, EventArgs e)
        {
            var mouseArgs = e as MouseEventArgs;
            if (mouseArgs?.Button == MouseButtons.Left)
            {
                ShowWindow();
            }
        }

        private void TrayIcon_DoubleClick(object? sender, EventArgs e)
        {
            ToggleTheme();
        }

        private async void Timer_Tick(object? sender, object? e)
        {
            if (_isAutomationEnabled)
            {
                // Update location periodically
                if (DateTime.Now.Subtract(_lastLocationUpdate).TotalHours >= LocationUpdateIntervalHours)
                {
                    await GetAndCacheLocationAsync();
                }

                // If a new day has started, update sunrise/sunset times
                if (DateTime.Today != _sunrise.Date)
                {
                    await UpdateSunriseSunsetTimesAsync();
                }
                
                UpdateUI();
                UpdateTheme();
                UpdateProgressBar();
            }
        }

        private async void RefreshLocationButton_Click(object sender, RoutedEventArgs e)
        {
            await RefreshLocationAndTimesAsync();
        }

        private async Task RefreshLocationAndTimesAsync()
        {
            LocationTextBlock.Text = "Refreshing location...";
            
            try
            {
                await GetAndCacheLocationAsync();
                await UpdateSunriseSunsetTimesAsync();
                UpdateUI();
                ShowNotification("Location and times updated successfully", "success");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error refreshing location: {ex.Message}");
                LocationTextBlock.Text = "Location unavailable";
                ShowNotification("Failed to refresh location", "error");
            }
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
                ShowNotification("Theme automation enabled", "info");
            }
            else
            {
                _timer.Stop();
                ShowNotification("Theme automation disabled", "warning");
            }
        }

        private void SaveAutomationSetting(bool isEnabled)
        {
            try
            {
                Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"SOFTWARE\AutoThemeSwitcher")
                    .SetValue("AutomationEnabled", isEnabled ? 1 : 0);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving automation setting: {ex.Message}");
            }
        }

        private bool LoadAutomationSetting()
        {
            try
            {
                var value = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\AutoThemeSwitcher")
                    ?.GetValue("AutomationEnabled");
                return value != null && (int)value == 1;
            }
            catch
            {
                return true; // Default to enabled
            }
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
            var appWindow = AppWindow.GetFromWindowId(windowId);
            appWindow.Hide();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            QuitApplication();
        }

        private void QuitApplication()
        {
            _isClosing = true;
            _timer.Stop();
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            Microsoft.UI.Xaml.Application.Current.Exit();
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
                geolocator.DesiredAccuracyInMeters = 100; // Reduced accuracy for faster response
                var position = await geolocator.GetGeopositionAsync();

                newLatitude = position.Coordinate.Point.Position.Latitude;
                newLongitude = position.Coordinate.Point.Position.Longitude;

                var basicPosition = new BasicGeoposition
                {
                    Latitude = newLatitude.Value,
                    Longitude = newLongitude.Value
                };

                var geopoint = new Geopoint(basicPosition);
                
                // Try to get civic address with timeout
                var civicAddressTask = Windows.Services.Maps.MapLocationFinder.FindLocationsAtAsync(geopoint);
                if (await Task.WhenAny(civicAddressTask, Task.Delay(5000)) == civicAddressTask)
                {
                    var civicAddress = civicAddressTask.Result;
                    if (civicAddress?.Locations.Count > 0)
                    {
                        var address = civicAddress.Locations[0].Address;
                        newLocation = $"{address.Town ?? address.Region ?? "Unknown"}, {address.Country ?? "Unknown"}";
                    }
                    else
                    {
                        newLocation = $"({newLatitude:F2}, {newLongitude:F2})";
                    }
                }
                else
                {
                    newLocation = $"({newLatitude:F2}, {newLongitude:F2})";
                }
                
                fetched = true;
                _lastLocationUpdate = DateTime.Now;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error fetching location: {ex.Message}");
            }

            // If we couldn't fetch a new location, attempt to load stored location data
            if (!fetched)
            {
                var storedData = LoadLocationData();
                if (storedData != null)
                {
                    _cachedLatitude = storedData.Latitude;
                    _cachedLongitude = storedData.Longitude;
                    _location = storedData.Location;
                    Debug.WriteLine($"Using stored location data: {_location}");
                }
                else
                {
                    _location = "Unknown location";
                }
                return;
            }

            // Define a threshold for detecting significant location change
            const double threshold = 0.001; // roughly 111 meters at the equator

            // If a cached location exists, check if the new location differs significantly
            if (_cachedLatitude.HasValue && _cachedLongitude.HasValue)
            {
                if (Math.Abs(newLatitude.Value - _cachedLatitude.Value) < threshold &&
                    Math.Abs(newLongitude.Value - _cachedLongitude.Value) < threshold)
                {
                    Debug.WriteLine("Location unchanged; using cached location.");
                    return;
                }
            }

            // Update cached values
            _cachedLatitude = newLatitude;
            _cachedLongitude = newLongitude;
            _location = newLocation;
            Debug.WriteLine($"New location detected: {_location}");

            // Persist the new location
            SaveLocationData(new LocationData
            {
                Latitude = newLatitude.Value,
                Longitude = newLongitude.Value,
                Location = newLocation,
                LastUpdated = DateTime.Now
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
                            ShowNotification("Application will run on startup", "success");
                        }
                        break;

                    case StartupTaskState.Enabled:
                        if (!enable)
                        {
                            startupTask.Disable();
                            StartupToggleSwitch.IsOn = false;
                            ShowNotification("Removed from startup", "info");
                        }
                        break;

                    case StartupTaskState.EnabledByPolicy:
                    case StartupTaskState.DisabledByPolicy:
                        StartupToggleSwitch.IsEnabled = false;
                        ShowNotification("Startup is managed by system policy", "warning");
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception occurred: {ex.Message}");
                ShowNotification("Failed to configure startup settings", "error");
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
                // Use default times as fallback
                _sunrise = DateTime.Today.AddHours(6).AddMinutes(30);
                _sunset = DateTime.Today.AddHours(18).AddMinutes(30);
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
            try
            {
                int hour = DateTime.Now.Hour;
                string greeting = hour switch
                {
                    >= 5 and < 12 => "🌅 Good morning!",
                    >= 12 and < 17 => "☀️ Good afternoon!",
                    >= 17 and < 22 => "🌇 Good evening!",
                    _ => "🌙 Good night!"
                };

                // Update location with greeting
                if (!string.IsNullOrEmpty(_location) && _location != "Unknown location")
                {
                    LocationTextBlock.Text = $"{greeting} - {_location}";
                }
                else
                {
                    LocationTextBlock.Text = greeting;
                }

                // Update times
                SunriseTextBlock.Text = $"{_sunrise:t}";
                SunsetTextBlock.Text = $"{_sunset:t}";

                // Calculate next switch time
                DateTime now = DateTime.Now;
                DateTime nextSwitch = now < _sunrise ? _sunrise : now < _sunset ? _sunset : _sunrise.AddDays(1);
                NextSwitchTextBlock.Text = $"{nextSwitch:t}";

                // Update theme status
                bool isDark = now < _sunrise || now >= _sunset;
                ThemeStatusTextBlock.Text = isDark ? "Dark Mode" : "Light Mode";
                
                // Update theme indicator
                if (ThemeIndicator != null)
                {
                    ThemeIndicator.Fill = isDark ? 
                        new SolidColorBrush(Color.FromArgb(0xFF, 0x1f, 0x29, 0x37)) : 
                        new SolidColorBrush(Color.FromArgb(0xFF, 0xfb, 0xbf, 0x24));
                }

                // Update theme icon
                if (ThemeIcon != null)
                {
                    ThemeIcon.Text = isDark ? "🌙" : "☀️";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error updating UI: {ex.Message}");
            }
        }

        private void UpdateProgressBar()
        {
            if (SwitchProgressBar == null) return;

            try
            {
                DateTime now = DateTime.Now;
                DateTime nextSwitch = now < _sunrise ? _sunrise : now < _sunset ? _sunset : _sunrise.AddDays(1);
                
                DateTime lastSwitch = now < _sunrise ? _sunset.AddDays(-1) : now < _sunset ? _sunrise : _sunset;
                
                TimeSpan totalDuration = nextSwitch - lastSwitch;
                TimeSpan elapsed = now - lastSwitch;
                
                double progress = (elapsed.TotalMinutes / totalDuration.TotalMinutes) * 100;
                progress = Math.Max(0, Math.Min(100, progress));
                
                SwitchProgressBar.Value = progress;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error updating progress bar: {ex.Message}");
            }
        }

        private void UpdateTheme()
        {
            try
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
            catch (Exception ex)
            {
                Debug.WriteLine($"Error updating theme: {ex.Message}");
            }
        }

        private ApplicationTheme GetCurrentTheme()
        {
            try
            {
                var settings = new UISettings();
                var bgColor = settings.GetColorValue(UIColorType.Background);
                return bgColor.R == 0 ? ApplicationTheme.Dark : ApplicationTheme.Light;
            }
            catch
            {
                return ApplicationTheme.Light; // Default fallback
            }
        }

        private async void SetTheme(ApplicationTheme theme)
        {
            try
            {
                // Use PowerShell to update registry settings for theme
                var psInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments =
                        $"-Command Set-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name SystemUsesLightTheme -Value {(theme == ApplicationTheme.Light ? 1 : 0)}; " +
                        $"Set-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name AppsUseLightTheme -Value {(theme == ApplicationTheme.Light ? 1 : 0)}",
                    CreateNoWindow = true,
                    UseShellExecute = false
                };
                
                using var process = Process.Start(psInfo);
                process?.WaitForExit(5000); // Add timeout

                // Broadcast the setting change
                const int HWND_BROADCAST = 0xffff;
                const int WM_SETTINGCHANGE = 0x001A;
                SendMessageTimeout(new IntPtr(HWND_BROADCAST), WM_SETTINGCHANGE, IntPtr.Zero, "ImmersiveColorSet",
                    SendMessageTimeoutFlags.SMTO_NORMAL, 1000, out _);

                await Task.Delay(500);

                // Update the application's requested theme
                _dispatcherQueue?.TryEnqueue(DispatcherQueuePriority.Normal, () =>
                {
                    if (Microsoft.UI.Xaml.Application.Current != null)
                    {
                        Microsoft.UI.Xaml.Application.Current.RequestedTheme = theme;
                        UpdateTitleBarButtonColors();
                        ReapplyBackdrop();
                        
                        // Show notification
                        ShowNotification($"Switched to {(theme == ApplicationTheme.Dark ? "Dark" : "Light")} theme", "info");
                    }
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error setting theme: {ex.Message}");
                ShowNotification("Failed to change theme", "error");
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
            if (titleBar == null) return;

            try
            {
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
            catch (Exception ex)
            {
                Debug.WriteLine($"Error updating title bar colors: {ex.Message}");
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
                StartupToggleSwitch.IsEnabled = false;
            }
        }

        #endregion

        #region Utility Methods

        private void ShowWindow()
        {
            try
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
            catch (Exception ex)
            {
                Debug.WriteLine($"Error showing window: {ex.Message}");
            }
        }

        private void ToggleTheme()
        {
            if (_isThemeChanging)
            {
                return;
            }

            _isThemeChanging = true;
            try
            {
                var currentTheme = GetCurrentTheme();
                SetTheme(currentTheme == ApplicationTheme.Dark ? ApplicationTheme.Light : ApplicationTheme.Dark);
                UpdateUI();
            }
            finally
            {
                _isThemeChanging = false;
            }
        }

        private void ShowSettingsDialog()
        {
            ShowWindow();
            // In a full implementation, this would open a settings dialog
            ShowNotification("Settings panel coming in a future update", "info");
        }

        private void ShowNotification(string message, string type)
        {
            ToolTipIcon icon = type switch
            {
                "error" => ToolTipIcon.Error,
                "warning" => ToolTipIcon.Warning,
                "info" => ToolTipIcon.Info,
                _ => ToolTipIcon.None
            };
            
            _trayIcon.ShowBalloonTip(2000, "Auto Theme Switcher", message, icon);
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