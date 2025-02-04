using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Globalization;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;
using Microsoft.UI;
using Windows.Graphics;

namespace AutoThemeSwitcher
{
    public sealed partial class MainWindow : Window
    {
        private readonly CancellationTokenSource _cts = new();
        private Timer? _monitorTimer;
        private double _lat;
        private double _lng;
        private DateTime _sunrise;
        private DateTime _sunset;
        private bool _isMonitoring = true;
        private readonly HttpClient _httpClient = new();

        public MainWindow()
        {
            InitializeComponent();
            SetupWindow();
            _ = InitializeAsync(); // Start initialization
        }

        private void SetupWindow()
        {
            var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
            var appWindow = AppWindow.GetFromWindowId(windowId);

            if (appWindow != null)
            {
                appWindow.Resize(new SizeInt32(400, 500));
                Title = "Auto Theme Switcher";
            }
        }

        private async Task InitializeAsync()
        {
            try
            {
                await GetLocationAsync();
                await GetSunriseSunsetAsync();
                UpdateDisplay();

                // Use a state object to carry the cancellation token
                var state = new TimerState { CancellationToken = _cts.Token };
                _monitorTimer = new Timer(
                    MonitorThemeCallback,
                    state,
                    TimeSpan.Zero,
                    TimeSpan.FromMinutes(1)
                );
            }
            catch (Exception ex)
            {
                await ShowErrorDialogAsync(ex.Message);
            }
        }

        private class TimerState
        {
            public CancellationToken CancellationToken { get; set; }
        }

        private void MonitorThemeCallback(object? state)
        {
            if (state is TimerState timerState && !timerState.CancellationToken.IsCancellationRequested)
            {
                _ = DispatcherQueue.TryEnqueue(() => MonitorTheme());
            }
        }

        private void MonitorTheme()
        {
            if (!_isMonitoring) return;

            try
            {
                var now = DateTime.Now;
                var shouldBeDark = now < _sunrise || now > _sunset;
                SetTheme(shouldBeDark);
                UpdateDisplay();
            }
            catch (Exception)
            {
                // Log error if needed
            }
        }

        private async Task ShowErrorDialogAsync(string message)
        {
            var errorDialog = new ContentDialog
            {
                XamlRoot = Content.XamlRoot,
                Title = "Error",
                Content = message,
                CloseButtonText = "OK"
            };
            await errorDialog.ShowAsync();
        }

        private async Task GetLocationAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync("http://ip-api.com/json/", _cts.Token);
                response.EnsureSuccessStatusCode();
                var json = await response.Content.ReadAsStringAsync(_cts.Token);

                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                if (root.GetProperty("status").GetString() != "success")
                    throw new Exception("Location service returned unsuccessful status");

                _lat = root.GetProperty("lat").GetDouble();
                _lng = root.GetProperty("lon").GetDouble();
                var city = root.GetProperty("city").GetString();
                var country = root.GetProperty("country").GetString();

                DispatcherQueue.TryEnqueue(() =>
                {
                    LocationTextBlock.Text = $"{city}, {country}";
                });
            }
            catch (Exception ex)
            {
                DispatcherQueue.TryEnqueue(() =>
                {
                    LocationTextBlock.Text = "Location detection failed";
                });
                throw new Exception("Failed to get location: " + ex.Message);
            }
        }

        private async Task GetSunriseSunsetAsync()
        {
            try
            {
                var url = $"https://api.sunrise-sunset.org/json?lat={_lat}&lng={_lng}&formatted=0";
                var response = await _httpClient.GetAsync(url, _cts.Token);
                response.EnsureSuccessStatusCode();
                var json = await response.Content.ReadAsStringAsync(_cts.Token);

                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                if (root.GetProperty("status").GetString() != "OK")
                    throw new Exception("Sunrise/sunset service returned unsuccessful status");

                var results = root.GetProperty("results");
                _sunrise = DateTime.Parse(
                    results.GetProperty("sunrise").GetString()!,
                    null,
                    DateTimeStyles.AdjustToUniversal
                ).ToLocalTime();

                _sunset = DateTime.Parse(
                    results.GetProperty("sunset").GetString()!,
                    null,
                    DateTimeStyles.AdjustToUniversal
                ).ToLocalTime();

                DispatcherQueue.TryEnqueue(UpdateDisplay);
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to get sunrise/sunset times: " + ex.Message);
            }
        }

        private DateTime CalculateNextSwitch()
        {
            var now = DateTime.Now;
            if (now < _sunrise)
                return _sunrise;
            if (now < _sunset)
                return _sunset;

            return _sunrise.AddDays(1);
        }

        private void UpdateDisplay()
        {
            SunriseTextBlock.Text = $"Sunrise: {_sunrise:HH:mm}";
            SunsetTextBlock.Text = $"Sunset: {_sunset:HH:mm}";
            NextSwitchTextBlock.Text = $"Next switch: {CalculateNextSwitch():HH:mm}";
        }

        private void SetTheme(bool darkMode)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                    true);

                if (key == null) return;

                var value = darkMode ? 0 : 1;
                var currentValue = key.GetValue("AppsUseLightTheme");

                if (currentValue != null && (int)currentValue == value)
                    return;

                key.SetValue("AppsUseLightTheme", value, RegistryValueKind.DWord);
                key.SetValue("SystemUseLightTheme", value, RegistryValueKind.DWord);

                DispatcherQueue.TryEnqueue(() =>
                {
                    ThemeStatusTextBlock.Text = $"Current theme: {(darkMode ? "Dark" : "Light")}";
                });
            }
            catch (Exception)
            {
                // Log error if needed
            }
        }

        private void ToggleButton_Click(object sender, RoutedEventArgs e)
        {
            _isMonitoring = !_isMonitoring;
            ToggleButton.Content = _isMonitoring ? "⏯  Pause" : "⏯  Resume";

            if (_monitorTimer == null) return;

            if (_isMonitoring)
                _monitorTimer.Change(TimeSpan.Zero, TimeSpan.FromMinutes(1));
            else
                _monitorTimer.Change(Timeout.Infinite, Timeout.Infinite);
        }

        private void QuitButton_Click(object sender, RoutedEventArgs e)
        {
            _cts.Cancel();
            _monitorTimer?.Dispose();
            Application.Current.Exit();
        }
    }
}