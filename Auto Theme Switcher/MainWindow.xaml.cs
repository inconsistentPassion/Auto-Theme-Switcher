using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Composition.SystemBackdrops;
using Microsoft.UI.Windowing;
using WinRT;
using System.Runtime.InteropServices;
using Windows.Devices.Geolocation;
using System.Net.Http;
using System.Text.Json;
using Windows.Storage;
using Windows.UI.ViewManagement;
using System.Diagnostics;
using System.Threading.Tasks;
using System;

namespace AutoThemeSwitcher
{
    public sealed partial class MainWindow : Window
    {
        private WindowsSystemDispatcherQueueHelper? wsdqHelper;
        private MicaController? micaController;
        private SystemBackdropConfiguration? backdropConfiguration;
        private Microsoft.UI.Dispatching.DispatcherQueue? dispatcherQueue;
        private DispatcherTimer timer;
        private DateTime sunrise;
        private DateTime sunset;
        private string location = "";
        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        public MainWindow()
        {
            this.InitializeComponent();
            var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
            var appWindow = AppWindow.GetFromWindowId(windowId);
            appWindow.Resize(new Windows.Graphics.SizeInt32(420, 420)); // Set your desired width and height
            var screenWidth = GetSystemMetrics(0);
            var screenHeight = GetSystemMetrics(1);
            var taskbarHeight = 40; // Approximate taskbar height
            var windowWidth = 420;
            var windowHeight = 420;
            appWindow.Move(new Windows.Graphics.PointInt32(
                screenWidth - windowWidth - 10,
                screenHeight - windowHeight - taskbarHeight -30
            ));
            TrySetMicaBackdrop();
            InitializeWindowStyle();
            InitializeThemeAutomation();
        }

        private void InitializeWindowStyle()
        {
            if (AppWindowTitleBar.IsCustomizationSupported())
            {
                var titleBar = AppWindow.TitleBar;
                titleBar.ExtendsContentIntoTitleBar = true;
                titleBar.PreferredHeightOption = TitleBarHeightOption.Standard;
                titleBar.ButtonBackgroundColor = Colors.Transparent;
                titleBar.ButtonInactiveBackgroundColor = Colors.Transparent;
                titleBar.ButtonHoverBackgroundColor = Colors.Transparent;
                titleBar.ButtonPressedBackgroundColor = Colors.Transparent;
                UpdateTitleBarButtonColors();
            }
        }

        private async void InitializeThemeAutomation()
        {
            await GetLocationAsync();
            await UpdateSunriseSunsetTimes();
            UpdateUI();

            timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromMinutes(1);
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private async Task GetLocationAsync()
        {
            var geolocator = new Geolocator();
            var position = await geolocator.GetGeopositionAsync();
            var lat = position.Coordinate.Point.Position.Latitude;
            var lon = position.Coordinate.Point.Position.Longitude;

            // Use Windows APIs to get location name
            var basicGeoposition = new BasicGeoposition { Latitude = lat, Longitude = lon };
            var geopoint = new Geopoint(basicGeoposition);
            var civicAddress = await Windows.Services.Maps.MapLocationFinder.FindLocationsAtAsync(geopoint);

            if (civicAddress?.Locations.Count > 0)
            {
                var address = civicAddress.Locations[0].Address;
                location = $"{address.Town}, {address.Country}";
            }
            else
            {
                location = $"({lat:F2}, {lon:F2})";
            }
        }

        private async Task UpdateSunriseSunsetTimes()
        {
            var geolocator = new Geolocator();
            var position = await geolocator.GetGeopositionAsync();
            var lat = position.Coordinate.Point.Position.Latitude;
            var lon = position.Coordinate.Point.Position.Longitude;

            (sunrise, sunset) = CalculateSunriseSunset(DateTime.Today, lat, lon);
        }

        private (DateTime Sunrise, DateTime Sunset) CalculateSunriseSunset(DateTime date, double latitude, double longitude)
        {
            const double DEG_TO_RAD = Math.PI / 180.0;

            // Day of year
            int dayOfYear = date.DayOfYear;

            // Convert latitude and longitude to radians
            double latRad = latitude * DEG_TO_RAD;

            // Solar declination
            double declination = 23.45 * DEG_TO_RAD * Math.Sin(DEG_TO_RAD * (360.0 / 365.0) * (dayOfYear - 81));

            // Hour angle
            double hourAngle = Math.Acos(-Math.Tan(latRad) * Math.Tan(declination));

            // Convert to hours, adjusting for longitude
            double solarNoon = 12.0 - (longitude / 15.0);
            double sunriseOffset = hourAngle * 180.0 / (15.0 * Math.PI);

            // Calculate sunrise and sunset times
            double sunriseHour = solarNoon - sunriseOffset;
            double sunsetHour = solarNoon + sunriseOffset;

            // Convert to local time
            var utcOffset = TimeZoneInfo.Local.GetUtcOffset(date);
            DateTime sunriseTime = date.Date.AddHours(sunriseHour).Add(utcOffset);
            DateTime sunsetTime = date.Date.AddHours(sunsetHour).Add(utcOffset);

            // Apply twilight adjustment (civil twilight is approximately 30 minutes)
            sunriseTime = sunriseTime.AddMinutes(-30);
            sunsetTime = sunsetTime.AddMinutes(30);

            return (sunriseTime, sunsetTime);
        }

        private void UpdateUI()
        {
            var hour = DateTime.Now.Hour;
            string greeting = hour switch
            {
                >= 5 and < 12 => "👋 Good morning!",
                >= 12 and < 17 => "👋 Good afternoon!",
                >= 17 and < 22 => "👋 Good evening!",
                _ => "🌙 Good night!"
            };
            LocationTextBlock.Text = greeting;

            SunriseTextBlock.Text = $"🌅 Sunrise: {sunrise.ToString("t")}";
            SunsetTextBlock.Text = $"🌇 Sunset: {sunset.ToString("t")}";

            var now = DateTime.Now;
            var nextSwitch = now < sunrise ? sunrise : now < sunset ? sunset : sunrise.AddDays(1);
            NextSwitchTextBlock.Text = $"⏱ Next switch at: {nextSwitch.ToString("t")}";

            var isDark = now < sunrise || now >= sunset;
            ThemeStatusTextBlock.Text = $"🎨 Current theme: {(isDark ? "Dark" : "Light")}";
        }

        private async void Timer_Tick(object sender, object e)
        {
            var now = DateTime.Now;
            if (now.Date != sunrise.Date)
            {
                await UpdateSunriseSunsetTimes();
            }

            UpdateUI();
            UpdateTheme();
        }

        private void UpdateTheme()
        {
            var now = DateTime.Now;
            var shouldBeDark = now < sunrise || now >= sunset;
            var currentTheme = GetCurrentTheme();

            if ((shouldBeDark && currentTheme != ApplicationTheme.Dark) ||
                (!shouldBeDark && currentTheme != ApplicationTheme.Light))
            {
                SetTheme(shouldBeDark ? ApplicationTheme.Dark : ApplicationTheme.Light);
            }
        }

        private ApplicationTheme GetCurrentTheme()
        {
            var settings = new UISettings();
            var background = settings.GetColorValue(UIColorType.Background);
            return background.R == 0 ? ApplicationTheme.Dark : ApplicationTheme.Light;
        }

        private void SetTheme(ApplicationTheme theme)
        {
            var ps = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-Command Set-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name SystemUsesLightTheme -Value {(theme == ApplicationTheme.Light ? 1 : 0)}",
                CreateNoWindow = true,
                UseShellExecute = false
            };
            Process.Start(ps);
        }

        private void UpdateTitleBarButtonColors()
        {
            var titleBar = AppWindow.TitleBar;
            if (Application.Current.RequestedTheme == ApplicationTheme.Dark)
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

        private bool TrySetMicaBackdrop()
        {
            if (MicaController.IsSupported())
            {
                wsdqHelper = new WindowsSystemDispatcherQueueHelper();
                wsdqHelper.EnsureWindowsSystemDispatcherQueueController();

                backdropConfiguration = new SystemBackdropConfiguration();
                dispatcherQueue = Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread();

                micaController = new MicaController();
                micaController.AddSystemBackdropTarget(this.As<Microsoft.UI.Composition.ICompositionSupportsSystemBackdrop>());
                micaController.SetSystemBackdropConfiguration(backdropConfiguration);

                return true;
            }

            return false;
        }

        private void ToggleButton_Click(object sender, RoutedEventArgs e)
        {
            var currentTheme = GetCurrentTheme();
            SetTheme(currentTheme == ApplicationTheme.Dark ? ApplicationTheme.Light : ApplicationTheme.Dark);
            UpdateUI();
        }

        private void QuitButton_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Exit();
        }
    }

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
        private static extern int CreateDispatcherQueueController([In] DispatcherQueueOptions options, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object dispatcherQueueController);

        private object? m_dispatcherQueueController;
        public void EnsureWindowsSystemDispatcherQueueController()
        {
            if (Windows.System.DispatcherQueue.GetForCurrentThread() != null)
                return;

            if (m_dispatcherQueueController == null)
            {
                DispatcherQueueOptions options;
                options.dwSize = Marshal.SizeOf(typeof(DispatcherQueueOptions));
                options.threadType = 2;    // DQTYPE_THREAD_CURRENT
                options.apartmentType = 2;  // DQTAT_COM_STA

                object dispatcherQueueController = new();
                CreateDispatcherQueueController(options, ref dispatcherQueueController);
                m_dispatcherQueueController = dispatcherQueueController;
            }
        }
    }
}
