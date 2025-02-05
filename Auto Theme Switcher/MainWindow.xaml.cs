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

        public MainWindow()
        {
            this.InitializeComponent();
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

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add("User-Agent", "AutoThemeSwitcher");
            try
            {
                var response = await client.GetStringAsync($"https://timeanddate.com/services/geolocator?x={lon}&y={lat}");
                var locationData = JsonSerializer.Deserialize<JsonElement>(response);
                location = $"{locationData.GetProperty("city").GetString()}, {locationData.GetProperty("country").GetString()}";
            }
            catch (HttpRequestException e)
            {
                // Log the exception or handle it accordingly
                Console.WriteLine($"Request error: {e.Message}");
                location = "Unknown location";
            }
        }

        private async Task UpdateSunriseSunsetTimes()
        {
            var geolocator = new Geolocator();
            var position = await geolocator.GetGeopositionAsync();
            var lat = position.Coordinate.Point.Position.Latitude;
            var lon = position.Coordinate.Point.Position.Longitude;

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add("User-Agent", "AutoThemeSwitcher");
            var response = await client.GetAsync($"https://timeanddate.com/sun/api?x={lon}&y={lat}");
            if (response.IsSuccessStatusCode)
            {
                var responseBody = await response.Content.ReadAsStringAsync();
                var sunData = JsonSerializer.Deserialize<JsonElement>(responseBody);

                sunrise = DateTime.Parse(sunData.GetProperty("sunrise").GetString()!);
                sunset = DateTime.Parse(sunData.GetProperty("sunset").GetString()!);
            }
            else
            {
                // Handle the error appropriately
                Console.WriteLine($"Error: {response.StatusCode}");
            }
        }

        private void UpdateUI()
        {
            LocationTextBlock.Text = $"📍 {location}";
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
