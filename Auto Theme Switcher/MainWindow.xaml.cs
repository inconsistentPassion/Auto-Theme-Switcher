using Microsoft.UI;
using Microsoft.UI.Composition.SystemBackdrops;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Windows.Devices.Geolocation;
using Windows.UI.ViewManagement;
using WinRT;
using System.IO;
using System.Drawing;
using System.Windows.Forms;



namespace AutoThemeSwitcher
{
    public sealed partial class MainWindow : Window
    {
        private const IconShowOptions showIconAndSystemMenu = IconShowOptions.ShowIconAndSystemMenu;
        private WindowsSystemDispatcherQueueHelper? wsdqHelper;
        private MicaController? micaController;
        private SystemBackdropConfiguration? backdropConfiguration;
        private NotifyIcon trayIcon;
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
            trayIcon = new NotifyIcon();
            string iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico");
            if (File.Exists(iconPath))
            {
                trayIcon.Icon = new Icon(iconPath);
            }
            else
            {
                trayIcon.Icon = SystemIcons.Application;
            }
            trayIcon.Visible = true;
            trayIcon.Click += TrayIcon_Click;

            trayIcon.Visible = true;
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
                screenHeight - windowHeight - taskbarHeight - 30
            ));
            TrySetMicaBackdrop();
            InitializeWindowStyle();
            timer = new DispatcherTimer(); // Initialize the timer here
            InitializeThemeAutomation();
            this.Closed += MainWindow_Closed;
        }

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

        private void MainWindow_Closed(object sender, Microsoft.UI.Xaml.WindowEventArgs e)
        {
            e.Handled = true; // Cancel the close event
            var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
            var appWindow = AppWindow.GetFromWindowId(windowId);
            appWindow.Hide(); // Hide the window
        }

        private void TrayIcon_Click(object? sender, EventArgs e)
        {
            var windowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(windowHandle);
            var appWindow = AppWindow.GetFromWindowId(windowId);
            appWindow.Show(); // Show the window
            appWindow.SetPresenter(AppWindowPresenterKind.Default); // Restore the window state
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

        private async void Timer_Tick(object? sender, object? e)
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

        private async void SetTheme(ApplicationTheme theme)
        {
            // First set the system theme via PowerShell
            var ps = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-Command Set-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name SystemUsesLightTheme -Value {(theme == ApplicationTheme.Light ? 1 : 0)}; Set-ItemProperty -Path HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize -Name AppsUseLightTheme -Value {(theme == ApplicationTheme.Light ? 1 : 0)}",
                CreateNoWindow = true,
                UseShellExecute = false
            };
            Process.Start(ps)?.WaitForExit();

            // Broadcast theme change message
            const int HWND_BROADCAST = 0xffff;
            const int WM_SETTINGCHANGE = 0x001A;
            SendMessageTimeout(new IntPtr(HWND_BROADCAST), WM_SETTINGCHANGE, IntPtr.Zero, "ImmersiveColorSet", SendMessageTimeoutFlags.SMTO_NORMAL, 1000, out IntPtr _);

            await Task.Delay(100);

            dispatcherQueue?.TryEnqueue(() =>
            {
                if (global::Microsoft.UI.Xaml.Application.Current.RequestedTheme != theme)
                {
                    global::Microsoft.UI.Xaml.Application.Current.RequestedTheme = theme;
                    UpdateTitleBarButtonColors();
                }
            });
        }

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, int Msg, IntPtr wParam, string lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [Flags]
        private enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0,
        }

        private void UpdateTitleBarButtonColors()
        {
            var titleBar = AppWindow.TitleBar;
            if (global::Microsoft.UI.Xaml.Application.Current.RequestedTheme == ApplicationTheme.Dark)
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
            if (DesktopAcrylicController.IsSupported())
            {
                wsdqHelper = new WindowsSystemDispatcherQueueHelper();
                wsdqHelper.EnsureWindowsSystemDispatcherQueueController();

                backdropConfiguration = new SystemBackdropConfiguration();
                dispatcherQueue = Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread();

                var acrylicController = new DesktopAcrylicController();
                acrylicController.AddSystemBackdropTarget(this.As<Microsoft.UI.Composition.ICompositionSupportsSystemBackdrop>());
                acrylicController.SetSystemBackdropConfiguration(backdropConfiguration);

                return true;
            }

            return false;
        }


        private void ToggleButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var currentTheme = GetCurrentTheme();
                SetTheme(currentTheme == ApplicationTheme.Dark ? ApplicationTheme.Light : ApplicationTheme.Dark);
                UpdateUI();
            }
            catch (COMException)
            {
                // Handle gracefully by retrying once
                var currentTheme = GetCurrentTheme();
                SetTheme(currentTheme == ApplicationTheme.Dark ? ApplicationTheme.Light : ApplicationTheme.Dark);
                UpdateUI();
            }
        }

        private void QuitButton_Click(object sender, RoutedEventArgs e)
        {
            trayIcon.Visible = false; // Hide the tray icon
            global::System.Windows.Forms.Application.Exit();
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
