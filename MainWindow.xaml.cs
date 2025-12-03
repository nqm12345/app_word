using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using Newtonsoft.Json;
using Microsoft.Win32;

namespace WordWebDAV
{
    public partial class MainWindow : Window
    {
        private WebDAVServer _server;
        private AppConfig _config = new AppConfig();
        private DispatcherTimer _uptimeTimer;
        private DateTime _startTime;
        private int _requestCount = 0;
        private const string AutoStartKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "OfficeWebDAV";

        public MainWindow()
        {
            InitializeComponent();
            LoadConfig();
            LoadAutoStartSetting();
            ConfigureOfficeRegistry();
            StartServer();
            StartUptimeTimer();
            CreateDesktopShortcut();
        }

        private void StartUptimeTimer()
        {
            _startTime = DateTime.Now;
            _uptimeTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _uptimeTimer.Tick += (s, e) => UpdateUptime();
            _uptimeTimer.Start();
        }

        private void UpdateUptime()
        {
            var elapsed = DateTime.Now - _startTime;
            txtUptime.Text = elapsed.ToString(@"hh\:mm\:ss");
        }

        public void IncrementRequestCount()
        {
            _requestCount++;
            Dispatcher.Invoke(() => txtRequests.Text = _requestCount.ToString());
        }

        private void LoadAutoStartSetting()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(AutoStartKey);
                chkAutoStart.IsChecked = key?.GetValue(AppName) != null;
            }
            catch { }
        }

        private void ConfigureOfficeRegistry()
        {
            try
            {
                string[] officeApps = { "Word", "Excel", "PowerPoint", "Visio" };
                int configured = 0;
                foreach (var app in officeApps)
                {
                    // Tắt Protected View
                    if (ConfigureProtectedView($@"Software\Microsoft\Office\16.0\{app}\Security\ProtectedView"))
                        configured++;
                    ConfigureProtectedView($@"Software\Microsoft\Office\15.0\{app}\Security\ProtectedView");
                    
                    // Thêm server URL vào Trusted Locations
                    AddTrustedLocation($@"Software\Microsoft\Office\16.0\{app}\Security\Trusted Locations\Location99", "https://administrator.lifetex.vn:316");
                    AddTrustedLocation($@"Software\Microsoft\Office\16.0\{app}\Security\Trusted Locations\Location98", "https://vps-tcsg.lifetex.vn");
                    
                    // Thêm URL từ config (nếu khác)
                    if (!string.IsNullOrEmpty(_config.CompanyApiUrl) && 
                        !_config.CompanyApiUrl.Contains("localhost") &&
                        !_config.CompanyApiUrl.Contains("administrator.lifetex") &&
                        !_config.CompanyApiUrl.Contains("vps-tcsg.lifetex"))
                    {
                        AddTrustedLocation($@"Software\Microsoft\Office\16.0\{app}\Security\Trusted Locations\Location97", _config.CompanyApiUrl);
                    }
                }
                txtLog.Text = $"✅ Đã cấu hình Registry cho {configured} ứng dụng Office\n";
            }
            catch { }
        }

        private void AddTrustedLocation(string keyPath, string path)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(keyPath))
                {
                    if (key != null)
                    {
                        key.SetValue("Path", path, RegistryValueKind.String);
                        key.SetValue("AllowSubfolders", 1, RegistryValueKind.DWord);
                    }
                }
            }
            catch { }
        }

        private bool ConfigureProtectedView(string keyPath)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(keyPath))
                {
                    if (key != null)
                    {
                        key.SetValue("DisableInternetFilesInPV", 1, RegistryValueKind.DWord);
                        key.SetValue("DisableAttachementsInPV", 1, RegistryValueKind.DWord);
                        key.SetValue("DisableUnsafeLocationsInPV", 1, RegistryValueKind.DWord);
                        return true;
                    }
                }
            }
            catch { }
            return false;
        }

        private void CreateDesktopShortcut()
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string shortcutPath = Path.Combine(desktopPath, "Chỉnh sửa Office.lnk");
                if (File.Exists(shortcutPath)) return;
                string exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? "";
                string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.ico");
                Type? shellType = Type.GetTypeFromProgID("WScript.Shell");
                if (shellType == null) return;
                dynamic shell = Activator.CreateInstance(shellType)!;
                dynamic shortcut = shell.CreateShortcut(shortcutPath);
                shortcut.TargetPath = exePath;
                shortcut.WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory;
                shortcut.Description = "Chỉnh sửa file Office";
                if (File.Exists(iconPath)) shortcut.IconLocation = iconPath;
                shortcut.Save();
            }
            catch { }
        }

        private void LoadConfig()
        {
            try
            {
                var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
                if (File.Exists(configPath))
                {
                    var json = File.ReadAllText(configPath);
                    _config = JsonConvert.DeserializeObject<AppConfig>(json) ?? new AppConfig();
                }
            }
            catch { }
            txtPort.Text = _config.Port.ToString();
            txtApi.Text = _config.CompanyApiUrl;
        }

        private void StartServer()
        {
            try
            {
                // Stop old server first if exists
                if (_server != null)
                {
                    _server.Stop();
                    _server = null;
                }
                
                _server = new WebDAVServer(_config);
                _server.OnLog += Log;
                _server.Start();
                SetStatus("running");
            }
            catch (Exception ex)
            {
                Log("Error: " + ex.Message);
                SetStatus("error");
            }
        }

        private void StopServer()
        {
            if (_server != null)
            {
                _server.Stop();
                _server = null;
            }
            SetStatus("stopped");
        }

        private void SetStatus(string status)
        {
            var green = System.Windows.Media.Color.FromRgb(34, 197, 94);
            var yellow = System.Windows.Media.Color.FromRgb(250, 204, 21);
            var red = System.Windows.Media.Color.FromRgb(239, 68, 68);
            
            switch (status)
            {
                case "running":
                    statusCircle.Fill = new System.Windows.Media.SolidColorBrush(green);
                    txtStatus.Text = "  •  Đang chạy";
                    txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(green);
                    txtStatusCard.Text = "RUNNING";
                    txtStatusCard.Foreground = new System.Windows.Media.SolidColorBrush(green);
                    txtBtnIcon.Text = "⏸";
                    break;
                case "stopped":
                    statusCircle.Fill = new System.Windows.Media.SolidColorBrush(yellow);
                    txtStatus.Text = "  •  Tạm dừng";
                    txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(yellow);
                    txtStatusCard.Text = "PAUSED";
                    txtStatusCard.Foreground = new System.Windows.Media.SolidColorBrush(yellow);
                    txtBtnIcon.Text = "▶";
                    break;
                case "error":
                    statusCircle.Fill = new System.Windows.Media.SolidColorBrush(red);
                    txtStatus.Text = "  •  Lỗi";
                    txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(red);
                    txtStatusCard.Text = "ERROR";
                    txtStatusCard.Foreground = new System.Windows.Media.SolidColorBrush(red);
                    txtBtnIcon.Text = "▶";
                    break;
            }
        }

        private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
                DragMove();
        }

        private void Log(string msg)
        {
            Dispatcher.Invoke(() => { txtLog.Text += msg + "\n"; logScroller.ScrollToEnd(); });
        }

        private void BtnStartStop_Click(object sender, RoutedEventArgs e)
        {
            if (_server != null && _server.IsRunning) 
                StopServer(); 
            else 
                StartServer();
        }

        private void BtnRestart_Click(object sender, RoutedEventArgs e)
        {
            Log("🔄 Restarting server...");
            StopServer();
            _startTime = DateTime.Now;
            _requestCount = 0;
            txtRequests.Text = "0";
            StartServer();
        }

        private void BtnCopyApi_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Clipboard.SetText(_config.CompanyApiUrl);
                Log("📋 API URL copied to clipboard");
            }
            catch { }
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            txtLog.Text = "";
        }

        private void ChkAutoStart_Changed(object sender, RoutedEventArgs e)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(AutoStartKey, true);
                if (key == null) return;
                
                if (chkAutoStart.IsChecked == true)
                {
                    string exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? "";
                    key.SetValue(AppName, $"\"{exePath}\"");
                    Log("✅ Auto-start enabled");
                }
                else
                {
                    key.DeleteValue(AppName, false);
                    Log("⏹ Auto-start disabled");
                }
            }
            catch { }
        }

        private void BtnMinimize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _uptimeTimer?.Stop();
                _server?.Stop();
                trayIcon?.Dispose();
            }
            catch { }
            System.Diagnostics.Process.GetCurrentProcess().Kill();
        }

        private void TrayIcon_Click(object sender, RoutedEventArgs e)
        {
            Show();
            WindowState = WindowState.Normal;
            Activate();
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            // Keep on taskbar when minimized
        }
    }
}
