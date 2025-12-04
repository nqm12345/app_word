using System;
using System.IO;
using System.Windows;
using Newtonsoft.Json;
using Microsoft.Win32;

namespace WordWebDAV
{
    public partial class MainWindow : Window
    {
        private WebDAVServer _server;
        private AppConfig _config = new AppConfig();
        private const string AutoStartKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "OfficeWebDAV";

        public MainWindow()
        {
            InitializeComponent();
            LoadConfig();
            ConfigureOfficeRegistry();
            SetupAutoStart();
            StartServer();
            SetupTrayIcon();
            Hide();
        }

        private void SetupTrayIcon()
        {
            try
            {
                trayIcon.IconSource = new System.Windows.Media.Imaging.BitmapImage(
                    new Uri("pack://application:,,,/app.ico", UriKind.Absolute));
            }
            catch { }
            
            trayIcon.ToolTipText = "Trình chỉnh sửa Office - Đang chạy";
            trayIcon.Visibility = Visibility.Visible;
            
            var menu = new System.Windows.Controls.ContextMenu();
            
            var statusItem = new System.Windows.Controls.MenuItem { Header = "✅ Đang chạy", IsEnabled = false };
            menu.Items.Add(statusItem);
            
            menu.Items.Add(new System.Windows.Controls.Separator());
            
            var exitItem = new System.Windows.Controls.MenuItem { Header = "❌ Thoát" };
            exitItem.Click += (s, e) => ExitApp();
            menu.Items.Add(exitItem);
            
            trayIcon.ContextMenu = menu;
        }

        private void ExitApp()
        {
            try
            {
                _server?.Stop();
                trayIcon?.Dispose();
            }
            catch { }
            Application.Current.Shutdown();
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
        }

        private void ConfigureOfficeRegistry()
        {
            try
            {
                string[] officeApps = { "Word", "Excel", "PowerPoint", "Visio" };
                // Hỗ trợ Office 2010 (14.0), 2013 (15.0), 2016/2019/2021/365 (16.0)
                string[] officeVersions = { "16.0", "15.0", "14.0" };
                
                foreach (var app in officeApps)
                {
                    foreach (var version in officeVersions)
                    {
                        ConfigureProtectedView($@"Software\Microsoft\Office\{version}\{app}\Security\ProtectedView");
                        
                        AddTrustedLocation($@"Software\Microsoft\Office\{version}\{app}\Security\Trusted Locations\Location99", "https://administrator.lifetex.vn:316");
                        AddTrustedLocation($@"Software\Microsoft\Office\{version}\{app}\Security\Trusted Locations\Location98", "https://vps-tcsg.lifetex.vn");
                        
                        if (!string.IsNullOrEmpty(_config.CompanyApiUrl) && 
                            !_config.CompanyApiUrl.Contains("localhost") &&
                            !_config.CompanyApiUrl.Contains("administrator.lifetex") &&
                            !_config.CompanyApiUrl.Contains("vps-tcsg.lifetex"))
                        {
                            AddTrustedLocation($@"Software\Microsoft\Office\{version}\{app}\Security\Trusted Locations\Location97", _config.CompanyApiUrl);
                        }
                    }
                }
            }
            catch { }
        }

        private void AddTrustedLocation(string keyPath, string path)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(keyPath);
                if (key != null)
                {
                    key.SetValue("Path", path, RegistryValueKind.String);
                    key.SetValue("AllowSubfolders", 1, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        private void ConfigureProtectedView(string keyPath)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(keyPath);
                if (key != null)
                {
                    key.SetValue("DisableInternetFilesInPV", 1, RegistryValueKind.DWord);
                    key.SetValue("DisableAttachementsInPV", 1, RegistryValueKind.DWord);
                    key.SetValue("DisableUnsafeLocationsInPV", 1, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        private void SetupAutoStart()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(AutoStartKey, true);
                if (key != null)
                {
                    string exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? "";
                    key.SetValue(AppName, $"\"{exePath}\"");
                }
            }
            catch { }
        }

        private void StartServer()
        {
            try
            {
                _server = new WebDAVServer(_config);
                _server.Start();
            }
            catch { }
        }
    }
}
