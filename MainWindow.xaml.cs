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
            // SetupAutoStart(); // Đã bỏ - để người dùng tự chọn khi cài đặt
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

            // Hiện thông báo khi app khởi động thành công
            MessageBox.Show("ChinhSuaOffice đã khởi động thành công!\n\nỨng dụng sẽ chạy ẩn trong system tray (góc phải dưới màn hình).", 
                "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
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
                // Hỗ trợ Office 2013 (15.0), 2016/2019/2021/365 (16.0)
                string[] officeVersions = { "16.0", "15.0" };
                
                foreach (var app in officeApps)
                {
                    foreach (var version in officeVersions)
                    {
                        // 1. Tắt Protected View
                        ConfigureProtectedView($@"Software\Microsoft\Office\{version}\{app}\Security\ProtectedView");
                        
                        // 2. Bật "Allow Trusted Locations on my network"
                        EnableNetworkTrustedLocations($@"Software\Microsoft\Office\{version}\{app}\Security\Trusted Locations");
                        
                        // 3. Cấu hình Security chung - tắt cảnh báo
                        ConfigureSecurity($@"Software\Microsoft\Office\{version}\{app}\Security");
                        
                        // 4. Thêm Trusted Locations
                        AddTrustedLocation($@"Software\Microsoft\Office\{version}\{app}\Security\Trusted Locations\Location99", "https://administrator.lifetex.vn:316");
                        AddTrustedLocation($@"Software\Microsoft\Office\{version}\{app}\Security\Trusted Locations\Location98", "https://lifeoffice-tcsg.lifetex.vn");
                        AddTrustedLocation($@"Software\Microsoft\Office\{version}\{app}\Security\Trusted Locations\Location100", "http://localhost:1901");
                        AddTrustedLocation($@"Software\Microsoft\Office\{version}\{app}\Security\Trusted Locations\Location101", "http://127.0.0.1:1901");
                        
                        // Thêm các URL từ config.json nếu khác URL mặc định
                        var apiUrls = _config.GetApiUrls();
                        int locationIndex = 90;
                        foreach (var apiUrl in apiUrls)
                        {
                            if (!apiUrl.Contains("localhost") &&
                                !apiUrl.Contains("127.0.0.1") &&
                                !apiUrl.Contains("administrator.lifetex") &&
                                !apiUrl.Contains("lifeoffice-tcsg.lifetex"))
                            {
                                AddTrustedLocation($@"Software\Microsoft\Office\{version}\{app}\Security\Trusted Locations\Location{locationIndex}", apiUrl);
                                locationIndex++;
                            }
                        }
                    }
                }
                
                // 5. Thêm Trusted Protocol cho ms-word, ms-excel, ms-powerpoint
                foreach (var version in officeVersions)
                {
                    AddTrustedProtocol($@"Software\Microsoft\Office\{version}\Common\Security\Trusted Protocols\All Applications\ms-word:");
                    AddTrustedProtocol($@"Software\Microsoft\Office\{version}\Common\Security\Trusted Protocols\All Applications\ms-excel:");
                    AddTrustedProtocol($@"Software\Microsoft\Office\{version}\Common\Security\Trusted Protocols\All Applications\ms-powerpoint:");
                }
                
                // 6. Thêm localhost vào Internet Explorer Trusted Sites (Office 2013 dùng IE Security Zones)
                AddToTrustedSites("localhost");
                AddToTrustedSites("127.0.0.1");
                
                // 7. Tắt cảnh báo cho Intranet zone
                ConfigureSecurityZones();
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

        private void EnableNetworkTrustedLocations(string keyPath)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(keyPath);
                if (key != null)
                {
                    // Bật "Allow Trusted Locations on my network" - cần cho Office 2013+
                    key.SetValue("AllowNetworkLocations", 1, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        private void ConfigureSecurity(string keyPath)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(keyPath);
                if (key != null)
                {
                    // Tắt cảnh báo "Some files contain viruses"
                    key.SetValue("DisableHyperlinkWarning", 1, RegistryValueKind.DWord);
                    // Tắt kiểm tra file validation
                    key.SetValue("EnableFileValidation", 0, RegistryValueKind.DWord);
                    // Cho phép mở file từ network
                    key.SetValue("AllowDDE", 2, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        private void AddTrustedProtocol(string keyPath)
        {
            try
            {
                // Tạo key rỗng để đánh dấu protocol là trusted
                Registry.CurrentUser.CreateSubKey(keyPath);
            }
            catch { }
        }

        private void AddToTrustedSites(string site)
        {
            try
            {
                // Thêm site vào IE Trusted Sites (Zone 2)
                // Office 2013 dùng IE Security Zones để kiểm tra URL
                string keyPath = $@"Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\{site}";
                using var key = Registry.CurrentUser.CreateSubKey(keyPath);
                if (key != null)
                {
                    // http = Zone 2 (Trusted Sites)
                    key.SetValue("http", 2, RegistryValueKind.DWord);
                    key.SetValue("https", 2, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        private void ConfigureSecurityZones()
        {
            try
            {
                // Zone 1 = Local Intranet, Zone 2 = Trusted Sites
                // Tắt cảnh báo "launching programs and files" cho cả 2 zones
                
                // Zone 1 (Intranet)
                using (var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1"))
                {
                    if (key != null)
                    {
                        // 2001 = Download signed ActiveX controls (0=enable)
                        // 1806 = Launching applications and unsafe files (0=enable)
                        key.SetValue("1806", 0, RegistryValueKind.DWord);
                    }
                }
                
                // Zone 2 (Trusted Sites)
                using (var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2"))
                {
                    if (key != null)
                    {
                        key.SetValue("1806", 0, RegistryValueKind.DWord);
                    }
                }
                
                // Thêm localhost vào Intranet zone
                using (var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range100"))
                {
                    if (key != null)
                    {
                        key.SetValue(":Range", "127.0.0.1", RegistryValueKind.String);
                        key.SetValue("http", 1, RegistryValueKind.DWord);
                    }
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
                bool started = _server.Start();
                
                if (!started)
                {
                    MessageBox.Show(
                        "Không thể khởi động WebDAV Server!\n\n" +
                        "Có thể do:\n" +
                        "1. Port 1901 đang bị sử dụng\n" +
                        "2. Thiếu quyền Administrator\n" +
                        "3. Windows Firewall chặn\n\n" +
                        "Thử chạy app với quyền Administrator.",
                        "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Lỗi khởi động server:\n\n{ex.Message}",
                    "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
