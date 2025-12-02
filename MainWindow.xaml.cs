using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using Newtonsoft.Json;

namespace WordWebDAV
{
    public partial class MainWindow : Window
    {
        private WebDAVServer _server;
        private AppConfig _config = new AppConfig();

        public MainWindow()
        {
            InitializeComponent();
            LoadConfig();
            StartServer();
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
            switch (status)
            {
                case "running":
                    statusCircle.Fill = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(34, 197, 94)); // Green
                    txtStatus.Text = "  •  Đang chạy";
                    txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(34, 197, 94));
                    txtStatusCard.Text = "RUNNING";
                    txtStatusCard.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(34, 197, 94));
                    break;
                case "stopped":
                    statusCircle.Fill = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(250, 204, 21)); // Yellow
                    txtStatus.Text = "  •  Đang tạm dừng";
                    txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(250, 204, 21));
                    txtStatusCard.Text = "STOPPED";
                    txtStatusCard.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(250, 204, 21));
                    break;
                case "error":
                    statusCircle.Fill = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(239, 68, 68)); // Red
                    txtStatus.Text = "  •  Đang lỗi";
                    txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(239, 68, 68));
                    txtStatusCard.Text = "ERROR";
                    txtStatusCard.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(239, 68, 68));
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

        private void BtnMinimize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;  // Thu nhỏ xuống taskbar
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _server?.Stop();
                trayIcon?.Dispose();
            }
            catch { }
            
            // Force kill process
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
            // Không ẩn khi minimize, giữ trên taskbar
        }
    }
}
