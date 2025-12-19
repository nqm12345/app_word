using System;
using System.Threading;
using System.Windows;

namespace WordWebDAV
{
    public partial class App : Application
    {
        private static Mutex _mutex;

        protected override void OnStartup(StartupEventArgs e)
        {
            const string appName = "ChinhSuaOffice_SingleInstance";
            bool createdNew;

            _mutex = new Mutex(true, appName, out createdNew);

            if (!createdNew)
            {
                // App đã chạy rồi, thoát ngay
                MessageBox.Show("ChinhSuaOffice đang chạy rồi!\n\nKiểm tra icon trong system tray (góc phải dưới màn hình).", 
                    "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                Environment.Exit(0);
                return;
            }

            base.OnStartup(e);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            if (_mutex != null)
            {
                _mutex.ReleaseMutex();
                _mutex.Dispose();
            }
            base.OnExit(e);
        }
    }
}

