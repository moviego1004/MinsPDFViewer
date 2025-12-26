using System;
using System.IO;
using System.Windows;

namespace MinsPDFViewer
{
    public partial class App : Application
    {
        public App()
        {
            // 1. UI 스레드에서 발생하는 예외 처리
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;

            // 2. 백그라운드 스레드(Task)에서 발생하는 예외 처리
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            System.Threading.Tasks.TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            LogException(e.Exception, "UI Thread Error");
            e.Handled = true; // true로 하면 죽지 않고 무시할 수도 있지만, 상태가 불안정할 수 있음
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            LogException(e.ExceptionObject as Exception, "Non-UI Thread Error (Critical)");
        }

        private void TaskScheduler_UnobservedTaskException(object sender, System.Threading.Tasks.UnobservedTaskExceptionEventArgs e)
        {
            LogException(e.Exception, "Background Task Error");
            e.SetObserved();
        }

        private void LogException(Exception? ex, string source)
        {
            string logFile = "crash_dump.txt";
            string message = $"[{DateTime.Now}] [{source}]\n{ex?.ToString() ?? "Unknown Error"}\n---------------------------------\n";

            try
            {
                File.AppendAllText(logFile, message);
                // 치명적인 에러일 경우 사용자에게 알림
                MessageBox.Show($"오류가 발생했습니다.\n로그 파일: {logFile}\n내용: {ex?.Message}", "오류 발생", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch { }
        }
    }
}