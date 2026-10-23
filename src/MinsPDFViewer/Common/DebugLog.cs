using System;
using System.IO;
using System.Text;

namespace MinsPDFViewer
{
    internal static class DebugLog
    {
        private static readonly object SyncRoot = new object();

        public static void Append(string category, string message)
        {
            string prefix = string.IsNullOrWhiteSpace(category)
                ? string.Empty
                : $"[{category}] ";

            AppendRaw("debug.log", $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {prefix}{message}{Environment.NewLine}");
        }

        public static void AppendRaw(string fileName, string message)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                fileName = "debug.log";

            string safeFileName = Path.GetFileName(fileName);
            string logDirectory = GetLogDirectory();
            string path = Path.Combine(logDirectory, safeFileName);

            lock (SyncRoot)
            {
                Directory.CreateDirectory(logDirectory);
                File.AppendAllText(path, message, Encoding.UTF8);
            }
        }

        private static string GetLogDirectory()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (!string.IsNullOrWhiteSpace(appData))
                return Path.Combine(appData, "MinsPDFViewer", "logs");

            return Path.Combine(Path.GetTempPath(), "MinsPDFViewer", "logs");
        }
    }
}
