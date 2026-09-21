using System;
using System.IO;
using System.Text;
using System.Threading;

namespace OMNIX.Core.Logging
{
    /// <summary>
    /// Layer 9 (Diagnostics): local file logger.
    /// Logs live in %LOCALAPPDATA%\OMNIX\logs\*.log and rotate at ~5 MB.
    /// NEVER log API keys or protected document content (Ironclad Rule 6).
    /// </summary>
    public static class Logger
    {
        private static readonly object Gate = new object();
        private static string _baseDir;
        private const long MaxLogBytes = 5L * 1024 * 1024;

        public static string BaseDir
        {
            get
            {
                if (_baseDir == null)
                {
                    string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                    _baseDir = Path.Combine(localAppData, "OMNIX");
                }
                return _baseDir;
            }
        }

        public static string LogsDir { get { return Path.Combine(BaseDir, "logs"); } }

        public static string InstallLogPath { get { return Path.Combine(LogsDir, "install-debug.log"); } }

        public static void Startup(string message) { Write("startup-debug", message); }
        public static void Ui(string message) { Write("ui-debug", message); }
        public static void Gateway(string message) { Write("gateway-debug", message); }
        public static void RuntimeJourney(string message) { Write("runtime-journey", message); }
        public static void RuntimeEventJson(string json) { WriteRawJsonLine("runtime-events", json); }
        public static void Install(string message) { Write("install-debug", message); }
        public static void Error(string source, string message, Exception ex) { Write(source, message + (ex == null ? "" : " | " + ex.GetType().Name + ": " + ex.Message + " | stack: " + ex.StackTrace)); }

        public static void Write(string streamName, string message)
        {
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(LogsDir);
                    string path = Path.Combine(LogsDir, streamName + ".log");
                    RotateIfNeeded(path);
                    var line = new StringBuilder();
                    line.Append(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff")).Append("Z [")
                        .Append(Thread.CurrentThread.ManagedThreadId).Append("] ").Append(message)
                        .Append(Environment.NewLine);
                    File.AppendAllText(path, line.ToString(), Encoding.UTF8);
                }
            }
            catch
            {
                // Logging must never crash the add-in.
            }
        }

        private static void WriteRawJsonLine(string streamName, string json)
        {
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(LogsDir);
                    string path = Path.Combine(LogsDir, streamName + ".jsonl");
                    RotateIfNeeded(path);
                    File.AppendAllText(path, (json ?? "{}") + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch
            {
                // Diagnostics must never crash the add-in.
            }
        }

        private static void RotateIfNeeded(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    var fi = new FileInfo(path);
                    if (fi.Length > MaxLogBytes)
                    {
                        string backup = path + ".1";
                        if (File.Exists(backup)) File.Delete(backup);
                        File.Move(path, backup);
                    }
                }
            }
            catch
            {
                // best effort
            }
        }
    }
}
