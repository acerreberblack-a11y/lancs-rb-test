using System;
using System.Globalization;
using System.IO;

namespace LandocsRobot.Logging
{
    internal sealed class FileLogger
    {
        private readonly string _logDirectory;
        private string _logFilePath;

        public FileLogger(string logDirectory)
        {
            _logDirectory = logDirectory;
            Directory.CreateDirectory(_logDirectory);
            UpdateLogFile(DateTime.Now);
        }

        public LogLevel LogLevel { get; private set; } = LogLevel.Info;

        public static string CreateDefaultLogDirectory(string baseDirectory)
        {
            string logDirectory = Path.Combine(baseDirectory, "logs");
            Directory.CreateDirectory(logDirectory);
            return logDirectory;
        }

        public void SetLogLevel(LogLevel level) => LogLevel = level;

        public void UpdateLogFile(DateTime date)
        {
            _logFilePath = Path.Combine(_logDirectory, $"{date:yyyy-MM-dd}.log");
        }

        public void Log(LogLevel level, string message, string ticketContext)
        {
            if (level > LogLevel)
            {
                return;
            }

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            string context = string.IsNullOrWhiteSpace(ticketContext) ? string.Empty : $"[{ticketContext}] ";
            string formattedMessage = $"{timestamp} [{level}] {context}{message}";

            if (!string.IsNullOrWhiteSpace(_logFilePath))
            {
                try
                {
                    File.AppendAllText(_logFilePath, formattedMessage + Environment.NewLine);
                }
                catch (IOException ex)
                {
                    Console.Error.WriteLine($"Не удалось записать сообщение в лог: {ex.Message}");
                }
            }

            Console.WriteLine(formattedMessage);
        }

        public void CleanOldLogs(int retentionDays)
        {
            if (retentionDays <= 0)
            {
                return;
            }

            foreach (string file in Directory.EnumerateFiles(_logDirectory, "*.log"))
            {
                try
                {
                    DateTime creationTime = File.GetCreationTime(file);
                    if (creationTime < DateTime.Now.AddDays(-retentionDays))
                    {
                        File.Delete(file);
                        Log(LogLevel.Info, $"Лог-файл {Path.GetFileName(file)} удален", null);
                    }
                }
                catch (Exception ex)
                {
                    Log(LogLevel.Error, $"Ошибка при удалении файла лога {Path.GetFileName(file)}: {ex.Message}", null);
                }
            }
        }
    }
}
