using System;
using System.IO;

namespace yuseok.kim.dw2docs.Common.Utils
{
    public static class FileLogger
    {
        private const string LogFilePath = @"C:\temp\Dw2Doc_ExcelError.log";

        public static void LogToFile(string message, Exception? ex = null)
        {
            try
            {
                string logContent = $"[{DateTime.Now}] {message}";
                if (ex != null)
                {
                    logContent += $"\nException Type: {ex.GetType().FullName}\nMessage: {ex.Message}\nStackTrace:\n{ex.StackTrace}";
                }
                logContent += "\n---------------------------------\n";
                File.AppendAllText(LogFilePath, logContent);
            }
            catch (Exception logEx)
            {
                Console.WriteLine($"!!! Failed to write to log file {LogFilePath}: {logEx.Message}");
            }
        }
    }
} 