using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Abstractions;
using NPOI.XSSF.UserModel;
using System.IO;

namespace yuseok.kim.dw2docs.Xlsx.VirtualGridWriter.XlsxWriter
{
    public class VirtualGridXlsxWriterBuilder : IVirtualGridWriterBuilder
    {
        private const string LogFilePath = @"C:\temp\Dw2Doc_ExcelError.log";

        private static void LogToFile(string message, Exception? ex = null)
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

        public bool Append { get; set; }
        public string? WriteToPath { get; set; }
        public string? AppendToSheetName { get; set; }

        public AbstractVirtualGridWriter? Build(VirtualGrid grid,
            out string? error)
        {
            LogToFile($"VirtualGridXlsxWriterBuilder.Build entered. Path: {WriteToPath}, Append: {Append}, Sheet: {AppendToSheetName}");
            error = null;

            if (string.IsNullOrEmpty(WriteToPath))
            {
                error = "Must specify a path";
                LogToFile("Error in Build: WriteToPath is null or empty.");
                return null;
            }

            XSSFWorkbook? workbook = null;
            FileStream? stream = null;
            string targetSheetName = AppendToSheetName ?? "Sheet1"; // Default sheet name if not appending or not specified
            bool fileExists = File.Exists(WriteToPath);

            try
            {
                if (Append)
                {
                    if (!fileExists)
                    {
                        error = $"Append specified, but file does not exist: {WriteToPath}";
                        LogToFile(error);
                        return null;
                    }
                    if (string.IsNullOrEmpty(AppendToSheetName))
                    {
                        error = "Append specified, but AppendToSheetName is null or empty";
                        LogToFile(error);
                        return null;
                    }

                    LogToFile($"Opening existing file for append: {WriteToPath}");
                    // Open stream to read existing workbook
                    using (var readStream = new FileStream(WriteToPath!, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) // Allow readwrite share temporarily
                    {
                        workbook = new XSSFWorkbook(readStream);
                    }
                    LogToFile("Existing workbook read.");

                    if (workbook.GetSheet(AppendToSheetName) != null)
                    {
                        error = $"Append specified, but sheet '{AppendToSheetName}' already exists in {WriteToPath}";
                        LogToFile(error);
                        workbook.Close(); // Close workbook if sheet exists
                        return null;
                    }
                    targetSheetName = AppendToSheetName; // Already set above, but explicit for clarity

                    // Open the final stream for writing. NPOI needs to rewrite the whole file.
                    // Use FileMode.Create to overwrite the existing file with the modified workbook.
                    stream = new FileStream(WriteToPath!, FileMode.Create, FileAccess.Write, FileShare.None);
                    LogToFile("Opened file stream for writing (FileMode.Create for append rewrite).");
                }
                else // Create new file or overwrite existing if Append is false
                {
                    LogToFile($"Creating new workbook. File will be created/overwritten at: {WriteToPath}");
                    workbook = CreateWorkbook();
                    stream = new FileStream(WriteToPath!, FileMode.Create, FileAccess.Write, FileShare.None);
                    targetSheetName = AppendToSheetName ?? "Sheet1"; // Use provided name or default
                    LogToFile($"Created stream for new/overwritten file (FileMode.Create). Target sheet: {targetSheetName}");
                }

                LogToFile("Attempting to create VirtualGridXlsxWriter instance.");
                // Pass the created workbook, stream, and target sheet name
                // The writer will now be responsible for disposing the stream and workbook
                return new VirtualGridXlsxWriter(grid, workbook, stream, targetSheetName!);
            }
            catch (Exception e)
            {
                error = e.Message;
                LogToFile("!!! EXCEPTION during stream/workbook creation or writer instantiation in Build", e);
                // Clean up resources if creation failed
                stream?.Close();
                stream?.Dispose();
                workbook?.Close(); // Close workbook on error
                return null;
            }
        }

        public AbstractVirtualGridWriter? BuildFromTemplate(VirtualGrid grid,
            string filePath,
            out string? error)
        {
            LogToFile($"VirtualGridXlsxWriterBuilder.BuildFromTemplate entered. FilePath: {filePath}");
            error = null;
            throw new NotImplementedException();
            // Future implementation will go here
        }

        protected virtual XSSFWorkbook CreateWorkbook()
        {
            string tempFile = Path.Combine(Path.GetTempPath(), $"excel_temp_{Guid.NewGuid()}.xlsx");
            try
            {
                // Use a hard-coded minimal Excel file
                byte[] minimalExcelFile = MinimalXlsxBytes.GetBytes();
                File.WriteAllBytes(tempFile, minimalExcelFile);
                using (var fs = new FileStream(tempFile, FileMode.Open, FileAccess.Read))
                {
                    var result = new XSSFWorkbook(fs);
                    // Clean up
                    try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
                    return result;
                }
            }
            catch (Exception ex)
            {
                LogToFile("!!! EXCEPTION during workbook creation in CreateWorkbook", ex);
                throw new InvalidOperationException(
                    "All attempts to create an Excel workbook have failed. " +
                    "This appears to be a bug in the NPOI library with your environment. " + 
                    "Consider using a pre-created template file or a different Excel library. " +
                    "Error: " + ex.Message, ex);
            }
        }
    }
}
