using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using yuseok.kim.dw2docs.Interop;

namespace yuseok.kim.dw2docs.test.Interop
{
    [TestClass]
    public class DatawindowExporterTests
    {
        private const string SimpleJson = "{\n  \"rows\": [\n    { \"Name\": \"Alice\", \"Age\": \"30\" },\n    { \"Name\": \"Bob\", \"Age\": \"25\" }\n  ]\n}";
        private string _testOutputDir;

        [TestInitialize]
        public void Setup()
        {
            _testOutputDir = Path.Combine(TestContext.TestRunDirectory, "TestOutput");
            Directory.CreateDirectory(_testOutputDir);
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (Directory.Exists(_testOutputDir))
            {
                foreach (var file in Directory.GetFiles(_testOutputDir))
                {
                    try { File.Delete(file); } catch { }
                }
            }
        }

        public TestContext TestContext { get; set; }

        [TestMethod]
        public void ExportToExcel_CreatesFileAndReturnsSuccess()
        {
            // Arrange
            var exporter = new DatawindowExporter();
            string outputPath = Path.Combine(_testOutputDir, "test.xlsx");

            // Act
            string result = exporter.ExportToExcel(SimpleJson, outputPath);

            // Assert
            Assert.IsTrue(result.StartsWith("Success"), $"Expected success, got: {result}");
            Assert.IsTrue(File.Exists(outputPath), "Excel file was not created.");
        }

        [TestMethod]
        public void ExportToWord_CreatesFileAndReturnsSuccess()
        {
            // Arrange
            var exporter = new DatawindowExporter();
            string outputPath = Path.Combine(_testOutputDir, "test.docx");

            // Act
            string result = exporter.ExportToWord(SimpleJson, outputPath);

            // Assert
            if (!File.Exists(outputPath))
            {
                Console.WriteLine($"ExportToWord result: {result}");
            }
            else
            {
                var fileInfo = new FileInfo(outputPath);
                Console.WriteLine($"Word file created. Size: {fileInfo.Length} bytes");
            }
            Assert.IsTrue(result.StartsWith("Success"), $"Expected success, got: {result}");
            Assert.IsTrue(File.Exists(outputPath), "Word file was not created.");
        }
    }
} 