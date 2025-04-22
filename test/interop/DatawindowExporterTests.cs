using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using yuseok.kim.dw2docs.Interop;

namespace yuseok.kim.dw2docs.test.Interop
{
    [TestClass]
    public class DatawindowExporterTests
    {
        private const string SimpleJson = @"{
  ""columns"": [
    { ""name"": ""Name"", ""width"": 100, ""type"": ""char(20)"", ""format"": """" },
    { ""name"": ""Age"", ""width"": 50, ""type"": ""int"", ""format"": """" }
  ],
  ""bands"": [
    { ""name"": ""detail"" }
  ],
  ""rows"": [
    { ""Name"": ""Alice"", ""Age"": ""30"" },
    { ""Name"": ""Bob"", ""Age"": ""25"" }
  ],
  ""cell_attributes"": {
    ""cell_0_0"": { ""text"": ""Alice"", ""is_visible"": true },
    ""cell_0_1"": { ""text"": ""30"", ""is_visible"": true },
    ""cell_1_0"": { ""text"": ""Bob"", ""is_visible"": true },
    ""cell_1_1"": { ""text"": ""25"", ""is_visible"": true }
  }
}";
        private string? _testOutputDir;

        [TestInitialize]
        public void Setup()
        {
            _testOutputDir = Path.Combine(TestContext.TestRunDirectory!, "TestOutput");
            Directory.CreateDirectory(_testOutputDir);
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (_testOutputDir != null && Directory.Exists(_testOutputDir))
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
            string outputPath = Path.Combine(_testOutputDir!, "test.xlsx");

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
            string outputPath = Path.Combine(_testOutputDir!, "test.docx");

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