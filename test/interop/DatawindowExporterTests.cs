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

        private const string EnhancedJson = @"{
  ""columns"": [
    { ""name"": ""Name"", ""width"": 100, ""type"": ""char(20)"", ""format"": """" },
    { ""name"": ""Age"", ""width"": 50, ""type"": ""int"", ""format"": """" }
  ],
  ""bands"": [
    { ""name"": ""header"" },
    { ""name"": ""detail"" },
    { ""name"": ""footer"" }
  ],
  ""objects"": [
    {
      ""name"": ""title_text"",
      ""type"": ""text"",
      ""x"": 100,
      ""y"": 50,
      ""width"": 300,
      ""height"": 60,
      ""band"": ""header"",
      ""text"": ""Employee Report"",
      ""alignment"": ""center"",
      ""font_face"": ""Arial"",
      ""font_height"": 14,
      ""font_weight"": 700
    },
    {
      ""name"": ""separator_line"",
      ""type"": ""line"",
      ""x1"": 0,
      ""y1"": 120,
      ""x2"": 500,
      ""y2"": 120,
      ""band"": ""header"",
      ""pen_width"": 2,
      ""pen_color"": 0
    },
    {
      ""name"": ""total_compute"",
      ""type"": ""compute"",
      ""x"": 200,
      ""y"": 10,
      ""width"": 100,
      ""height"": 30,
      ""band"": ""footer"",
      ""expression"": ""Count(Name)"",
      ""format"": ""#,##0"",
      ""alignment"": ""right""
    }
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

        [TestMethod]
        public void ExportToExcel_WithVisualObjects_CreatesFileAndReturnsSuccess()
        {
            // Arrange
            var exporter = new DatawindowExporter();
            string outputPath = Path.Combine(_testOutputDir!, "enhanced_test.xlsx");

            // Act
            string result = exporter.ExportToExcel(EnhancedJson, outputPath);

            // Assert
            Assert.IsTrue(result.StartsWith("Success"), $"Expected success, got: {result}");
            Assert.IsTrue(File.Exists(outputPath), "Enhanced Excel file was not created.");
            
            // Additional validation - check file size is reasonable
            var fileInfo = new FileInfo(outputPath);
            Assert.IsTrue(fileInfo.Length > 0, "Excel file should have content");
            Console.WriteLine($"Enhanced Excel file created. Size: {fileInfo.Length} bytes");
        }

        [TestMethod]
        public void ExportToWord_WithVisualObjects_CreatesFileAndReturnsSuccess()
        {
            // Arrange
            var exporter = new DatawindowExporter();
            string outputPath = Path.Combine(_testOutputDir!, "enhanced_test.docx");

            // Act
            string result = exporter.ExportToWord(EnhancedJson, outputPath);

            // Assert
            Assert.IsTrue(result.StartsWith("Success"), $"Expected success, got: {result}");
            Assert.IsTrue(File.Exists(outputPath), "Enhanced Word file was not created.");
            
            // Additional validation - check file size is reasonable
            var fileInfo = new FileInfo(outputPath);
            Assert.IsTrue(fileInfo.Length > 0, "Word file should have content");
            Console.WriteLine($"Enhanced Word file created. Size: {fileInfo.Length} bytes");
        }
    }
} 