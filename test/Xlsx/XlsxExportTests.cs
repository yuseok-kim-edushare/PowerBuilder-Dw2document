using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Xlsx.VirtualGridWriter.XlsxWriter;
using yuseok.kim.dw2docs.test.TestSupport;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.XSSF.UserModel;
using System.IO;

namespace yuseok.kim.dw2docs.test.Xlsx
{
    [TestClass]
    public class XlsxExportTests
    {
        private const string TestOutputDir = "TestOutput";
        private readonly string _testFilePath;

        public XlsxExportTests()
        {
            // Create test output directory if it doesn't exist
            if (!Directory.Exists(TestOutputDir))
            {
                Directory.CreateDirectory(TestOutputDir);
            }

            _testFilePath = Path.Combine(TestOutputDir, "export_test.xlsx");
            
            // Make sure we start with a clean file
            if (File.Exists(_testFilePath))
            {
                File.Delete(_testFilePath);
            }
        }

        [TestMethod]
        public void ExportSimpleGrid_CreatesValidXlsxFile()
        {
            // Arrange - Create VirtualGrid and builder
            var grid = TestVirtualGridFactory.CreateSimpleVirtualGrid();
            var builder = new VirtualGridXlsxWriterBuilder
            {
                WriteToPath = _testFilePath
            };

            // Act - Create writer and export grid to Excel
            var writer = builder.Build(grid, out string? error);
            var result = writer!.Write("SimpleSheet", out error);

            // Assert - Verify basic export succeeded
            Assert.IsTrue(result);
            Assert.IsNull(error);
            Assert.IsTrue(File.Exists(_testFilePath));
            
            // Verify file content
            VerifyBasicExcelStructure(_testFilePath, "SimpleSheet");
        }

        [TestMethod]
        public void ExportComplexGrid_CreatesCorrectXlsxStructure()
        {
            // Arrange - Create complex VirtualGrid and builder
            var grid = TestVirtualGridFactory.CreateComplexVirtualGrid();
            var builder = new VirtualGridXlsxWriterBuilder
            {
                WriteToPath = _testFilePath
            };

            // Act - Create writer and export grid to Excel
            var writer = builder.Build(grid, out string? error);
            var result = writer!.Write("ComplexSheet", out error);

            // Assert - Verify export succeeded
            Assert.IsTrue(result);
            Assert.IsNull(error);
            Assert.IsTrue(File.Exists(_testFilePath));
            
            // Verify complex file structure
            using (var workbook = new XSSFWorkbook(_testFilePath))
            {
                var sheet = workbook.GetSheet("ComplexSheet");
                Assert.IsNotNull(sheet);
                
                // Verify sheet has content
                Assert.IsTrue(sheet.PhysicalNumberOfRows > 0);
                
                // Verify header row
                var headerRow = sheet.GetRow(0);
                Assert.IsNotNull(headerRow);
                Assert.IsNotNull(headerRow.GetCell(0));
                
                // Verify data row (if available)
                if (sheet.PhysicalNumberOfRows > 1)
                {
                    var dataRow = sheet.GetRow(1);
                    Assert.IsNotNull(dataRow);
                }
            }
        }

        [TestMethod]
        public void AppendToExistingFile_AddsNewSheet()
        {
            // Arrange - First create a file with one sheet
            {
                var grid1 = TestVirtualGridFactory.CreateSimpleVirtualGrid();
                var builder1 = new VirtualGridXlsxWriterBuilder
                {
                    WriteToPath = _testFilePath
                };
                var writer1 = builder1.Build(grid1, out string? error1);
                writer1!.Write("Sheet1", out error1);
            }
            
            // Act - Now append a second sheet
            var grid2 = TestVirtualGridFactory.CreateComplexVirtualGrid();
            var builder2 = new VirtualGridXlsxWriterBuilder
            {
                WriteToPath = _testFilePath,
                Append = true,
                AppendToSheetName = "Sheet2"
            };
            
            var writer2 = builder2.Build(grid2, out string? error);
            var result = writer2!.Write(null, out error);
            
            // Assert - Verify append succeeded
            Assert.IsTrue(result);
            Assert.IsNull(error);
            
            // Verify file contains both sheets
            using (var workbook = new XSSFWorkbook(_testFilePath))
            {
                Assert.AreEqual(2, workbook.NumberOfSheets);
                Assert.IsNotNull(workbook.GetSheet("Sheet1"));
                Assert.IsNotNull(workbook.GetSheet("Sheet2"));
            }
        }
        
        /// <summary>
        /// Helper method to verify basic Excel file structure
        /// </summary>
        private void VerifyBasicExcelStructure(string filePath, string sheetName)
        {
            using (var workbook = new XSSFWorkbook(filePath))
            {
                // Verify sheet exists
                var sheet = workbook.GetSheet(sheetName);
                Assert.IsNotNull(sheet);
                
                // Verify sheet has content
                Assert.IsTrue(sheet.PhysicalNumberOfRows > 0);
                
                // Verify at least one row exists
                var row = sheet.GetRow(0);
                Assert.IsNotNull(row);
                
                // Verify at least one cell exists
                Assert.IsTrue(row.PhysicalNumberOfCells > 0);
            }
        }
    }
} 