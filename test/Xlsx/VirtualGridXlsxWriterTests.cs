using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Xlsx.VirtualGridWriter.XlsxWriter;
using yuseok.kim.dw2docs.test.TestSupport;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Collections.Generic;

namespace yuseok.kim.dw2docs.test.Xlsx
{
    [TestClass]
    public class VirtualGridXlsxWriterTests
    {
        private const string TestOutputDir = "TestOutput";
        private readonly string _testFilePath;

        public VirtualGridXlsxWriterTests()
        {
            // Create test output directory if it doesn't exist
            if (!Directory.Exists(TestOutputDir))
            {
                Directory.CreateDirectory(TestOutputDir);
            }

            _testFilePath = Path.Combine(TestOutputDir, "writer_test.xlsx");
            
            // Make sure we start with a clean file for each test
            if (File.Exists(_testFilePath))
            {
                File.Delete(_testFilePath);
            }
        }

        [TestMethod]
        public void Write_WithValidGrid_CreatesXlsxFile()
        {
            // Arrange
            var grid = TestVirtualGridFactory.CreateSimpleVirtualGrid();
            var builder = new VirtualGridXlsxWriterBuilder
            {
                WriteToPath = _testFilePath
            };

            // Act
            var writer = builder.Build(grid, out string? error);
            var result = writer!.Write(null, out error);

            // Assert
            Assert.IsTrue(result);
            Assert.IsNull(error);
            Assert.IsTrue(File.Exists(_testFilePath));
            
            // Verify file structure
            VerifyExcelFile(_testFilePath);
        }

        [TestMethod]
        public void Write_WithSheetName_CreatesXlsxFileWithSpecifiedSheetName()
        {
            // Arrange
            var grid = TestVirtualGridFactory.CreateSimpleVirtualGrid();
            var builder = new VirtualGridXlsxWriterBuilder
            {
                WriteToPath = _testFilePath
            };
            const string sheetName = "TestSheet";

            // Act
            var writer = builder.Build(grid, out string? error);
            var result = writer!.Write(sheetName, out error);

            // Assert
            Assert.IsTrue(result);
            Assert.IsNull(error);
            Assert.IsTrue(File.Exists(_testFilePath));
            
            // Verify the sheet name
            using (var workbook = new XSSFWorkbook(_testFilePath))
            {
                var sheet = workbook.GetSheet(sheetName);
                Assert.IsNotNull(sheet);
            }
        }

        [TestMethod]
        public void Build_WithNullPath_ReturnsErrorAndNullWriter()
        {
            // Arrange
            var grid = TestVirtualGridFactory.CreateSimpleVirtualGrid();
            var builder = new VirtualGridXlsxWriterBuilder
            {
                WriteToPath = null
            };

            // Act
            var writer = builder.Build(grid, out string? error);

            // Assert
            Assert.IsNull(writer);
            Assert.IsNotNull(error);
            Assert.AreEqual("Must specify a path", error);
        }

        [TestMethod]
        public void Build_WithAppendAndNonExistentFile_ReturnsErrorAndNullWriter()
        {
            // Arrange
            var grid = TestVirtualGridFactory.CreateSimpleVirtualGrid();
            var nonExistentPath = Path.Combine(TestOutputDir, "non_existent.xlsx");
            
            // Ensure the file doesn't exist
            if (File.Exists(nonExistentPath))
            {
                File.Delete(nonExistentPath);
            }

            var builder = new VirtualGridXlsxWriterBuilder
            {
                WriteToPath = nonExistentPath,
                Append = true,
                AppendToSheetName = "TestSheet"
            };

            // Act
            var writer = builder.Build(grid, out string? error);

            // Assert
            Assert.IsNull(writer);
            Assert.IsNotNull(error);
            Assert.IsTrue(error.Contains("does not exist"));
        }

        [TestMethod]
        public void Build_WithAppendAndNoSheetName_ReturnsErrorAndNullWriter()
        {
            // Arrange
            // First create a file to append to
            {
                var firstGrid = TestVirtualGridFactory.CreateSimpleVirtualGrid();
                var firstBuilder = new VirtualGridXlsxWriterBuilder
                {
                    WriteToPath = _testFilePath
                };
                var firstWriter = firstBuilder.Build(firstGrid, out string? error1);
                firstWriter!.Write(null, out error1);
            }

            // Now try to append without a sheet name
            var grid2 = TestVirtualGridFactory.CreateSimpleVirtualGrid();
            var appendBuilder = new VirtualGridXlsxWriterBuilder
            {
                WriteToPath = _testFilePath,
                Append = true,
                AppendToSheetName = null
            };

            // Act
            var writer = appendBuilder.Build(grid2, out string? error);

            // Assert
            Assert.IsNull(writer);
            Assert.IsNotNull(error);
            Assert.AreEqual("Did not specify a sheet name", error);
        }
        
        [TestMethod]
        public void Append_ToExistingFile_CreatesNewSheet()
        {
            // Arrange
            // First create a file with Sheet1
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
            var grid2 = TestVirtualGridFactory.CreateSimpleVirtualGrid();
            var builder2 = new VirtualGridXlsxWriterBuilder
            {
                WriteToPath = _testFilePath,
                Append = true,
                AppendToSheetName = "Sheet2"
            };
            
            var writer2 = builder2.Build(grid2, out string? error);
            var result = writer2!.Write(null, out error);

            // Assert
            Assert.IsTrue(result);
            Assert.IsNull(error);
            
            using (var workbook = new XSSFWorkbook(_testFilePath))
            {
                Assert.AreEqual(2, workbook.NumberOfSheets);
                Assert.IsNotNull(workbook.GetSheet("Sheet1"));
                Assert.IsNotNull(workbook.GetSheet("Sheet2"));
            }
        }

        /// <summary>
        /// Helper method to verify Excel file structure
        /// </summary>
        private void VerifyExcelFile(string filePath)
        {
            using (var workbook = new XSSFWorkbook(filePath))
            {
                Assert.AreEqual(1, workbook.NumberOfSheets);
                var sheet = workbook.GetSheetAt(0);
                Assert.IsNotNull(sheet);
                
                // Default sheet name should be "Sheet1"
                Assert.AreEqual("Sheet1", sheet.SheetName);
                
                // Verify sheet has content
                Assert.IsTrue(sheet.PhysicalNumberOfRows > 0);
            }
        }
    }
} 