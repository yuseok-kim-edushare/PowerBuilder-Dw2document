using yuseok.kim.dw2docs.Common.DwObjects;
using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.test.TestSupport;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;

namespace yuseok.kim.dw2docs.test.Common
{
    [TestClass]
    public class VirtualGridTests
    {
        [TestMethod]
        public void CreateSimpleVirtualGrid_HasCorrectStructure()
        {
            // Act
            var grid = TestVirtualGridFactory.CreateSimpleVirtualGrid();
            
            // Assert
            Assert.IsNotNull(grid);
            
            // Verify core properties
            Assert.AreEqual(1, grid.Rows.Count);
            Assert.AreEqual(2, grid.Columns.Count);
            Assert.AreEqual(1, grid.BandRows.Count);
            
            // Verify row
            var row = grid.Rows[0];
            Assert.AreEqual(30, row.Size);
            Assert.AreEqual("detail", row.BandName);
            Assert.AreEqual(1, row.Objects.Count);
            
            // Verify cell
            var cell = row.Objects[0];
            Assert.AreEqual("text1", cell.Object.Name);
            Assert.AreEqual(10, cell.X);
            Assert.AreEqual(10, cell.Y);
            Assert.AreEqual(80, cell.Width);
            Assert.AreEqual(20, cell.Height);
        }
        
        [TestMethod]
        public void CreateComplexVirtualGrid_HasCorrectStructure()
        {
            // Act
            var grid = TestVirtualGridFactory.CreateComplexVirtualGrid();
            
            // Assert
            Assert.IsNotNull(grid);
            
            // Verify core structure
            Assert.AreEqual(2, grid.Rows.Count);
            Assert.AreEqual(3, grid.Columns.Count);
            Assert.AreEqual(2, grid.BandRows.Count);
            
            // Verify bands exist
            var headerBand = grid.BandRows.FirstOrDefault(b => b.Name == "header");
            var detailBand = grid.BandRows.FirstOrDefault(b => b.Name == "detail");
            Assert.IsNotNull(headerBand);
            Assert.IsNotNull(detailBand);
            
            // Verify header row
            var headerRow = headerBand.Rows[0];
            Assert.AreEqual(30, headerRow.Size);
            Assert.AreEqual("header", headerRow.BandName);
            Assert.AreEqual(1, headerRow.Objects.Count);
            
            // Verify header cell
            var headerCell = headerRow.Objects[0];
            Assert.AreEqual("header", headerCell.Object.Name);
            Assert.AreEqual(3, headerCell.ColumnSpan);
            
            // Verify detail row
            var detailRow = detailBand.Rows[0];
            Assert.AreEqual(25, detailRow.Size);
            Assert.AreEqual("detail", detailRow.BandName);
            Assert.AreEqual(2, detailRow.Objects.Count);
            
            // Verify text cell
            var textCell = detailRow.Objects.FirstOrDefault(c => c.Object.Name == "text1");
            Assert.IsNotNull(textCell);
            Assert.AreEqual(10, textCell.X);
            Assert.AreEqual(40, textCell.Y);
            
            // Verify button cell
            var buttonCell = detailRow.Objects.FirstOrDefault(c => c.Object.Name == "button1");
            Assert.IsNotNull(buttonCell);
            Assert.AreEqual(120, buttonCell.X);
            Assert.AreEqual(40, buttonCell.Y);
        }
    }
} 