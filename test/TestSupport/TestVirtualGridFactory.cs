using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using yuseok.kim.dw2docs.Common.DwObjects;
using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.Enums;
using yuseok.kim.dw2docs.Common.VirtualGrid;

namespace yuseok.kim.dw2docs.test.TestSupport
{
    /// <summary>
    /// Test support class for creating test grid objects for testing
    /// </summary>
    public static class TestVirtualGridFactory
    {
        /// <summary>
        /// Creates a minimal but functional grid for testing
        /// </summary>
        public static VirtualGrid CreateSimpleVirtualGrid()
        {
            // Setup rows, columns, band rows
            var rows = new List<RowDefinition>();
            var columns = new List<ColumnDefinition>();
            var bandRows = new List<BandRows>();
            
            // Add columns to grid
            var column1 = new ColumnDefinition { Size = 100 };
            var column2 = new ColumnDefinition { Size = 100 };
            columns.Add(column1);
            columns.Add(column2);
            
            // Add row to grid
            var row = new RowDefinition { Size = 30, BandName = "detail" };
            rows.Add(row);
            
            // Create a text object for grid
            var textObject = new Dw2DObject(
                "text1",      // name
                "text",       // text
                10,           // x
                10,           // y
                80,           // width
                20);          // height
            
            // Create cell and add to row
            var textCell = new VirtualCell(textObject);
            textCell.OwningColumn = column1;
            textCell.OwningRow = row;
            row.Objects.Add(textCell);
            
            // Create cell repository
            var cellRepo = CreateCellRepository(rows);
            
            // Create band rows definition
            var detailBand = new BandRows("detail", rows);
            bandRows.Add(detailBand);
            
            // Create the VirtualGrid directly using internal constructor
            var grid = new VirtualGrid(rows, columns, bandRows, cellRepo, DwType.Default);
            
            // Add attributes to the grid
            var attributes = new Dictionary<string, DwObjectAttributesBase>
            {
                { textObject.Name, new DwTextAttributes { Text = "Sample Text", IsVisible = true } }
            };
            
            SetAttributes(grid, attributes);
            
            return grid;
        }
        
        /// <summary>
        /// Creates a more complex grid with multiple cell types
        /// </summary>
        public static VirtualGrid CreateComplexVirtualGrid()
        {
            // Setup rows, columns, band rows
            var rows = new List<RowDefinition>();
            var columns = new List<ColumnDefinition>();
            var bandRows = new List<BandRows>();
            
            // Add columns to grid
            var column1 = new ColumnDefinition { Size = 100 };
            var column2 = new ColumnDefinition { Size = 100 };
            var column3 = new ColumnDefinition { Size = 100 };
            columns.Add(column1);
            columns.Add(column2);
            columns.Add(column3);
            
            // Add header row
            var headerRow = new RowDefinition { Size = 30, BandName = "header" };
            rows.Add(headerRow);
            
            // Create header cell
            var headerObject = new Dw2DObject(
                "header",     // name
                "Header",     // text
                10,           // x
                10,           // y
                280,          // width (spans all columns)
                20);          // height
            
            var headerCell = new VirtualCell(headerObject);
            headerCell.OwningColumn = column1;
            headerCell.ColumnSpan = 3;
            headerCell.OwningRow = headerRow;
            headerRow.Objects.Add(headerCell);
            
            // Add data row
            var dataRow = new RowDefinition { Size = 25, BandName = "detail" };
            rows.Add(dataRow);
            
            // Add text cell
            var textObject = new Dw2DObject(
                "text1",      // name
                "Cell Text",  // text
                10,           // x
                40,           // y
                80,           // width
                20);          // height
            
            var textCell = new VirtualCell(textObject);
            textCell.OwningColumn = column1;
            textCell.OwningRow = dataRow;
            dataRow.Objects.Add(textCell);
            
            // Add button cell
            var buttonObject = new Dw2DObject(
                "button1",    // name
                "Click Me",   // text
                120,          // x
                40,           // y
                80,           // width
                20);          // height
            
            var buttonCell = new VirtualCell(buttonObject);
            buttonCell.OwningColumn = column2;
            buttonCell.OwningRow = dataRow;
            dataRow.Objects.Add(buttonCell);
            
            // Create cell repository
            var cellRepo = CreateCellRepository(rows);
            
            // Create band rows definitions
            var headerBand = new BandRows("header", new List<RowDefinition> { headerRow });
            var detailBand = new BandRows("detail", new List<RowDefinition> { dataRow });
            bandRows.Add(headerBand);
            bandRows.Add(detailBand);
            
            // Create the VirtualGrid directly using internal constructor
            var grid = new VirtualGrid(rows, columns, bandRows, cellRepo, DwType.Default);
            
            // Add attributes to the grid
            var attributes = new Dictionary<string, DwObjectAttributesBase>
            {
                { headerObject.Name, new DwTextAttributes { Text = "Header Text", IsVisible = true } },
                { textObject.Name, new DwTextAttributes { Text = "Sample Text", IsVisible = true } },
                { buttonObject.Name, new DwButtonAttributes { Text = "Click Me", IsVisible = true } }
            };
            
            SetAttributes(grid, attributes);
            
            return grid;
        }
        
        /// <summary>
        /// Creates a properly populated VirtualCellRepository from the given rows
        /// </summary>
        private static VirtualCellRepository CreateCellRepository(IList<RowDefinition> rows)
        {
            var cellsByName = new Dictionary<string, VirtualCell>();
            var rowsByBand = new Dictionary<string, IDictionary<string, VirtualCell>>();
            var cellsByX = new SortedList<int, IList<VirtualCell>>();
            var cellsByY = new SortedList<int, IList<VirtualCell>>();
            var cells = new HashSet<VirtualCell>();
            
            // Collect all cells from rows
            foreach (var row in rows)
            {
                foreach (var cell in row.Objects)
                {
                    // Add to cells collection
                    cells.Add(cell);
                    
                    // Add to cellsByName
                    cellsByName[cell.Object.Name] = cell;
                    
                    // Add to rowsByBand
                    string bandName = row.BandName ?? "detail";
                    if (!rowsByBand.ContainsKey(bandName))
                    {
                        rowsByBand[bandName] = new Dictionary<string, VirtualCell>();
                    }
                    rowsByBand[bandName][cell.Object.Name] = cell;
                    
                    // Add to cellsByX
                    int x = cell.X;
                    if (!cellsByX.ContainsKey(x))
                    {
                        cellsByX[x] = new List<VirtualCell>();
                    }
                    cellsByX[x].Add(cell);
                    
                    // Add to cellsByY
                    int y = cell.Y;
                    if (!cellsByY.ContainsKey(y))
                    {
                        cellsByY[y] = new List<VirtualCell>();
                    }
                    cellsByY[y].Add(cell);
                }
            }
            
            return new VirtualCellRepository(cellsByName, rowsByBand, cellsByX, cellsByY, cells);
        }
        
        /// <summary>
        /// Set attributes on a VirtualGrid
        /// </summary>
        private static void SetAttributes(VirtualGrid grid, Dictionary<string, DwObjectAttributesBase> attributes)
        {
            // Try to find and set _controlAttributes field directly
            // The InternalsVisibleTo attribute should give us access
            var controlAttributesField = typeof(VirtualGrid).GetField("_controlAttributes", 
                System.Reflection.BindingFlags.NonPublic | 
                System.Reflection.BindingFlags.Instance | 
                System.Reflection.BindingFlags.Public);
            
            if (controlAttributesField != null)
            {
                // Get the current value
                var currentAttributes = controlAttributesField.GetValue(grid) as Dictionary<string, DwObjectAttributesBase>;
                
                if (currentAttributes == null)
                {
                    // If field exists but is null, create a new dictionary
                    currentAttributes = new Dictionary<string, DwObjectAttributesBase>();
                    controlAttributesField.SetValue(grid, currentAttributes);
                }
                
                // Add our attributes to the existing ones
                foreach (var attr in attributes)
                {
                    currentAttributes[attr.Key] = attr.Value;
                }
            }
        }
    }
} 