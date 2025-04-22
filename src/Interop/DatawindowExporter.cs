using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using yuseok.kim.dw2docs.Common.DwObjects;
using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.Enums;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Abstractions;
using yuseok.kim.dw2docs.Docx.VirtualGridWriter.DocxWriter;
using yuseok.kim.dw2docs.Xlsx.VirtualGridWriter.XlsxWriter;
using yuseok.kim.dw2docs.Common.Utils;

namespace yuseok.kim.dw2docs.Interop
{
    /// <summary>
    /// A simplified exporter class that PowerBuilder can use directly for exporting datawindows to Excel and Word
    /// </summary>
    public class DatawindowExporter
    {

        /// <summary>
        /// Export datawindow data in JSON format to Excel
        /// </summary>
        /// <param name="jsonData">Datawindow data in JSON format</param>
        /// <param name="outputPath">Path where the Excel file will be saved</param>
        /// <param name="sheetName">Optional sheet name</param>
        /// <returns>Success message or error message</returns>
        public string ExportToExcel(string jsonData, string outputPath, string sheetName = "Sheet1")
        {
            try
            {
                // Create a virtual grid from the JSON data
                var grid = CreateVirtualGridFromJson(jsonData);
                if (grid == null)
                {
                    return "Error: Failed to create virtual grid from JSON data";
                }

                // Create Excel writer
                var builder = new VirtualGridXlsxWriterBuilder
                {
                    WriteToPath = outputPath
                };

                // Build the writer
                var writer = builder.Build(grid, out string? error);
                if (writer == null)
                {
                    return $"Error: Failed to create Excel writer - {error}";
                }

                // Set path and write to Excel
                if (writer is VirtualGridXlsxWriter xlsxWriter)
                {
                    xlsxWriter.SetWritePath(outputPath);
                }

                // Write to Excel
                bool success = writer.WriteEntireGrid(sheetName, out error);
                
                // Dispose writer if it implements IDisposable
                if (writer is IDisposable disposable)
                {
                    disposable.Dispose();
                }

                if (!success)
                {
                    return $"Error: {error}";
                }

                return $"Success: Excel file created at {outputPath}";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}\nStackTrace: {ex.StackTrace}\nInnerException: {ex.InnerException}";
            }
        }

        /// <summary>
        /// Export datawindow data in JSON format to Word
        /// </summary>
        /// <param name="jsonData">Datawindow data in JSON format</param>
        /// <param name="outputPath">Path where the Word file will be saved</param>
        /// <returns>Success message or error message</returns>
        public string ExportToWord(string jsonData, string outputPath)
        {
            try
            {
                // Create a virtual grid from the JSON data
                var grid = CreateVirtualGridFromJson(jsonData);
                if (grid == null)
                {
                    return "Error: Failed to create virtual grid from JSON data";
                }

                // Create Word writer
                var builder = new VirtualGridDocxWriterBuilder
                {
                    WriteToPath = outputPath
                };

                // Build the writer
                var writer = builder.Build(grid, out string? error);
                if (writer == null)
                {
                    return $"Error: Failed to create Word writer - {error}";
                }

                // Write to Word
                bool success = writer.WriteEntireGrid(outputPath, out error);
                
                // Dispose writer if it implements IDisposable
                if (writer is IDisposable disposable)
                {
                    disposable.Dispose();
                }

                if (!success)
                {
                    return $"Error: {error}";
                }

                // Debug: Print ControlAttributes keys
                var controlAttributesProp = typeof(VirtualGrid).GetProperty("ControlAttributes");
                var controlAttributes = controlAttributesProp?.GetValue(grid) as IDictionary<string, object>;
                if (controlAttributes != null)
                {
                    FileLogger.LogToFile("[ExportToWord] ControlAttributes keys: " + string.Join(", ", controlAttributes.Keys));
                }
                else
                {
                    FileLogger.LogToFile("[ExportToWord] ControlAttributes is null");
                }

                return $"Success: Word file created at {outputPath}";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// Helper method to create a virtual grid from JSON data
        /// </summary>
        private VirtualGrid CreateVirtualGridFromJson(string jsonData)
        {
            try
            {
                var jsonDoc = JsonDocument.Parse(jsonData);
                var root = jsonDoc.RootElement;

                // Initialize grid components
                var rows = new List<RowDefinition>();
                var columns = new List<ColumnDefinition>();
                var bands = new List<BandRows>();
                
                // Dictionary to store cell values by cell name
                var cellValues = new Dictionary<string, string>();
                
                // Process rows from JSON
                if (root.TryGetProperty("rows", out var rowsElement))
                {
                    // First pass: create column definitions
                    var columnNames = new HashSet<string>();
                    foreach (var row in rowsElement.EnumerateArray())
                    {
                        foreach (var prop in row.EnumerateObject())
                        {
                            columnNames.Add(prop.Name);
                        }
                    }
                    
                    // Create column definitions
                    int colIndex = 0;
                    foreach (var colName in columnNames)
                    {
                        columns.Add(new ColumnDefinition { Size = 100 });
                        colIndex++;
                    }
                    
                    // Second pass: create rows and add cells
                    int rowIndex = 0;
                    foreach (var row in rowsElement.EnumerateArray())
                    {
                        // Create row
                        var rowDef = new RowDefinition { Size = 30, BandName = "detail" };
                        rows.Add(rowDef);
                        
                        // Process cells
                        colIndex = 0;
                        foreach (var colName in columnNames)
                        {
                            if (row.TryGetProperty(colName, out var value))
                            {
                                // Get the cell value as string
                                string? cellValue = value.GetString();
                                string cellName = $"cell_{rowIndex}_{colIndex}";
                                
                                // Store the value for later use
                                cellValues[cellName] = cellValue ?? "";
                                
                                // Create a text object for the cell
                                var textObject = new Dw2DObject(
                                    cellName,      // name
                                    "detail",      // band
                                    colIndex * 100, // x
                                    rowIndex * 30,  // y
                                    100,           // width
                                    30);           // height
                                
                                // Create cell and add to row
                                var cell = new VirtualCell(textObject);
                                cell.OwningColumn = columns[colIndex];
                                cell.OwningRow = rowDef;
                                rowDef.Objects.Add(cell);
                            }
                            colIndex++;
                        }
                        rowIndex++;
                    }
                }
                
                // Create band rows definition
                var detailBand = new BandRows("detail", rows);
                bands.Add(detailBand);
                
                // Create cell repository
                var cellRepo = CreateCellRepository(rows);
                
                // Create the VirtualGrid
                var grid = new VirtualGrid(rows, columns, bands, cellRepo, DwType.Default);
                
                // Add attributes to the grid
                var attributes = new Dictionary<string, DwObjectAttributesBase>();
                
                // Create attributes based on stored cell values
                foreach (var cellEntry in cellValues)
                {
                    attributes[cellEntry.Key] = new DwTextAttributes 
                    { 
                        Text = cellEntry.Value,
                        IsVisible = true
                    };
                }
                
                // Set attributes using reflection
                var method = typeof(VirtualGrid).GetMethod("SetAttributes", 
                    System.Reflection.BindingFlags.NonPublic | 
                    System.Reflection.BindingFlags.Instance);
                method?.Invoke(grid, new object[] { attributes });

                // Also set ControlAttributes property if available
                var controlAttributesProp = typeof(VirtualGrid).GetProperty("ControlAttributes");
                if (controlAttributesProp != null && controlAttributesProp.CanWrite)
                {
                    controlAttributesProp.SetValue(grid, attributes);
                }
                
                return grid;
            }
            catch (Exception ex)
            {
                FileLogger.LogToFile($"[CreateVirtualGridFromJson] Error: {ex.Message}\n{ex.StackTrace}", ex);
                throw; // rethrow so ExportToExcel can catch it
            }
        }
        
        /// <summary>
        /// Helper method to create a cell repository from rows
        /// </summary>
        private VirtualCellRepository CreateCellRepository(IList<RowDefinition> rows)
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
            
            // Create VirtualCellRepository using constructor with parameters
            var constructor = typeof(VirtualCellRepository).GetConstructor(
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance,
                null,
                new Type[] {
                    typeof(Dictionary<string, VirtualCell>),
                    typeof(Dictionary<string, IDictionary<string, VirtualCell>>),
                    typeof(SortedList<int, IList<VirtualCell>>),
                    typeof(SortedList<int, IList<VirtualCell>>),
                    typeof(HashSet<VirtualCell>)
                },
                null);
                
            if (constructor == null)
            {
                throw new InvalidOperationException("Could not find constructor for VirtualCellRepository");
            }
            
            return (VirtualCellRepository)constructor.Invoke(new object[] {
                cellsByName,
                rowsByBand,
                cellsByX,
                cellsByY,
                cells
            });
        }
    }
} 