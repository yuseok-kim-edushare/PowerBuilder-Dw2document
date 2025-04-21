using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Abstractions;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Common;
using yuseok.kim.dw2docs.Docx.VirtualGridWriter.Renderers;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System.Diagnostics.CodeAnalysis;
using yuseok.kim.dw2docs.Common.DwObjects; // Added for DwText

namespace yuseok.kim.dw2docs.Docx.VirtualGridWriter.DocxWriter
{
    public class VirtualGridDocxWriter : AbstractVirtualGridWriter
    {
        private readonly XWPFDocument _document;
        private readonly DocxWriterContext _context;
        private readonly RendererLocator _rendererLocator;
        private bool _documentInitialized = false;
        private string? _writeToPath;

        private XWPFTable _workingTable;

        public VirtualGridDocxWriter(VirtualGrid virtualGrid) : base(virtualGrid)
        {
            _document = new XWPFDocument();
            _workingTable = _document.CreateTable();
            _context = new DocxWriterContext(_document);
            _rendererLocator = new RendererLocator();
            RegisterRenderers();
        }

        public VirtualGridDocxWriter(VirtualGrid virtualGrid, string? writeToPath) : this(virtualGrid)
        {
            _writeToPath = writeToPath;
        }

        public void SetWritePath(string path)
        {
            _writeToPath = path;
        }

        private void RegisterRenderers()
        {
            // Register renderers for different attribute types
            _rendererLocator.RegisterRenderer(typeof(DwTextAttributes), new DocxTextRenderer(_context));
            
            // TODO: Register other renderers as they are implemented
        }

        private void InitDocument()
        {
            if (_documentInitialized) return;

            _context.CurrentTable = _workingTable;

            _documentInitialized = true;

            // Table setup remains the same
            _workingTable.GetCTTbl().tblPr.tblBorders.left.val = (ST_Border.none);
            _workingTable.GetCTTbl().tblPr.tblBorders.right.val = (ST_Border.none);
            _workingTable.GetCTTbl().tblPr.tblBorders.top.val = (ST_Border.none);
            _workingTable.GetCTTbl().tblPr.tblBorders.bottom.val = (ST_Border.none);
            _workingTable.GetCTTbl().tblPr.tblBorders.insideV.val = (ST_Border.none);
            _workingTable.GetCTTbl().tblPr.tblBorders.insideH.val = (ST_Border.none);

            for (int i = 0; i < VirtualGrid.Columns.Count; ++i)
            {
                _workingTable.AddNewCol();
            }
        }

        protected override IList<ExportedCellBase>? WriteRows(IList<RowDefinition> rows, IDictionary<string, DwObjectAttributesBase> data)
        {
            // Initialize document and make sure _workingTable is created
            if (!_documentInitialized)
            {
                InitDocument();
            }
            else if (_workingTable == null)
            {
                // Ensure _workingTable is initialized even if _documentInitialized is somehow true
                _workingTable = _document.CreateTable();
                _context.CurrentTable = _workingTable;
            }

            if (data is null)
                return null;

            var exportedCells = new List<ExportedCellBase>();

            foreach (var row in rows)
            {
                var newRow = _context.CurrentTable!.CreateRow();
                newRow.Height = row.Size;
                _context.CurrentRow = newRow;

                foreach (var cellObj in row.Objects.Concat(row.FloatingObjects))
                {
                    if (!data.TryGetValue(cellObj.Object.Name, out var attribute))
                    {
                        Console.WriteLine($"Warning: No attribute found for cell {cellObj.Object.Name}");
                        continue;
                    }

                    if (!attribute.IsVisible)
                    {
                        Console.WriteLine($"Cell {cellObj.Object.Name} is not visible, skipping");
                        continue;
                    }

                    var renderer = _rendererLocator.Find(attribute.GetType());
                    if (renderer == null)
                    {
                        Console.WriteLine($"Warning: No renderer found for attribute type {attribute.GetType().Name}");
                        continue;
                    }

                    // Process regular cells
                    if (cellObj is VirtualCell cell)
                    {
                        var columnIndex = cell.OwningColumn?.IndexOffset ?? 0;
                        XWPFTableCell? targetCell = null;
                        
                        try 
                        {
                            targetCell = newRow.GetCell(columnIndex);
                        }
                        catch
                        {
                            // Cell might not exist, create it
                            while (newRow.GetTableCells().Count <= columnIndex)
                            {
                                newRow.CreateCell();
                            }
                            targetCell = newRow.GetCell(columnIndex);
                        }

                        if (targetCell != null)
                        {
                            var result = renderer.Render(_context, cell, attribute, targetCell);
                            if (result != null)
                            {
                                exportedCells.Add(result);
                            }
                        }
                    }
                    // Process floating objects
                    else if (cellObj is FloatingVirtualCell floatingCell)
                    {
                        // For floating cells, we'll need to handle them differently
                        // Possibly create a paragraph in the document for floating elements
                        var result = renderer.Render(_context, floatingCell, attribute, newRow);
                        if (result != null)
                        {
                            exportedCells.Add(result);
                        }
                    }
                }
            }

            return exportedCells;
        }

        public override bool Write(string? path, out string? error)
        {
            error = null;

            string targetPath = path ?? _writeToPath ?? string.Empty;
            if (string.IsNullOrEmpty(targetPath))
            {
                error = "No file specified";
                return false;
            }

            try
            {
                // Process the grid to ensure all tables/rows/cells are created
                base.ProcessBands();
                
                // Write the document to the file
                Console.WriteLine($"[DOCX] Attempting to write to {targetPath}");
                using var stream = File.Create(targetPath);
                _document.Write(stream);
                Console.WriteLine($"[DOCX] Write complete. File exists: {File.Exists(targetPath)}");
                
                return true;
            }
            catch (IOException e)
            {
                error = e.ToString();
                return false;
            }
            catch (Exception ex)
            {
                error = $"Error writing document: {ex.Message}";
                return false;
            }
        }
    }
}
