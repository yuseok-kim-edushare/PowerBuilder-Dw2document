using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Abstractions;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Common;
using yuseok.kim.dw2docs.Docx.VirtualGridWriter.Renderers; // Added for RendererLocator
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
        private readonly RendererLocator<DocxWriterContext> _rendererLocator;
        private bool _documentInitialized = false;

        private XWPFTable? _workingTable; // This might become part of the context or managed differently

        public VirtualGridDocxWriter(VirtualGrid virtualGrid) : base(virtualGrid)
        {
            _document = new XWPFDocument();
            // Initialize context and locator later in InitDocument to ensure table exists
            // Need to handle potential nullability before InitDocument is called.
            // Let's initialize them partially here.
            _context = new DocxWriterContext(_document, null!); // Temp null! - Will be set in InitDocument
            _rendererLocator = new RendererLocator<DocxWriterContext>(_context);
            RegisterRenderers(); // Register renderers needed
        }

        private void RegisterRenderers()
        {
            // Register the actual DocxTextRenderer
            _rendererLocator.RegisterRenderer<DwText>((model, locator) => new DocxTextRenderer(_context, model, locator));

            // TODO: Register other actual Docx Renderers once created
            // _rendererLocator.RegisterRenderer<DwColumn>((model, locator) => new DocxColumnRenderer(_context, model, locator));
            // _rendererLocator.RegisterRenderer<DwBand>((model, locator) => new DocxBandRenderer(_context, model, locator));
            // ... register other renderers (Line, Rectangle, etc.)
        }


        [MemberNotNull(nameof(_workingTable))]
        private void InitDocument()
        {
            if (_documentInitialized) return;

            _workingTable = _document.CreateTable();
            // Set the table in the context AFTER it's created
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
            InitDocument(); // Ensures _workingTable and _context.CurrentTable are initialized

            if (data is null) // Keep data dictionary for potential future use by renderers
                return null;

            foreach (var row in rows)
            {
                var newRow = _context.CurrentTable.CreateRow();
                newRow.Height = row.Size;
                // Maybe set CurrentRow in context? -> _context.CurrentRow = newRow;

                foreach (var @object in row.Objects.Concat(row.FloatingObjects))
                {
                    // Get the appropriate renderer for the object's model type
                    var renderer = _rendererLocator.GetRenderer(@object.Model);

                    // How to pass the cell/row context? 
                    // Option 1: Modify context (e.g., _context.CurrentRow = newRow;)
                    // Option 2: Modify Render signature (e.g., Render(IVirtualCell cell, XWPFTableRow row))
                    // Let's try modifying the Render signature for clarity for now.
                    // This requires changing ObjectRendererBase and all implementing renderers.
                    // For now, we will just pass the cell object as before, 
                    // the renderer will need refinement to get the NPOI cell/row.
                    renderer.Render(@object);
                }
            }

            return null; // Return value might change depending on requirements

        }

        // Write method remains the same
        public override bool Write(string? path, out string? error)
        {
            error = null;

            if (path is null)
            {
                error = "No file specified";
                return false;
            }

            try
            {
                // Ensure all rows are processed before writing
                // The base class logic handles calling ProcessBands/WriteRows, so just write the document
                using var stream = File.Create(path);
                _document.Write(stream);
            }
            catch (IOException e)
            {
                error = e.ToString();
                return false;
            }

            return true;
        }
    }
}
