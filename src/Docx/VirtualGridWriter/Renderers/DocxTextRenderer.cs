using yuseok.kim.dw2docs.Common.DwObjects;
using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Common;
using yuseok.kim.dw2docs.Docx.VirtualGridWriter.DocxWriter;
using NPOI.XWPF.UserModel;

namespace yuseok.kim.dw2docs.Docx.VirtualGridWriter.Renderers;

public class DocxTextRenderer : ObjectRendererBase
{
    private readonly DocxWriterContext _context;

    public DocxTextRenderer(DocxWriterContext context)
    {
        _context = context;
    }

    public override ExportedCellBase? Render(object context, VirtualCell cell, DwObjectAttributesBase attribute, object renderTarget)
    {
        if (attribute is not DwTextAttributes textAttributes)
        {
            return null;
        }

        // TODO: Determine the target XWPFTableCell and apply text formatting
        Console.WriteLine($"DOCX Rendering Text: {textAttributes.Text} (Placeholder)");
        
        // Return a basic exported cell
        return new ExportedCell(cell, attribute)
        {
            // Set any relevant output information
        };
    }

    public override ExportedCellBase? Render(object context, FloatingVirtualCell cell, DwObjectAttributesBase attribute, object renderTarget)
    {
        if (attribute is not DwTextAttributes textAttributes)
        {
            return null;
        }

        // TODO: Implement floating text rendering for DOCX
        Console.WriteLine($"DOCX Rendering Floating Text: {textAttributes.Text} (Placeholder)");
        
        // Return a basic exported floating cell
        return new ExportedFloatingCell(cell, attribute)
        {
            // Set any relevant output information
        };
    }
} 