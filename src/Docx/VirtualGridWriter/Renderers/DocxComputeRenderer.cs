using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Common;
using yuseok.kim.dw2docs.Docx.VirtualGridWriter.DocxWriter;

namespace yuseok.kim.dw2docs.Docx.VirtualGridWriter.Renderers;

public class DocxComputeRenderer : ObjectRendererBase
{
    private readonly DocxWriterContext _context;

    public DocxComputeRenderer(DocxWriterContext context)
    {
        _context = context;
    }

    public override ExportedCellBase? Render(object context, VirtualCell cell, DwObjectAttributesBase attribute, object renderTarget)
    {
        if (attribute is not DwComputeAttributes computeAttributes)
        {
            return null;
        }

        // Render computed field with expression and formatting
        Console.WriteLine($"DOCX Rendering Compute: Expression='{computeAttributes.Expression}', Format='{computeAttributes.FormatString}'");
        
        return new ExportedCell(cell, attribute)
        {
            // Set any relevant output information
        };
    }

    public override ExportedCellBase? Render(object context, FloatingVirtualCell cell, DwObjectAttributesBase attribute, object renderTarget)
    {
        if (attribute is not DwComputeAttributes computeAttributes)
        {
            return null;
        }

        // Render floating computed field
        Console.WriteLine($"DOCX Rendering Floating Compute: Expression='{computeAttributes.Expression}' at ({computeAttributes.X},{computeAttributes.Y}) in band {computeAttributes.Band}");
        
        return new ExportedFloatingCell(cell, attribute)
        {
            // Set any relevant output information
        };
    }
}