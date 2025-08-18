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

        // Enhanced text rendering with positioning and styling information
        var positionInfo = $"at ({textAttributes.X},{textAttributes.Y}) size ({textAttributes.Width}x{textAttributes.Height})";
        var styleInfo = $"Font: {textAttributes.FontFace} {textAttributes.FontSize}pt, Align: {textAttributes.Alignment}";
        var bandInfo = $"Band: {textAttributes.Band}";
        
        Console.WriteLine($"DOCX Rendering Text: '{textAttributes.Text}' {positionInfo} {styleInfo} {bandInfo}");
        
        // In a full implementation, this would apply font, alignment, and positioning to the Word cell
        // For now, we're demonstrating that the enhanced attributes are being processed
        
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

        // Enhanced floating text rendering with positioning and styling information
        var positionInfo = $"at ({textAttributes.X},{textAttributes.Y}) size ({textAttributes.Width}x{textAttributes.Height})";
        var styleInfo = $"Font: {textAttributes.FontFace} {textAttributes.FontSize}pt Weight: {textAttributes.FontWeight}, Align: {textAttributes.Alignment}";
        var bandInfo = $"Band: {textAttributes.Band}";
        var decorationInfo = $"Underline: {textAttributes.Underline}, Italic: {textAttributes.Italics}, Strikethrough: {textAttributes.Strikethrough}";
        
        Console.WriteLine($"DOCX Rendering Floating Text: '{textAttributes.Text}' {positionInfo} {styleInfo} {bandInfo} {decorationInfo}");
        
        // In a full implementation, this would use Word's absolute positioning features
        // to place text at exact coordinates within the document
        
        return new ExportedFloatingCell(cell, attribute)
        {
            // Set any relevant output information
        };
    }
} 