using yuseok.kim.dw2docs.Common.DwObjects;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Common;
using yuseok.kim.dw2docs.Docx.VirtualGridWriter.DocxWriter;
using NPOI.XWPF.UserModel;

namespace yuseok.kim.dw2docs.Docx.VirtualGridWriter.Renderers;

public class DocxTextRenderer : ObjectRendererBase<DocxWriterContext, DwText>
{
    public DocxTextRenderer(DocxWriterContext context, DwText model, RendererLocator<DocxWriterContext> locator)
        : base(context, model, locator)
    {
    }

    public override void Render(IVirtualCell cell)
    {
        // TODO: Determine the target XWPFTableCell. 
        // We likely need the current XWPFTableRow passed here or accessible from the context.
        // Assume we get the row: XWPFTableRow currentRow = ...; 
        // Assume we get the cell index: int cellIndex = cell.OwningColumn?.IndexOffset ?? -1;
        // if (cellIndex == -1) { /* Handle floating? */ return; }
        // XWPFTableCell targetCell = currentRow.GetCell(cellIndex);

        // Placeholder: Get the text content
        var text = Model.Attributes.Text; // Or potentially cell.CalculatedValue?

        // TODO: Add text to the targetCell using NPOI
        // Example:
        // var paragraph = targetCell.AddParagraph(); // Or get existing?
        // var run = paragraph.CreateRun();
        // run.SetText(text);
        // Apply formatting based on Model.Attributes (font, color, size, alignment etc.)

        Console.WriteLine($"DOCX Rendering Text: {text} (Placeholder)"); // Placeholder
    }
} 