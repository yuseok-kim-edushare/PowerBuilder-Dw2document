using NPOI.XWPF.UserModel;

namespace yuseok.kim.dw2docs.Docx.VirtualGridWriter.DocxWriter;

/// <summary>
/// Holds the context required for rendering objects to a DOCX document using NPOI.
/// </summary>
public class DocxWriterContext
{
    public XWPFDocument Document { get; }
    public XWPFTable CurrentTable { get; set; } // Assuming rendering happens within a table structure
    // Add other relevant context properties as needed, e.g.:
    // public XWPFTableRow CurrentRow { get; set; }
    // public XWPFTableCell CurrentCell { get; set; }

    public DocxWriterContext(XWPFDocument document, XWPFTable initialTable)
    {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        CurrentTable = initialTable ?? throw new ArgumentNullException(nameof(initialTable));
    }
} 