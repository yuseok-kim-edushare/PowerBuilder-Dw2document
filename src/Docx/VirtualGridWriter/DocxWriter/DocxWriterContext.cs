using NPOI.XWPF.UserModel;

namespace yuseok.kim.dw2docs.Docx.VirtualGridWriter.DocxWriter;

/// <summary>
/// Holds the context required for rendering objects to a DOCX document using NPOI.
/// </summary>
public class DocxWriterContext
{
    public XWPFDocument Document { get; }
    public XWPFTable? CurrentTable { get; set; }
    public XWPFTableRow? CurrentRow { get; set; }
    // Add other relevant context properties as needed, e.g.:
    // public XWPFTableCell CurrentCell { get; set; }

    public DocxWriterContext(XWPFDocument document, XWPFTable? initialTable = null)
    {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        CurrentTable = initialTable;
    }
} 