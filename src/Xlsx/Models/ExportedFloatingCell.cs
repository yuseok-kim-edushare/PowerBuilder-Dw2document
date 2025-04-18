using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using NPOI.XSSF.UserModel;

namespace yuseok.kim.dw2docs.Xlsx.Models;

public class ExportedFloatingCell : ExportedCellBase
{
    public XSSFShape? OutputShape { get; set; }

    public ExportedFloatingCell(
        FloatingVirtualCell cell,
        DwObjectAttributesBase attribute) : base(cell, attribute)
    {
    }
}
