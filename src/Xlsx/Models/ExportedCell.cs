using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using NPOI.SS.UserModel;

namespace yuseok.kim.dw2docs.Xlsx.Models;

public class ExportedCell : ExportedCellBase
{
    public ICell? OutputCell { get; set; }

    public ExportedCell(VirtualCell cell,
        DwObjectAttributesBase attribute) : base(cell, attribute)
    {
    }
}
