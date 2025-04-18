using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;

namespace yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;

public abstract class ExportedCellBase
{
    public VirtualCell Cell { get; init; }
    public DwObjectAttributesBase Attribute { get; init; }
    public Type AttributeType => Attribute.GetType();

    protected ExportedCellBase(
        VirtualCell cell,
        DwObjectAttributesBase attribute
        )
    {
        Cell = cell;
        Attribute = attribute;
    }
}
