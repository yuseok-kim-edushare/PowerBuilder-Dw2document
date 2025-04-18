using yuseok.skim.dw2docs.Common.DwObjects;

namespace yuseok.kim.dw2docs.Common.VirtualGrid
{
    public class FloatingVirtualCell : VirtualCell
    {
        public FloatingCellOffset Offset { get; }

        public FloatingVirtualCell(Dw2DObject @object, FloatingCellOffset offset) : base(@object)
        {
            Offset = offset;
        }

        public static FloatingVirtualCell FromVirtualCell(VirtualCell cell, FloatingCellOffset offset)
            => new(cell.Object, offset)
            {
                OwningColumn = cell.OwningColumn,
                OwningRow = offset.StartRow,
            };
    }
}
