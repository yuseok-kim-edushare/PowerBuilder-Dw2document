using System.Drawing;

namespace yuseok.kim.dw2docs.Common.VirtualGrid
{
    public class FloatingCellOffset
    {
        public ColumnDefinition StartColumn { get; set; }
        public RowDefinition StartRow { get; set; }
        public short ColSpan { get; set; }
        public short RowSpan { get; set; }
        public Point StartOffset { get; set; }
        public Point EndOffset { get; set; }

        public FloatingCellOffset(ColumnDefinition startCol, RowDefinition startRow)
        {
            StartColumn = startCol;
            StartRow = startRow;
        }
    }
}
