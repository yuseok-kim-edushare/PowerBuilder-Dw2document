namespace yuseok.kim.dw2docs.Common.VirtualGrid
{
    internal class ColumnOverlap
    {
        public IList<ColumnDefinition> Columns { get; set; }
        public int OffsetLeft { get; set; }
        public int OffsetRight { get; set; }


        public ColumnOverlap(IList<ColumnDefinition> columns, int offsetLeft, int offsetRight)
        {
            Columns = columns;
            OffsetLeft = offsetLeft;
            OffsetRight = offsetRight;
        }
    }
}

