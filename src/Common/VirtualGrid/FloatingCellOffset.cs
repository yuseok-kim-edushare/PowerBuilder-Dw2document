using System.Drawing;

namespace yuseok.kim.dw2docs.Common.VirtualGrid
{
    /// <summary>
    /// Represents the offset information for a floating cell in the virtual grid
    /// </summary>
    public class FloatingCellOffset
    {
        /// <summary>
        /// Gets or sets the start column
        /// </summary>
        public ColumnDefinition? StartColumn { get; set; }
        
        /// <summary>
        /// Gets or sets the start row
        /// </summary>
        public RowDefinition? StartRow { get; set; }
        
        /// <summary>
        /// Gets or sets the start offset from the top-left corner
        /// </summary>
        public Point StartOffset { get; set; } = new Point();
        
        /// <summary>
        /// Gets or sets the end offset from the bottom-right corner
        /// </summary>
        public Point EndOffset { get; set; } = new Point();
        
        /// <summary>
        /// Gets or sets the row span of this cell
        /// </summary>
        public int RowSpan { get; set; } = 1;
        
        /// <summary>
        /// Gets or sets the column span of this cell
        /// </summary>
        public int ColSpan { get; set; } = 1;
        
        /// <summary>
        /// Initializes a new instance of the <see cref="FloatingCellOffset"/> class
        /// </summary>
        public FloatingCellOffset()
        {
        }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="FloatingCellOffset"/> class with the specified column and row
        /// </summary>
        /// <param name="startCol">The start column</param>
        /// <param name="startRow">The start row</param>
        public FloatingCellOffset(ColumnDefinition startCol, RowDefinition startRow)
        {
            StartColumn = startCol;
            StartRow = startRow;
        }
    }
}
