namespace yuseok.kim.dw2docs.Common.VirtualGrid
{
    /// <summary>
    /// Represents the offset information for a cell in the virtual grid
    /// </summary>
    public class CellOffset
    {
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
    }
    
    /// <summary>
    /// Represents a point in 2D space
    /// </summary>
    public class Point
    {
        /// <summary>
        /// Gets or sets the X coordinate
        /// </summary>
        public int X { get; set; }
        
        /// <summary>
        /// Gets or sets the Y coordinate
        /// </summary>
        public int Y { get; set; }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="Point"/> class
        /// </summary>
        public Point()
        {
            X = 0;
            Y = 0;
        }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="Point"/> class with the specified coordinates
        /// </summary>
        /// <param name="x">The X coordinate</param>
        /// <param name="y">The Y coordinate</param>
        public Point(int x, int y)
        {
            X = x;
            Y = y;
        }
    }
} 