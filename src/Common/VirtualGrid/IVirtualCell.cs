using yuseok.kim.dw2docs.Common.DwObjects;

namespace yuseok.kim.dw2docs.Common.VirtualGrid
{
    /// <summary>
    /// Interface for a virtual cell in a virtual grid
    /// </summary>
    public interface IVirtualCell
    {
        /// <summary>
        /// Gets the object contained in this cell
        /// </summary>
        IDwObject Object { get; }
        
        /// <summary>
        /// Gets the owning column, if any
        /// </summary>
        ColumnDefinition? OwningColumn { get; }
        
        /// <summary>
        /// Gets the offset of this cell
        /// </summary>
        CellOffset Offset { get; }
    }
} 