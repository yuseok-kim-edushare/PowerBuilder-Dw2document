using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;

namespace yuseok.kim.dw2docs.Common.VirtualGridWriter.Models
{
    /// <summary>
    /// Represents an exported floating cell
    /// </summary>
    public class ExportedFloatingCell : ExportedCellBase
    {
        /// <summary>
        /// Gets or sets the output shape object
        /// </summary>
        public object? OutputShape { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExportedFloatingCell"/> class
        /// </summary>
        /// <param name="cell">The floating virtual cell</param>
        /// <param name="attribute">The DW object attributes</param>
        public ExportedFloatingCell(FloatingVirtualCell cell, DwObjectAttributesBase attribute)
            : base(cell, attribute)
        {
        }
    }
} 