using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;

namespace yuseok.kim.dw2docs.Common.VirtualGridWriter.Models
{
    /// <summary>
    /// Represents an exported non-floating cell
    /// </summary>
    public class ExportedCell : ExportedCellBase
    {
        /// <summary>
        /// Gets or sets the output cell
        /// </summary>
        public object? OutputCell { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExportedCell"/> class
        /// </summary>
        /// <param name="cell">The virtual cell</param>
        /// <param name="attribute">The DW object attributes</param>
        public ExportedCell(VirtualCell cell, DwObjectAttributesBase attribute)
            : base(cell, attribute)
        {
        }
    }
} 