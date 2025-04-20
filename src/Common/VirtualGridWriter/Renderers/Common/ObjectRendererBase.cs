using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;

namespace yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Common
{
    /// <summary>
    /// Base class for rendering DwObjects to a specific output format context.
    /// </summary>
    public abstract class ObjectRendererBase
    {
        /// <summary>
        /// Renders a standard (non-floating) virtual cell.
        /// </summary>
        /// <param name="context">The rendering context (e.g., ISheet for NPOI, Document for DocX).</param>
        /// <param name="cell">The virtual cell definition.</param>
        /// <param name="attribute">The specific attributes of the DW object.</param>
        /// <param name="renderTarget">The target object to render onto (e.g., ICell for NPOI).</param>
        /// <returns>An ExportedCellBase representation, or null if rendering fails.</returns>
        public abstract ExportedCellBase? Render(object context, VirtualCell cell, DwObjectAttributesBase attribute, object renderTarget);

        /// <summary>
        /// Renders a floating virtual cell.
        /// </summary>
        /// <param name="context">The rendering context (e.g., ISheet for NPOI).</param>
        /// <param name="cell">The floating virtual cell definition.</param>
        /// <param name="attribute">The specific attributes of the DW object.</param>
        /// <param name="renderTarget">The target object or coordinates for rendering (e.g., a tuple containing offsets and an NPOI IDrawing).</param>
        /// <returns>An ExportedCellBase representation, or null if rendering fails.</returns>
        public abstract ExportedCellBase? Render(object context, FloatingVirtualCell cell, DwObjectAttributesBase attribute, object renderTarget);

        // Optional: Add methods for specific initialization if needed by certain renderers
        // public virtual void Initialize(object writerContext) { }
    }
} 