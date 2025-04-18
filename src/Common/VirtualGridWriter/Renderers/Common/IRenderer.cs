using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;

namespace yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Common
{
    public interface IRenderer<TContext, TCell, TAttrib, TTarget>
        where TCell : VirtualCell
    {
        /// <summary>
        /// Render the object
        /// </summary>
        ExportedCellBase? Render(TContext context, TCell cell, TAttrib attribute, TTarget renderTarget);
    }
}
