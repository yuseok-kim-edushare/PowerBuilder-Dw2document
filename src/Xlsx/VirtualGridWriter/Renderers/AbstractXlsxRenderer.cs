using yuseok.kim.dw2docs.Common.DwObjects;
using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Common;
using yuseok.kim.dw2docs.Xlsx.Helpers;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace yuseok.kim.dw2docs.Xlsx.VirtualGridWriter.Renderers
{
    internal abstract class AbstractXlsxRenderer : ObjectRendererBase
    {
        public abstract ExportedCellBase? Render(ISheet sheet, VirtualCell cell, DwObjectAttributesBase attribute, ICell renderTarget);
        public abstract ExportedCellBase? Render(ISheet sheet, FloatingVirtualCell cell, DwObjectAttributesBase attribute, (int x, int y, XSSFDrawing draw) renderTarget);

        public override ExportedCellBase? Render(object context, VirtualCell cell, DwObjectAttributesBase attribute, object renderTarget)
        {
            if (context is ISheet sheet && renderTarget is ICell targetCell)
            {
                return Render(sheet, cell, attribute, targetCell);
            }
            Console.WriteLine($"Error: Invalid context or renderTarget type for non-floating render in {this.GetType().Name}");
            return null;
        }

        public override ExportedCellBase? Render(object context, FloatingVirtualCell cell, DwObjectAttributesBase attribute, object renderTarget)
        {
            if (context is ISheet sheet && renderTarget is ValueTuple<int, int, XSSFDrawing> tuple)
            {
                return Render(sheet, cell, attribute, (tuple.Item1, tuple.Item2, tuple.Item3));
            }
            Console.WriteLine($"Error: Invalid context or renderTarget type for floating render in {this.GetType().Name}");
            return null;
        }

        protected static T CheckAttributeType<T>(DwObjectAttributesBase attribute)
            where T : DwObjectAttributesBase
        {
            var instanceType = typeof(T);
            if (attribute.GetType() != instanceType)
                throw new ArgumentException($"Attribute {nameof(attribute)} is not of type {instanceType.Name}");

            return (T)attribute;
        }

        protected static XSSFClientAnchor GetAnchor(XSSFDrawing drawing,
                                                    FloatingVirtualCell floatingCell,
                                                    int startColumn,
                                                    int startRow,
                                                    double widthAdjustment = 0.0,
                                                    double heightAdjustment = 0.0) =>
            (drawing.CreateAnchor(
                UnitConversion.PixelsToEMU((floatingCell.Offset.StartOffset.X)),
                UnitConversion.PixelsToEMU(floatingCell.Offset.StartOffset.Y),
                UnitConversion.PixelsToEMU((int)(-floatingCell.Offset.EndOffset.X + widthAdjustment)),
                UnitConversion.PixelsToEMU((int)(-floatingCell.Offset.EndOffset.Y + heightAdjustment)),
                startColumn,
                startRow,
                startColumn + floatingCell.Offset.ColSpan,
                startRow + floatingCell.Offset.RowSpan) as XSSFClientAnchor)!;

    }
}
