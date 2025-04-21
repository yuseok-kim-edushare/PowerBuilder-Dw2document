using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Attributes;
using yuseok.kim.dw2docs.Xlsx.Extensions;
using yuseok.kim.dw2docs.Xlsx.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace yuseok.kim.dw2docs.Xlsx.VirtualGridWriter.Renderers
{
    [RendererFor(typeof(DwShapeAttributes), typeof(XlsxShapeRenderer))]
    internal class XlsxShapeRenderer : AbstractXlsxRenderer
    {
        public override ExportedCellBase? Render(ISheet sheet, VirtualCell cell, DwObjectAttributesBase attribute, ICell renderTarget)
        {
            throw new NotImplementedException();
        }

        public override ExportedCellBase? Render(ISheet sheet, FloatingVirtualCell cell, DwObjectAttributesBase attribute, (int x, int y, XSSFDrawing draw) renderTarget)
        {
            var shapeAttribute = CheckAttributeType<DwShapeAttributes>(attribute);

            XSSFClientAnchor anchor = GetAnchor(
                renderTarget.draw,
                cell,
                renderTarget.x,
                renderTarget.y);

            anchor.AnchorType = AnchorType.MoveDontResize;

            var shape = renderTarget.draw.CreateSimpleShape(anchor);
            shape.FillColor = shapeAttribute.FillColor.Value.ToRgb();
            if (shapeAttribute.FillStyle == Common.Enums.FillStyle.Transparent)
            {
                shape.IsNoFill = true;
            }
            shape.LineStyle = shapeAttribute.OutlineStyle.DwLineStyleToNpoiLineStyle();
            shape.SetLineStyleColor(shapeAttribute.OutlineColor.Value.ToRgb());
            shape.LineWidth = shapeAttribute.OutlineWidth;
            shape.ShapeType = (int)(shapeAttribute.Shape switch
            {
                Common.Enums.Shape.Circle => ShapeTypes.Ellipse,
                Common.Enums.Shape.Rectangle => ShapeTypes.Rectangle,
                _ => throw new NotImplementedException("Unsupported shape type")
            });
            //textBox.SetText();

            return new yuseok.kim.dw2docs.Xlsx.Models.ExportedFloatingCell(cell, attribute)
            {
                OutputShape = shape,
            };
        }
    }
}
