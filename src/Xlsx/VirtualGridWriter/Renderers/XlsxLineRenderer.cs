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
    [RendererFor(typeof(DwLineAttributes), typeof(XlsxLineRenderer))]
    internal class XlsxLineRenderer : AbstractXlsxRenderer
    {
        public override ExportedCellBase? Render(ISheet sheet, VirtualCell cell, DwObjectAttributesBase attribute, ICell renderTarget)
        {
            throw new NotImplementedException();
        }

        public override ExportedCellBase? Render(ISheet sheet, FloatingVirtualCell cell, DwObjectAttributesBase attribute, (int x, int y, XSSFDrawing draw) renderTarget)
        {
            var lineAttribute = CheckAttributeType<DwLineAttributes>(attribute);

            XSSFClientAnchor anchor = GetAnchor(
                renderTarget.draw,
                cell,
                renderTarget.x,
                renderTarget.y);

            anchor.AnchorType = AnchorType.MoveDontResize;

            XSSFShape shapeBase;
            if (cell.Object.Height == 0 || cell.Object.Width == 0)
            { // Line is straight 
                var shape = renderTarget.draw.CreateConnector(anchor);
                shapeBase = shape;
                shape.ShapeType = NPOI.OpenXmlFormats.Dml.ST_ShapeType.line;
            }
            else
            {
                var shape = renderTarget.draw.CreateSimpleShape(anchor);
                shapeBase = shape;
                shape.ShapeType = (int)ShapeTypes.Line;
            }

            shapeBase.LineStyle = lineAttribute.LineStyle.DwLineStyleToNpoiLineStyle();
            shapeBase.SetLineStyleColor(lineAttribute.LineColor.Value.ToRgb());
            shapeBase.LineWidth = lineAttribute.LineWidth;

            return new yuseok.kim.dw2docs.Xlsx.Models.ExportedFloatingCell(cell, attribute)
            {
                OutputShape = shapeBase,
            };
        }
    }
}
