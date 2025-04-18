using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Attributes;
using yuseok.kim.dw2docs.Xlsx.Extensions;
using yuseok.kim.dw2docs.Xlsx.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;

namespace yuseok.kim.dw2docs.Xlsx.VirtualGridWriter.Renderers
{
    [RendererFor(typeof(DwButtonAttributes), typeof(XlsxButtonRenderer))]
    internal class XlsxButtonRenderer : AbstractXlsxRenderer
    {
        public override ExportedCellBase? Render(ISheet sheet, VirtualCell cell, DwObjectAttributesBase attribute, ICell renderTarget)
        {
            var buttonAttributes = CheckAttributeType<DwButtonAttributes>(attribute);
            
            // Create a cell style
            if (sheet.Workbook is XSSFWorkbook workbook)
            {
                var style = workbook.CreateCellStyle();
                var font = workbook.CreateFont();
                
                // Set up font - only use the FontSize that exists
                if (font is XSSFFont xFont)
                {
                    xFont.FontHeightInPoints = buttonAttributes.FontSize;
                    xFont.FontName = "Arial"; // Default font
                    xFont.IsBold = true; // Make buttons bold by default
                }
                
                style.SetFont(font);
                style.Alignment = HorizontalAlignment.Center;
                style.VerticalAlignment = VerticalAlignment.Center;
                
                // Add fill color (light gray for buttons)
                style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
                style.FillPattern = FillPattern.SolidForeground;
                
                // Add border
                style.BorderTop = BorderStyle.Thin;
                style.BorderBottom = BorderStyle.Thin;
                style.BorderLeft = BorderStyle.Thin;
                style.BorderRight = BorderStyle.Thin;
                
                renderTarget.CellStyle = style;
            }
            
            // Set the button text
            renderTarget.SetCellValue(buttonAttributes.Text);
            
            return new ExportedCell(cell, attribute)
            {
                OutputCell = renderTarget
            };
        }

        public override ExportedCellBase? Render(ISheet sheet, FloatingVirtualCell cell, DwObjectAttributesBase attribute, (int x, int y, XSSFDrawing draw) renderTarget)
        {
            var buttonAttributes = CheckAttributeType<DwButtonAttributes>(attribute);

            XSSFClientAnchor anchor = GetAnchor(
                renderTarget.draw,
                cell,
                renderTarget.x,
                renderTarget.y);

            anchor.AnchorType = AnchorType.MoveDontResize;

            var shape = renderTarget.draw.CreateSimpleShape(anchor);
            shape.FillColor = Color.White.ToRgb();
            shape.LineStyle = LineStyle.Solid;
            shape.LineStyleColor = Color.Gray.ToRgb();
            shape.LineWidth = 1;
            shape.ShapeType = (int)ShapeTypes.RoundedRectangle;
            var paragraph = shape.TextParagraphs[0];
            paragraph.TextAlign = TextAlign.CENTER;
            var run = paragraph.AddNewTextRun();
            run.FontSize = buttonAttributes.FontSize;
            run.Text = buttonAttributes.Text;

            return new ExportedFloatingCell(cell, attribute)
            {
                OutputShape = shape,
            };
        }
    }
}
