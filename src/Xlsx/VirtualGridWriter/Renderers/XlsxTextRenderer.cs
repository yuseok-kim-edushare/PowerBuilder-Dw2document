using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Attributes;
using yuseok.kim.dw2docs.Xlsx.Extensions;
using yuseok.kim.dw2docs.Xlsx.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SixLabors.ImageSharp.PixelFormats;

namespace yuseok.kim.dw2docs.Xlsx.VirtualGridWriter.Renderers.Xlsx
{
    [RendererFor(typeof(DwTextAttributes), typeof(XlsxTextRenderer))]
    internal class XlsxTextRenderer : AbstractXlsxRenderer
    {
        private const double TextBoxWidthAdjustment = 5;
        private const string LogFilePath = @"C:\temp\Dw2Doc_ExcelError.log";

        private static void LogToFile(string message, Exception? ex = null)
        {
            try
            {
                string logContent = $"[{DateTime.Now}] {message}";
                if (ex != null)
                {
                    logContent += $"\nException Type: {ex.GetType().FullName}\nMessage: {ex.Message}\nStackTrace:\n{ex.StackTrace}";
                }
                logContent += "\n---------------------------------\n";
                File.AppendAllText(LogFilePath, logContent);
            }
            catch (Exception logEx)
            {
                Console.WriteLine($"!!! Failed to write to log file {LogFilePath}: {logEx.Message}");
            }
        }
        private static double ConvertFontSize(DwTextAttributes attributes)
        {
            return attributes.FontSize - 1;
        }

        public override ExportedCellBase? Render(
            ISheet sheet,
            VirtualCell cell,
            DwObjectAttributesBase attribute,
            ICell renderTarget)
        {
            var textAttribute = CheckAttributeType<DwTextAttributes>(attribute);
            
            // Use null-conditional operator and explicitly check for null
            var cellStyle = renderTarget.Row.Sheet.Workbook.CreateCellStyle();
            if (cellStyle is not XSSFCellStyle xssfStyle) 
            {
                throw new InvalidCastException("Could not get XSSFCellStyle");
            }
            
            XSSFCellStyle style = xssfStyle;
            var font = renderTarget.Row.Sheet.Workbook.CreateFont();
            if (font is not XSSFFont xssfFont)
            {
                throw new InvalidCastException("Could not get XSSFFont");
            }
            
            xssfFont.FontHeightInPoints = ConvertFontSize(textAttribute);
            xssfFont.FontName = textAttribute.FontFace;
            xssfFont.SetColor(new XSSFColor(new byte[] {
                textAttribute.FontColor.Value.R,
                textAttribute.FontColor.Value.G,
                textAttribute.FontColor.Value.B}));
            xssfFont.IsBold = textAttribute.FontWeight >= 700;
            xssfFont.IsItalic = textAttribute.Italics;
            xssfFont.IsStrikeout = textAttribute.Strikethrough;
            xssfFont.Underline = textAttribute.Underline ? FontUnderlineType.Single : FontUnderlineType.None;
            style.SetFont(xssfFont);
            style.SetFillForegroundColor(new XSSFColor(new byte[3] {
                textAttribute.BackgroundColor.Value.R,
                textAttribute.BackgroundColor.Value.G,
                textAttribute.BackgroundColor.Value.B,
            }));
            style.FillPattern = textAttribute.BackgroundColor.Value.A == 0 ? FillPattern.NoFill : FillPattern.SolidForeground;
            style.Alignment = textAttribute.Alignment.ToNpoiHorizontalAlignment();
            style.VerticalAlignment = VerticalAlignment.Top;


            // Use null-conditional operator and pattern matching to avoid null warnings
            if (renderTarget is XSSFCell xTarget &&
                textAttribute.FormatString is not null
                && textAttribute.FormatString.ToLower() != "[general]")
            {
                switch (textAttribute.DataType)
                {
                    case Common.Enums.DataType.Money:
                    case Common.Enums.DataType.Number:
                        if (sheet.Workbook?.CreateDataFormat() != null)
                        {
                            style.SetDataFormat(sheet
                                .Workbook
                                .CreateDataFormat()
                                .GetFormat(textAttribute.FormatString ?? string.Empty));
                        }
                        break;
                }
            }

            renderTarget.CellStyle = style;
            if (!string.IsNullOrEmpty(textAttribute.Text) &&
                (textAttribute.DataType is Common.Enums.DataType.Money ||
                textAttribute.DataType is Common.Enums.DataType.Number) &&
                !string.IsNullOrEmpty(textAttribute.RawText))
            {
                double value = double.Parse(textAttribute.RawText);
                LogToFile($"[XlsxTextRenderer] Setting numeric cell value at row {renderTarget.Row.RowNum}, col {renderTarget.ColumnIndex}: {value}");
                renderTarget.SetCellValue(value);
            }
            else
            {
                LogToFile($"[XlsxTextRenderer] Setting text cell value at row {renderTarget.Row.RowNum}, col {renderTarget.ColumnIndex}: '{textAttribute.Text}'");
                renderTarget.SetCellValue(textAttribute.Text);
            }

            return new yuseok.kim.dw2docs.Xlsx.Models.ExportedCell(cell, attribute)
            {
                OutputCell = renderTarget,
            };
        }

        public override ExportedCellBase? Render(ISheet sheet, FloatingVirtualCell cell, DwObjectAttributesBase attribute, (int x, int y, XSSFDrawing draw) renderTarget)
        {
            var textAttribute = CheckAttributeType<DwTextAttributes>(attribute);

            var anchor = GetAnchor(
                renderTarget.draw,
                cell,
                renderTarget.x,
                renderTarget.y,
                widthAdjustment: TextBoxWidthAdjustment);

            anchor.AnchorType = AnchorType.MoveDontResize;

            var textBox = renderTarget.draw.CreateTextbox(anchor);
            textBox.BottomInset = 0;
            textBox.LeftInset = 0;
            textBox.TopInset = 0;
            textBox.RightInset = 0;
            if (textAttribute.BackgroundColor.Value.A != 0)
            {
                textBox.SetFillColor(textAttribute.BackgroundColor.Value.R,
                textAttribute.BackgroundColor.Value.G,
                textAttribute.BackgroundColor.Value.B);
            }


            var paragraph = textBox.TextParagraphs[0];
            var run = paragraph.AddNewTextRun();
            run.FontColor = new Rgb24(
                textAttribute.FontColor.Value.R,
                textAttribute.FontColor.Value.G,
                textAttribute.FontColor.Value.B
            );
            run.Text = textAttribute.Text;
            run.FontSize = ConvertFontSize(textAttribute);
            run.IsItalic = textAttribute.Italics;
            run.IsStrikethrough = textAttribute.Strikethrough;
            run.IsUnderline = textAttribute.Underline;
            run.SetFont(textAttribute.FontFace);
            run.IsBold = textAttribute.FontWeight >= 700;
            paragraph.TextAlign = textAttribute.Alignment.ToNpoiTextAlignment();
            textBox.LineStyle = LineStyle.None;
            //textBox.SetText();

            return new yuseok.kim.dw2docs.Xlsx.Models.ExportedFloatingCell(cell, attribute)
            {
                OutputShape = textBox,
            };
        }
    }
}
