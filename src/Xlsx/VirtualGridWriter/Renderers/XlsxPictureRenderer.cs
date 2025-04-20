using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Attributes;
using yuseok.kim.dw2docs.Xlsx.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;

namespace yuseok.kim.dw2docs.Xlsx.VirtualGridWriter.Renderers
{
    [RendererFor(typeof(DwPictureAttributes), typeof(XlsxPictureRenderer))]
    internal class XlsxPictureRenderer : AbstractXlsxRenderer
    {
        public IDictionary<string, int> PictureCache { get; private set; } = new Dictionary<string, int>();
        public byte[]? PictureData { get; private set; }

        public void SetPictureCache(IDictionary<string, int> pictureCache)
        {
            PictureCache = pictureCache;
        }

        public override ExportedCellBase? Render(ISheet sheet, VirtualCell cell, DwObjectAttributesBase attribute, ICell renderTarget)
        {
            throw new NotImplementedException();
        }

        public override ExportedCellBase? Render(ISheet sheet, FloatingVirtualCell cell, DwObjectAttributesBase attribute, (int x, int y, XSSFDrawing draw) renderTarget)
        {
            var pictureAttribute = CheckAttributeType<DwPictureAttributes>(attribute);

            if (pictureAttribute.FileName is null)
                return null;

            int pictureIndex = -1;
            XSSFShape? outputShape = null;

            XSSFClientAnchor anchor = GetAnchor(
                renderTarget.draw,
                cell,
                renderTarget.x,
                renderTarget.y);

            if (PictureCache.ContainsKey(pictureAttribute.FileName))
            {
                pictureIndex = PictureCache[pictureAttribute.FileName];
            }
            else
            {
                if (!File.Exists(pictureAttribute.FileName))
                {
                    var textBox = renderTarget.draw.CreateTextbox(anchor);
                    textBox.BottomInset = 0;
                    textBox.LeftInset = 0;
                    textBox.TopInset = 0;
                    textBox.RightInset = 0;
                    var paragraph = textBox.TextParagraphs[0];
                    var run = paragraph.AddNewTextRun();
                    run.Text = $"{pictureAttribute.FileName} not found";
                    textBox.LineStyle = LineStyle.Solid;

                    outputShape = textBox;
                }
                else
                {
                    var image = Image.FromFile(pictureAttribute.FileName);
                    using var memStream = new MemoryStream();
                    image.Save(memStream, image.RawFormat);

                    PictureType t;
                    if (image.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Bmp))
                    {
                        t = PictureType.BMP;
                    }
                    else if (image.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Jpeg))
                    {
                        t = PictureType.JPEG;
                    }
                    else if (image.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Png))
                    {
                        t = PictureType.PNG;
                    }
                    else if (image.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Gif))
                    {
                        t = PictureType.GIF;
                    }
                    else if (image.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Tiff))
                    {
                        t = PictureType.TIFF;
                    }
                    else
                    {
                        throw new NotSupportedException("Unsupported image format");
                    }

                    PictureData = memStream.GetBuffer();
                    PictureCache[pictureAttribute.FileName] = pictureIndex = sheet.Workbook.AddPicture(PictureData, t);
                }
            }



            anchor.AnchorType = AnchorType.MoveDontResize;

            if (pictureIndex != -1)
            {
                outputShape = (XSSFPicture)renderTarget.draw.CreatePicture(anchor, pictureIndex);
            }
            return new yuseok.kim.dw2docs.Xlsx.Models.ExportedFloatingCell(cell, attribute)
            {
                OutputShape = outputShape,
            };
        }
    }
}
