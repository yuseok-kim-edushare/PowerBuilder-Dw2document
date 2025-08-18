using yuseok.kim.dw2docs.Common.Enums;
using System.Drawing;

namespace yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes
{
    /// <summary>
    /// Attributes for PowerBuilder DataWindow compute objects
    /// </summary>
    public class DwComputeAttributes : DwObjectAttributesBase
    {
        public DataType DataType { get; set; }
        public string? Expression { get; set; }
        public string? FormatString { get; set; }
        public string? Text { get; set; }
        public Alignment Alignment { get; set; } = Alignment.Left;
        public byte FontSize { get; set; }
        public short FontWeight { get; set; }
        public bool Underline { get; set; }
        public bool Italics { get; set; }
        public bool Strikethrough { get; set; }
        public string? FontFace { get; set; }
        public DwColorWrapper FontColor { get; set; } = new DwColorWrapper()
        {
            Value = Color.FromArgb(255, 0, 0, 0)
        };
        public DwColorWrapper BackgroundColor { get; set; } = new DwColorWrapper()
        {
            Value = Color.FromArgb(255, 255, 255, 255)
        };

        // Position and size attributes
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string? Band { get; set; }

        public DwComputeAttributes() { }

        // override object.Equals
        public override bool Equals(object? obj)
        {
            return base.Equals(obj)
                && obj is DwComputeAttributes other
                && Expression == other.Expression
                && Text == other.Text
                && FontFace == other.FontFace
                && FontSize == other.FontSize
                && FontWeight == other.FontWeight
                && Italics == other.Italics
                && Underline == other.Underline
                && Strikethrough == other.Strikethrough
                && FontColor.Equals(other.FontColor)
                && BackgroundColor.Equals(other.BackgroundColor)
                && Alignment == other.Alignment
                && FormatString == other.FormatString
                && DataType == other.DataType
                && X == other.X
                && Y == other.Y
                && Width == other.Width
                && Height == other.Height
                && Band == other.Band;
        }

        // override object.GetHashCode
        public override int GetHashCode()
        {
            var hash = base.GetHashCode();
            hash = (hash * 17) ^ (Expression?.GetHashCode() ?? 0);
            hash = (hash * 17) ^ (Text?.GetHashCode() ?? 0);
            hash = (hash * 17) ^ (FontFace?.GetHashCode() ?? 0);
            hash = (hash * 17) ^ FontSize.GetHashCode();
            hash = (hash * 17) ^ FontWeight.GetHashCode();
            hash = (hash * 17) ^ Italics.GetHashCode();
            hash = (hash * 17) ^ Underline.GetHashCode();
            hash = (hash * 17) ^ Strikethrough.GetHashCode();
            hash = (hash * 17) ^ FontColor.GetHashCode();
            hash = (hash * 17) ^ BackgroundColor.GetHashCode();
            hash = (hash * 17) ^ Alignment.GetHashCode();
            hash = (hash * 17) ^ (FormatString?.GetHashCode() ?? 0);
            hash = (hash * 17) ^ DataType.GetHashCode();
            hash = (hash * 17) ^ X.GetHashCode();
            hash = (hash * 17) ^ Y.GetHashCode();
            hash = (hash * 17) ^ Width.GetHashCode();
            hash = (hash * 17) ^ Height.GetHashCode();
            hash = (hash * 17) ^ (Band?.GetHashCode() ?? 0);
            return hash;
        }

        public override string ToString()
        {
            return $"{Expression} ({Text})";
        }
    }
}