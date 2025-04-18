using yuseok.kim.dw2docs.Common.Enums;
using System.Drawing;

namespace yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes
{
    public class DwTextAttributes : DwObjectAttributesBase
    {
        public DataType DataType { get; set; }
        public string? FormatString { get; set; }
        public string? RawText { get; set; }
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
            Value = Color.FromArgb(0, 255, 255, 255)
        };

        public DwTextAttributes()
        {
            Floating = false;
        }

        // override object.Equals
        public override bool Equals(object? obj)
        {
            return base.Equals(obj)
                && obj is DwTextAttributes that
                && DataType == that.DataType
                && FormatString == that.FormatString
                && RawText == that.RawText
                && Text == that.Text
                && Alignment == that.Alignment
                && FontSize == that.FontSize
                && FontWeight == that.FontWeight
                && Underline == that.Underline
                && Italics == that.Italics
                && Strikethrough == that.Strikethrough
                && FontFace == that.FontFace
                && FontColor.Equals(that.FontColor)
                && BackgroundColor.Equals(that.BackgroundColor);
        }

        // override object.GetHashCode
        public override int GetHashCode()
        {
            return HashCode.Combine(base.GetHashCode()
                , DataType
                , FormatString
                , RawText
                , Text
                , Alignment
                , FontSize
                , HashCode.Combine(
                    FontWeight
                    , Underline
                    , Italics
                    , Strikethrough
                    , FontFace
                    , FontColor
                    , BackgroundColor
                    )
                );
        }

        public override string ToString()
        {
            return $"{Text}";
        }
    }
}
