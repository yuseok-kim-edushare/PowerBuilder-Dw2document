using yuseok.kim.dw2docs.Common.Enums;
using System.Drawing;

namespace yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes
{
    public class DwShapeAttributes : DwObjectAttributesBase
    {
        public Shape Shape { get; set; }
        public DwColorWrapper FillColor { get; set; } = new DwColorWrapper();
        public FillStyle FillStyle { get; set; }
        public DwColorWrapper OutlineColor { get; set; } = new DwColorWrapper();
        public LineStyle OutlineStyle { get; set; }
        public ushort OutlineWidth { get; set; } = 1;

        public DwShapeAttributes()
        {
            Floating = true;
        }

        // override object.Equals
        public override bool Equals(object? obj)
        {
            return base.Equals(obj)
                && obj is DwShapeAttributes that
                && that.Shape == this.Shape
                && that.FillColor.Equals(this.FillColor)
                && FillStyle == that.FillStyle
                && OutlineColor == that.OutlineColor
                && OutlineStyle == that.OutlineStyle
                && OutlineWidth == that.OutlineWidth;
        }

        // override object.GetHashCode
        public override int GetHashCode()
        {
            int hash = 17;
            hash = hash * 31 + (int)Shape;
            hash = hash * 31 + FillColor.GetHashCode();
            hash = hash * 31 + (int)FillStyle;
            hash = hash * 31 + OutlineColor.GetHashCode();
            hash = hash * 31 + (int)OutlineStyle;
            hash = hash * 31 + OutlineWidth.GetHashCode();
            return hash;
        }
    }
}
