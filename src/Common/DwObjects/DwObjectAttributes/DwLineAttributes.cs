using yuseok.kim.dw2docs.Common.Enums;
using System.Drawing;

namespace yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes
{
    public class DwLineAttributes : DwObjectAttributesBase
    {
        public Point Start { get; set; }
        public Point End { get; set; }
        public ushort LineWidth { get; set; }
        public DwColorWrapper LineColor { get; set; } = new DwColorWrapper();
        public LineStyle LineStyle { get; set; }

        public void SetStart(int x, int y)
        {
            Start = new Point(x, y);
        }

        public void SetEnd(int x, int y)
        {
            End = new Point(x, y);
        }

        public DwLineAttributes() { }

        // override object.Equals
        public override bool Equals(object? obj)
        {
            return base.Equals(obj)
                && obj is DwLineAttributes other
                && Start == other.Start
                && End == other.End
                && LineWidth == other.LineWidth
                && LineColor.Equals(other.LineColor)
                && LineStyle == other.LineStyle;
        }

        // override object.GetHashCode
        public override int GetHashCode()
        {
            return HashCode.Combine(base.GetHashCode(),
                Start,
                End,
                LineWidth,
                LineColor,
                LineStyle);
        }
    }
}
