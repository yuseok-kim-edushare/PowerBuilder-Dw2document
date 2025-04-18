using yuseok.kim.dw2docs.Common.Enums;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace yuseok.kim.dw2docs.Xlsx.Extensions
{
    public static class HorizontalAlignmentExtensions
    {
        public static TextAlign ToNpoiTextAlignment(this Alignment alignment)
        {
            return alignment switch
            {
                Alignment.Left => TextAlign.LEFT,
                Alignment.Right => TextAlign.RIGHT,
                Alignment.Center => TextAlign.CENTER,
                Alignment.Justify => TextAlign.JUSTIFY,
                _ => throw new ArgumentException($"Unsupported enum value {alignment} in argument {nameof(alignment)}"),
            };
        }

        public static HorizontalAlignment ToNpoiHorizontalAlignment(this Alignment alignment)
        {
            return alignment switch
            {
                Alignment.Left => HorizontalAlignment.Left,
                Alignment.Right => HorizontalAlignment.Right,
                Alignment.Center => HorizontalAlignment.Center,
                Alignment.Justify => HorizontalAlignment.Justify,
                _ => throw new ArgumentException($"Unsupported enum value {alignment} in argument {nameof(alignment)}"),
            };
        }
    }
}
