using yuseok.kim.dw2docs.Common.DwObjects;

namespace yuseok.kim.dw2docs.Common.Utils.Color
{
    public class ColorFactory
    {
        public static DwColorWrapper FromRgb(byte red, byte green, byte blue)
            => new()
            {
                Value = System.Drawing.Color.FromArgb(red, green, blue)
            };

        public static DwColorWrapper FromArgb(byte alpha, byte red, byte green, byte blue)
            => new()
            {
                Value = System.Drawing.Color.FromArgb(alpha, red, green, blue)
            };
    }
}
