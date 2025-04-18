using yuseok.kim.dw2docs.Xlsx.Windows;
using static yuseok.kim.dw2docs.Xlsx.Windows.ScreenTools;

namespace yuseok.kim.dw2docs.Xlsx.Extensions
{
    internal static class DimensionExtensions
    {
        public static int TwipsToPixels(this int twips, MeasureDirection direction)
        {
            return (int)(twips * ((double)GetPPI(direction)) / 1440.0);
        }

        public static int PixelsToTwips(this int pixels, MeasureDirection direction)
        {
            return (int)(pixels * 1440.0 / GetPPI(direction));
        }
    }
}
