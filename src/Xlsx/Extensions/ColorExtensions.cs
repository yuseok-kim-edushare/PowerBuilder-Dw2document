using System.Drawing;

namespace yuseok.kim.dw2docs.Xlsx.Extensions
{
    public static class ColorExtensions
    {
        public static int ToRgb(this Color color)
        {
            return color.ToArgb() & 16777215; /// this mask removes the highest 8 bits
        }

        public static Color ToColor(this byte[] bytes)
        {
            if (bytes.Length != 3)
                throw new ArgumentException("Argument is not a valid color description", nameof(bytes));

            return Color.FromArgb(bytes[0] << 16 | bytes[1] << 8 | bytes[2] | 0xff << 24);
        }
    }
}
