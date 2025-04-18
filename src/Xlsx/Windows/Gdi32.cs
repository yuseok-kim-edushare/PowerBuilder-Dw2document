using System.Runtime.InteropServices;

namespace yuseok.kim.dw2docs.Xlsx.Windows
{
    internal static class Gdi32
    {
        [DllImport("gdi32.dll")]
        internal static extern int GetDeviceCaps(IntPtr hdc, int devCap);
    }
}
