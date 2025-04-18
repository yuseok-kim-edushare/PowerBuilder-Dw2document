using System.Runtime.InteropServices;

namespace yuseok.kim.dw2docs.Xlsx.Windows
{
    internal static class User32
    {
        [DllImport("user32.dll")]
        internal static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);
    }
}
