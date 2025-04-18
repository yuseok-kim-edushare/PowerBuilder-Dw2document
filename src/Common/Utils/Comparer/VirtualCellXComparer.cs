using yuseok.kim.dw2docs.Common.VirtualGrid;

namespace yuseok.kim.dw2docs.Common.Utils.Comparer
{
    public class VirtualCellXComparer : IComparer<VirtualCell>
    {
        public int Compare(VirtualCell? x, VirtualCell? y)
        {
            if (x == null)
                throw new ArgumentNullException(nameof(x));
            if (y == null)
                throw new ArgumentNullException(nameof(y));
            return x.X.CompareTo(y.X);
        }
    }
}
