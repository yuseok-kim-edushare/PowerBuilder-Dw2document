namespace yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes
{
    public class DwPictureAttributes : DwObjectAttributesBase
    {
        public string? FileName { get; set; }
        public byte Transparency { get; set; }

        public override bool Equals(object? obj)
        {
            return base.Equals(obj)
                && obj is DwPictureAttributes that
                && that.FileName == FileName
                && Transparency == that.Transparency;
        }

        public override int GetHashCode()
        {
            int hash = base.GetHashCode();
            hash = hash * 31 + (FileName != null ? FileName.GetHashCode() : 0);
            hash = hash * 31 + Transparency.GetHashCode();
            return hash;
        }
    }
}
