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
            return HashCode.Combine(base.GetHashCode(),
                FileName,
                Transparency);
        }
    }
}
