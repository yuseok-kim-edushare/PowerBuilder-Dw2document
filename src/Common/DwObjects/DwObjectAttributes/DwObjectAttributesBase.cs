namespace yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes
{
    public abstract class DwObjectAttributesBase
    {
        public bool IsVisible { get; set; }
        public bool Floating { get; protected set; }

        public DwObjectAttributesBase()
        {
        }

        // override object.Equals
        public override bool Equals(object? obj)
        {
            return obj is DwObjectAttributesBase other
                && IsVisible == other.IsVisible
                && Floating == other.Floating;
        }

        // override object.GetHashCode
        public override int GetHashCode()
        {
            return HashCode.Combine(IsVisible,
                Floating);
        }
    }
}
