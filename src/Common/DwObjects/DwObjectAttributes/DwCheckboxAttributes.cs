namespace yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;

public class DwCheckboxAttributes : DwTextAttributes
{
    public string? Label { get; set; }
    public string? CheckedValue { get; set; }
    public string? UncheckedValue { get; set; }
    public bool LeftText { get; set; }

    // override object.Equals
    public override bool Equals(object? obj)
    {
        return base.Equals(obj)
            && obj is DwCheckboxAttributes that
            && that.Label == Label
            && that.CheckedValue == CheckedValue
            && that.UncheckedValue == UncheckedValue
            && that.LeftText == LeftText;
    }

    // override object.GetHashCode
    public override int GetHashCode()
    {
        int hash = base.GetHashCode();
        hash = hash * 31 + (Label != null ? Label.GetHashCode() : 0);
        hash = hash * 31 + (CheckedValue != null ? CheckedValue.GetHashCode() : 0);
        hash = hash * 31 + (UncheckedValue != null ? UncheckedValue.GetHashCode() : 0);
        hash = hash * 31 + LeftText.GetHashCode();
        return hash;
    }
}
