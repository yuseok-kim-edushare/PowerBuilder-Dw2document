﻿namespace yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes
{
    public class DwButtonAttributes : DwObjectAttributesBase
    {
        public string Text { get; set; } = string.Empty;
        public int FontSize { get; set; } = 9;

        // override object.Equals
        public override bool Equals(object? obj)
        {
            return base.Equals(obj)
                && obj is DwButtonAttributes that
                && that.Text == Text
                && that.FontSize == FontSize;
        }

        // override object.GetHashCode
        public override int GetHashCode()
        {
            int hash = base.GetHashCode();
            hash = hash * 31 + (Text != null ? Text.GetHashCode() : 0);
            hash = hash * 31 + FontSize.GetHashCode();
            return hash;
        }
    }
}
