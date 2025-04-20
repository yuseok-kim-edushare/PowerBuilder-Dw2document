namespace yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;

public class DwRadioButtonAttributes : DwTextAttributes
{
    public IDictionary<string, string>? CodeTable { get; set; }
    public short Columns { get; set; }
    public bool LeftText { get; set; }

    // override object.Equals
    public override bool Equals(object? obj)
    {
        return obj is DwRadioButtonAttributes other
            && Columns == other.Columns
            && LeftText == other.LeftText
            && ((CodeTable is null) == (other.CodeTable is null))
            && (CodeTable is null || CodeTable.Keys.SequenceEqual(other.CodeTable!.Keys))
            && (CodeTable is null || CodeTable.Values.SequenceEqual(other.CodeTable!.Values));
    }

    // override object.GetHashCode
    public override int GetHashCode()
    {
        int hash = 17;
        hash = hash * 31 + Columns.GetHashCode();
        hash = hash * 31 + LeftText.GetHashCode();
        
        if (CodeTable != null)
        {
            foreach (var key in CodeTable.Keys)
            {
                hash = hash * 31 + (key != null ? key.GetHashCode() : 0);
            }
            
            foreach (var value in CodeTable.Values)
            {
                hash = hash * 31 + (value != null ? value.GetHashCode() : 0);
            }
        }
        
        return hash;
    }
}
