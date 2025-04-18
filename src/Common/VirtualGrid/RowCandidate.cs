namespace yuseok.kim.dw2docs.Common.VirtualGrid;

internal class RowCandidate
{
    private int offset;
    private int height;

    public int Offset
    {
        get => offset;

        set
        {
            offset = value;
            UpdateChildren();
        }
    }
    public int Height
    {
        get => height; set
        {
            height = value;
            UpdateChildren();
        }
    }

    public bool IsFiller { get; set; }
    public int Bound => Offset + Height;
    public string? Band { get; set; }

    public IList<VirtualCell> Objects { get; set; }

    public RowCandidate()
    {
        Objects = new List<VirtualCell>();
    }

    public RowCandidate(IList<VirtualCell> objects)
    {
        Objects = objects;
    }

    private void UpdateChildren()
    {
        foreach (var child in Objects)
        {
            child.Y = Offset;
            child.Height = Height;
        }
    }

    public override string ToString()
    {
        return $"{Offset}+{Height}|{Bound} -> {Objects.Count} objs ";
    }
}
