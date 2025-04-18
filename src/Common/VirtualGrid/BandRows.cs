using yuseok.kim.dw2docs.Common.Enums;

namespace yuseok.kim.dw2docs.Common.VirtualGrid
{
    public class BandRows
    {
        public string Name { get; }
        public IList<RowDefinition> Rows { get; }
        public BandRows? ParentBand { get; set; }
        public BandRows? RelatedHeader { get; set; }
        public IList<BandRows> RelatedTrailers { get; }
        public bool IsRepeatable { get; set; }
        public BandType BandType { get; set; }


        public BandRows(string name, IList<RowDefinition> rows)
        {
            RelatedTrailers = new List<BandRows>();
            Name = name;
            Rows = rows;
        }
    }
}
