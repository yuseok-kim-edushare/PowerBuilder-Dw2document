using yuseok.kim.dw2docs.Common.Enums;
using yuseok.kim.dw2docs.Common.Extensions;
using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using System.Text;

namespace yuseok.kim.dw2docs.Common.VirtualGrid
{
    public class VirtualGrid
    {
        public IList<RowDefinition> Rows { get; }
        public IList<ColumnDefinition> Columns { get; }
        public IList<BandRows> BandRows { get; }
        internal VirtualCellRepository CellRepository { get; }
        public DwType DwType { get; set; }
        
        // Missing field that TestVirtualGrid.SetAttributes tries to access
        private Dictionary<string, DwObjectAttributesBase> _controlAttributes;

        internal VirtualGrid(
            IList<RowDefinition> rows,
            IList<ColumnDefinition> columns,
            IList<BandRows> rowsPerBand,
            VirtualCellRepository cellRepository,
            DwType DwType
            )
        {
            Rows = rows;
            Columns = columns;
            BandRows = rowsPerBand;
            CellRepository = cellRepository;
            this.DwType = DwType;
            _controlAttributes = new Dictionary<string, DwObjectAttributesBase>();
        }

        // Add a method to get the control attributes for renderers to use
        internal Dictionary<string, DwObjectAttributesBase> GetControlAttributes()
        {
            return _controlAttributes;
        }

        // Add a setter method for reflection to use
        internal void SetAttributes(Dictionary<string, DwObjectAttributesBase> attributes)
        {
            if (attributes == null)
                return;
                
            foreach (var attr in attributes)
            {
                _controlAttributes[attr.Key] = attr.Value;
            }
        }

        public override string ToString()
        {
            var builder = new StringBuilder();
            builder.AppendLine("Rows ->");
            builder.AppendLine(Rows?.CollectionToString() ?? "null");

            builder.AppendLine("Columns -> ");
            builder.AppendLine(Columns?.CollectionToString() ?? "null");

            builder.AppendLine("Rows/Band-> ");
            foreach (var bandRow in BandRows)
            {
                builder.AppendLine($"--- {bandRow.Name} ---");
                builder.AppendLine(bandRow.Rows.CollectionToString());
            }

            return builder.ToString();
        }
    }
}
