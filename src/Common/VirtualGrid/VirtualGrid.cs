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

        /// <summary>
        /// Public accessor for control attributes.
        /// </summary>
        public IReadOnlyDictionary<string, DwObjectAttributesBase> ControlAttributes => _controlAttributes;

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

        /// <summary>
        /// Gets the attributes for a control with the specified name.
        /// </summary>
        /// <param name="controlName">The name of the control.</param>
        /// <returns>The attributes for the control, or null if not found.</returns>
        public DwObjectAttributesBase? GetAttributes(string controlName)
        {
            if (string.IsNullOrEmpty(controlName) || !_controlAttributes.ContainsKey(controlName))
                return null;
                
            return _controlAttributes[controlName];
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
