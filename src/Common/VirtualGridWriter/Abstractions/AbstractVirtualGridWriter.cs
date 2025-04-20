using yuseok.kim.dw2docs.Common.DwObjects;
using yuseok.kim.dw2docs.Common.DwObjects.DwObjectAttributes;
using yuseok.kim.dw2docs.Common.Extensions;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using System.Text;

namespace yuseok.kim.dw2docs.Common.VirtualGridWriter.Abstractions
{
    public abstract class AbstractVirtualGridWriter
    {
        protected VirtualGrid.VirtualGrid VirtualGrid { get; }
        private IDictionary<string, DwObjectAttributesBase>? _previousDataSet0;
        protected bool IsClosed { get; set; }

        public IDictionary<string, DwObjectAttributesBase>? PreviousDataSet
        {
            get { return _previousDataSet0; }
            set { _previousDataSet0 = value; }
        }

        private List<BandRows>? _bandsWithoutRelatedHeaders;
        private ISet<string>? _unrepeatableBands;
        private Dictionary<string, bool>? _bandsWithChanges = new();
        private int _currentBandIndex;
        private BandRows? _currentBandRow;

        protected AbstractVirtualGridWriter(VirtualGrid.VirtualGrid virtualGrid)
        {
            _unrepeatableBands = new HashSet<string>(
                virtualGrid.BandRows
                    .Where(br => !br.IsRepeatable)
                    .Select(br => br.Name)
            );

            VirtualGrid = virtualGrid;
            _bandsWithoutRelatedHeaders = VirtualGrid.BandRows
                .Where(br => br.RelatedHeader is null)
                .ToList() ?? throw new InvalidOperationException("Virtual Grid has no Bands configured");
        }

        private int UpdateCurrentBand(int newBandIndex)
        {
            if (newBandIndex < 0)
                _currentBandRow = null;
            else if (_bandsWithoutRelatedHeaders != null)
                _currentBandRow = _bandsWithoutRelatedHeaders[newBandIndex];
            _currentBandIndex = newBandIndex;
            return newBandIndex;
        }

        protected IList<ExportedCellBase>? ProcessNextLine(
            IDictionary<string, DwObjectAttributesBase>? dataSet,
            IList<string> bandsWithChanges)
        {
            // write headers

            // write groups (loop)

            var exportedCells = new List<ExportedCellBase>();

            if (_bandsWithoutRelatedHeaders == null || _currentBandIndex < 0 || _currentBandIndex >= _bandsWithoutRelatedHeaders.Count)
            {
                return exportedCells;
            }

            _currentBandRow = _bandsWithoutRelatedHeaders[_currentBandIndex];
            if (dataSet is null)
            {
                while (_currentBandIndex >= 0)
                {
                    if (_bandsWithoutRelatedHeaders[_currentBandIndex].RelatedTrailers is not null)
                    {
                        foreach (var trailer in _bandsWithoutRelatedHeaders[_currentBandIndex].RelatedTrailers)
                        {
                            var data = PreviousDataSet ?? new Dictionary<string, DwObjectAttributesBase>();
                            var exported = (WriteRows(trailer.Rows, data));

                            if (exported is not null)
                            {
                                exportedCells.AddRange(exported);
                            }
                        }
                        UpdateCurrentBand(_currentBandIndex - 1);
                    }
                }
                return exportedCells;
            }

            if (_currentBandRow != null && bandsWithChanges.Count > 0 && _currentBandRow.Name != bandsWithChanges[0])
            {
                // We need this loop to repeat once more after the condition is met
                bool oneMore = false;
                while (_currentBandRow != null && _currentBandRow.Name != bandsWithChanges[0] || oneMore)
                {   // if dataset changed including bands previous to the current one,
                    // write the trailers of the pending bands 
                    if (_currentBandRow != null && _currentBandRow.RelatedTrailers != null && _currentBandRow.RelatedTrailers.Count > 0)
                    {
                        foreach (var trailer in _currentBandRow.RelatedTrailers)
                        {
                            var data = PreviousDataSet ?? new Dictionary<string, DwObjectAttributesBase>();
                            var _exportedCells = WriteRows(trailer.Rows, data);
                            if (_exportedCells is not null)
                                exportedCells.AddRange(_exportedCells);
                        }
                    }

                    // Guard against incorrectly looping past the condition
                    if (!oneMore)
                        UpdateCurrentBand(_currentBandIndex - 1);
                    if (_currentBandRow != null && _currentBandRow.Name == bandsWithChanges[0])
                        oneMore = !oneMore;
                }
            }

            if (_bandsWithoutRelatedHeaders != null)
            {
                for (; _currentBandIndex < _bandsWithoutRelatedHeaders.Count; ++_currentBandIndex)
                {
                    var _exportedCells = WriteRows(_bandsWithoutRelatedHeaders[_currentBandIndex].Rows, dataSet);

                    if (_exportedCells is not null) exportedCells.AddRange(_exportedCells);
                }
                --_currentBandIndex;
            }

            return exportedCells;
        }

        public IList<ExportedCellBase>? EnterData(IDictionary<string, DwObjectAttributesBase>? dataSet, out string? error)
        {
            error = null;
            try
            {
                IList<ExportedCellBase>? exportedCells;
                if (IsClosed)
                {
                    throw new InvalidOperationException("Writer is already closed");
                }

                if (_bandsWithChanges == null)
                {
                    _bandsWithChanges = new Dictionary<string, bool>();
                }
                else
                {
                    _bandsWithChanges.Clear();
                }

                if (dataSet is not null)
                {
                    // Determine the bands that have changed
                    if (PreviousDataSet is null)
                    {
                        if (VirtualGrid.BandRows != null)
                        {
                            foreach (var band in VirtualGrid.BandRows)
                            {
                                _bandsWithChanges[band.Name] = band.Rows.Count > 0;
                            }
                        }
                    }
                    else
                    {
                        foreach (var (_, cellSet) in VirtualGrid.CellRepository.CellsByY)
                        {
                            foreach (var cell in cellSet)
                            {
                                // Skip null objects or those with empty band names
                                if (cell.Object == null || string.IsNullOrEmpty(cell.Object.Band) || cell.Object.Name == null)
                                {
                                    continue;
                                }

                                // Skip unupdatable bands if _unrepeatableBands collection is available
                                if (_unrepeatableBands != null && _unrepeatableBands.Contains(cell.Object.Band))
                                {
                                    continue;
                                }

                                // Check if the value has changed
                                bool hasChanged = false;
                                if (PreviousDataSet != null && PreviousDataSet.ContainsKey(cell.Object.Name))
                                {
                                    var prevValue = PreviousDataSet[cell.Object.Name];
                                    var currValue = dataSet[cell.Object.Name];
                                    
                                    if (prevValue == null && currValue != null)
                                    {
                                        hasChanged = true;
                                    }
                                    else if (prevValue != null && !prevValue.Equals(currValue))
                                    {
                                        hasChanged = true;
                                    }
                                }
                                else
                                {
                                    // No previous value, consider it changed
                                    hasChanged = true;
                                }

                                if (hasChanged)
                                {
                                    _bandsWithChanges[cell.Object.Band] = true;
                                }
                            }
                        }
                    }
                }

                exportedCells = ProcessNextLine(dataSet, _bandsWithChanges.Keys.ToList());
                PreviousDataSet?.Clear();
                PreviousDataSet = dataSet;

                return exportedCells;
            }
            catch (Exception e)
            {
                error = e.Message;
                return null;
            }
        }

        public void Dispose()
        {
            _previousDataSet0?.Clear();
            _bandsWithChanges?.Clear();
            _bandsWithoutRelatedHeaders?.Clear();
            _unrepeatableBands?.Clear();
            _previousDataSet0 = null;
            _bandsWithoutRelatedHeaders = null;
            _unrepeatableBands = null;
            _bandsWithChanges = null;
        }

        /// <summary>
        /// Process all bands in the virtual grid for rendering
        /// </summary>
        /// <returns>A list of the exported cells</returns>
        protected List<ExportedCellBase> ProcessBands()
        {
            var exportedCells = new List<ExportedCellBase>();
            
            if (_bandsWithoutRelatedHeaders == null || _bandsWithoutRelatedHeaders.Count == 0)
            {
                return exportedCells;
            }
            
            // Reset band index to start from the first band
            UpdateCurrentBand(0);
            
            // Create a dictionary for data - either use previous dataset or create empty
            var data = new Dictionary<string, DwObjectAttributesBase>();
                
            // Copy data from PreviousDataSet if it exists
            if (PreviousDataSet != null)
            {
                foreach (var kvp in PreviousDataSet)
                {
                    data[kvp.Key] = kvp.Value;
                }
            }
            
            // Process each band in order
            foreach (var band in _bandsWithoutRelatedHeaders)
            {
                var result = WriteRows(band.Rows, data);
                
                if (result != null)
                {
                    exportedCells.AddRange(result);
                }
                
                // Process any related trailers for this band
                if (band.RelatedTrailers != null && band.RelatedTrailers.Count > 0)
                {
                    foreach (var trailer in band.RelatedTrailers)
                    {
                        var trailerResult = WriteRows(trailer.Rows, data);
                        if (trailerResult != null)
                        {
                            exportedCells.AddRange(trailerResult);
                        }
                    }
                }
            }
            
            return exportedCells;
        }

        protected abstract IList<ExportedCellBase>? WriteRows(IList<RowDefinition> rows, IDictionary<string, DwObjectAttributesBase> data);

        public abstract bool Write(string? sheetname, out string? error);

        public bool WriteEntireGrid(string? sheetname, out string? error)
        {
            error = null;
            
            try
            {
                // Get the control attributes directly using the public property
                var controlAttributes = VirtualGrid.ControlAttributes;
                
                if (controlAttributes == null || controlAttributes.Count == 0)
                {
                    error = "No control attributes found in VirtualGrid";
                    return false;
                }
                
                // Convert IReadOnlyDictionary to IDictionary by creating a new Dictionary
                var writableAttributes = controlAttributes.ToDictionary(x => x.Key, x => x.Value);
                
                // Process all rows using the control attributes
                foreach (var band in VirtualGrid.BandRows)
                {
                    if (band.Rows.Count > 0)
                    {
                        WriteRows(band.Rows, writableAttributes);
                    }
                }
                
                // Delegate to the concrete implementation to finish writing
                return Write(sheetname, out error);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }
    }
}
