using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using yuseok.kim.dw2docs.Common.DwObjects;

namespace yuseok.kim.dw2docs.Common.Positioning
{
    /// <summary>
    /// Interface for document writers that support streaming and batching
    /// </summary>
    public interface IDocumentWriter : IDisposable
    {
        Task WriteHeaderAsync(DocumentHeader header);
        Task WriteRowBatchAsync(IEnumerable<DocumentRow> rows);
        Task WriteFooterAsync(DocumentFooter footer);
        Task FlushAsync();
    }

    /// <summary>
    /// Document header information
    /// </summary>
    public class DocumentHeader
    {
        public string Title { get; set; } = "";
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public string Author { get; set; } = "";
        public List<ColumnDefinition> Columns { get; set; } = new();
    }

    /// <summary>
    /// Document footer information
    /// </summary>
    public class DocumentFooter
    {
        public int TotalRows { get; set; }
        public DateTime CompletedDate { get; set; } = DateTime.Now;
        public string Summary { get; set; } = "";
    }

    /// <summary>
    /// Column definition for document structure
    /// </summary>
    public class ColumnDefinition
    {
        public string Name { get; set; } = "";
        public string DataType { get; set; } = "";
        public int Width { get; set; }
        public string Format { get; set; } = "";
    }

    /// <summary>
    /// Document row containing cell data
    /// </summary>
    public class DocumentRow
    {
        public int RowIndex { get; set; }
        public List<DocumentCell> Cells { get; set; } = new();
        public string BandName { get; set; } = "";
    }

    /// <summary>
    /// Individual cell in a document row
    /// </summary>
    public class DocumentCell
    {
        public object? Value { get; set; }
        public string FormattedValue { get; set; } = "";
        public DocumentFont? Font { get; set; }
        public DocumentColor? BackgroundColor { get; set; }
        public DocumentColor? ForegroundColor { get; set; }
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }

    /// <summary>
    /// Performance monitoring and memory usage tracking
    /// </summary>
    public class PerformanceMonitor
    {
        private long _initialMemory;
        private DateTime _startTime;
        private int _processedRows;

        public void Start()
        {
            _initialMemory = GC.GetTotalMemory(true);
            _startTime = DateTime.Now;
            _processedRows = 0;
        }

        public void RecordRowProcessed()
        {
            Interlocked.Increment(ref _processedRows);
        }

        public PerformanceStats GetStats()
        {
            var currentMemory = GC.GetTotalMemory(false);
            var elapsed = DateTime.Now - _startTime;

            return new PerformanceStats
            {
                ElapsedTime = elapsed,
                ProcessedRows = _processedRows,
                MemoryUsed = currentMemory - _initialMemory,
                RowsPerSecond = elapsed.TotalSeconds > 0 ? _processedRows / elapsed.TotalSeconds : 0
            };
        }
    }

    /// <summary>
    /// Performance statistics
    /// </summary>
    public class PerformanceStats
    {
        public TimeSpan ElapsedTime { get; set; }
        public int ProcessedRows { get; set; }
        public long MemoryUsed { get; set; }
        public double RowsPerSecond { get; set; }

        public override string ToString()
        {
            return $"Processed {ProcessedRows} rows in {ElapsedTime:hh\\:mm\\:ss}, " +
                   $"Memory: {MemoryUsed / 1024 / 1024:F2} MB, " +
                   $"Rate: {RowsPerSecond:F1} rows/sec";
        }
    }

    /// <summary>
    /// High-performance renderer optimized for large DataWindows
    /// Supports streaming, batching, memory management, and progress reporting
    /// </summary>
    public class PerformanceOptimizedRenderer
    {
        private const int DefaultBatchSize = 100;
        private const long MaxMemoryThreshold = 500 * 1024 * 1024; // 500MB

        private readonly CoordinateMapper _coordinateMapper;
        private readonly StyleMapper _styleMapper;
        private readonly PerformanceMonitor _performanceMonitor;

        public PerformanceOptimizedRenderer()
        {
            _coordinateMapper = new CoordinateMapper();
            _styleMapper = new StyleMapper();
            _performanceMonitor = new PerformanceMonitor();
        }

        /// <summary>
        /// Renders a large DataWindow with streaming and progress reporting
        /// </summary>
        /// <param name="dataWindow">PowerBuilder DataWindow</param>
        /// <param name="outputStream">Output stream</param>
        /// <param name="progress">Progress reporter (0-100)</param>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <param name="batchSize">Number of rows to process in each batch</param>
        /// <returns>Performance statistics</returns>
        public async Task<PerformanceStats> RenderDocumentAsync(
            IPowerBuilderDataWindow dataWindow,
            Stream outputStream,
            IProgress<int>? progress = null,
            CancellationToken cancellationToken = default,
            int batchSize = DefaultBatchSize)
        {
            if (dataWindow == null) throw new ArgumentNullException(nameof(dataWindow));
            if (outputStream == null) throw new ArgumentNullException(nameof(outputStream));

            _performanceMonitor.Start();

            try
            {
                int totalRows = dataWindow.GetRowCount();
                int processedRows = 0;

                using var documentWriter = CreateDocumentWriter(outputStream);

                // Process header once
                await ProcessHeaderAsync(dataWindow, documentWriter);

                // Process rows in batches
                for (int startRow = 1; startRow <= totalRows; startRow += batchSize)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // Check memory usage and force garbage collection if needed
                    await CheckMemoryUsageAsync();

                    int endRow = Math.Min(startRow + batchSize - 1, totalRows);
                    
                    await ProcessRowBatchAsync(dataWindow, documentWriter, startRow, endRow);

                    processedRows += (endRow - startRow + 1);
                    
                    // Update progress
                    if (progress != null && totalRows > 0)
                    {
                        int percentComplete = (int)((double)processedRows / totalRows * 100);
                        progress.Report(percentComplete);
                    }

                    // Record performance metrics
                    for (int i = startRow; i <= endRow; i++)
                    {
                        _performanceMonitor.RecordRowProcessed();
                    }
                }

                // Process footer once
                await ProcessFooterAsync(dataWindow, documentWriter, totalRows);

                return _performanceMonitor.GetStats();
            }
            catch (OperationCanceledException)
            {
                // Handle cancellation gracefully
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Document rendering failed: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Processes header section with band information
        /// </summary>
        private async Task ProcessHeaderAsync(IPowerBuilderDataWindow dataWindow, IDocumentWriter writer)
        {
            var header = new DocumentHeader
            {
                Title = "DataWindow Export",
                CreatedDate = DateTime.Now,
                Author = Environment.UserName
            };

            // Extract column definitions from DataWindow
            var bandNames = dataWindow.GetBandNames();
            foreach (var bandName in bandNames)
            {
                var objects = dataWindow.GetBandObjects(bandName);
                foreach (var obj in objects)
                {
                    // Extract column information and add to header
                    if (obj is DwText textObj)
                    {
                        header.Columns.Add(new ColumnDefinition
                        {
                            Name = obj.Name,
                            DataType = "text",
                            Width = _coordinateMapper.ConvertFromPixels(textObj.Attributes.Width),
                            Format = textObj.Attributes.FormatString ?? ""
                        });
                    }
                }
            }

            await writer.WriteHeaderAsync(header);
        }

        /// <summary>
        /// Processes a batch of rows with optimized memory usage
        /// </summary>
        private async Task ProcessRowBatchAsync(IPowerBuilderDataWindow dataWindow, 
            IDocumentWriter writer, int startRow, int endRow)
        {
            var rows = new List<DocumentRow>();

            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                var row = ProcessSingleRow(dataWindow, rowIndex);
                rows.Add(row);
            }

            await writer.WriteRowBatchAsync(rows);
        }

        /// <summary>
        /// Processes a single row with all its cells
        /// </summary>
        private DocumentRow ProcessSingleRow(IPowerBuilderDataWindow dataWindow, int rowIndex)
        {
            var row = new DocumentRow { RowIndex = rowIndex };

            // Process each band's objects for this row
            var bandNames = dataWindow.GetBandNames();
            foreach (var bandName in bandNames)
            {
                row.BandName = bandName;
                var objects = dataWindow.GetBandObjects(bandName);

                foreach (var obj in objects)
                {
                    var cell = ProcessObject(obj);
                    row.Cells.Add(cell);
                }
            }

            return row;
        }

        /// <summary>
        /// Processes a single DataWindow object into a document cell
        /// </summary>
        private DocumentCell ProcessObject(IDwObject obj)
        {
            var cell = new DocumentCell();

            switch (obj)
            {
                case DwText textObj:
                    cell.Value = textObj.Attributes.Text;
                    cell.FormattedValue = textObj.Attributes.Text ?? "";
                    cell.Font = _styleMapper.MapFont(
                        textObj.Attributes.FontFace,
                        textObj.Attributes.FontSize,
                        textObj.Attributes.FontWeight,
                        textObj.Attributes.Italics,
                        textObj.Attributes.Underline,
                        textObj.Attributes.Strikethrough);
                    cell.BackgroundColor = _styleMapper.MapColor(textObj.Attributes.BackgroundColor.Value);
                    cell.ForegroundColor = _styleMapper.MapColor(textObj.Attributes.FontColor.Value);
                    cell.X = textObj.Attributes.X;
                    cell.Y = textObj.Attributes.Y;
                    cell.Width = textObj.Attributes.Width;
                    cell.Height = textObj.Attributes.Height;
                    break;

                case DwCompute computeObj:
                    cell.Value = computeObj.Attributes.Text;
                    cell.FormattedValue = computeObj.Attributes.Text ?? "";
                    cell.Font = _styleMapper.MapFont(
                        computeObj.Attributes.FontFace,
                        computeObj.Attributes.FontSize,
                        computeObj.Attributes.FontWeight,
                        computeObj.Attributes.Italics,
                        computeObj.Attributes.Underline,
                        computeObj.Attributes.Strikethrough);
                    cell.BackgroundColor = _styleMapper.MapColor(computeObj.Attributes.BackgroundColor.Value);
                    cell.ForegroundColor = _styleMapper.MapColor(computeObj.Attributes.FontColor.Value);
                    cell.X = computeObj.Attributes.X;
                    cell.Y = computeObj.Attributes.Y;
                    cell.Width = computeObj.Attributes.Width;
                    cell.Height = computeObj.Attributes.Height;
                    break;

                default:
                    cell.Value = obj.Name;
                    cell.FormattedValue = obj.Name;
                    break;
            }

            return cell;
        }

        /// <summary>
        /// Processes footer section
        /// </summary>
        private async Task ProcessFooterAsync(IPowerBuilderDataWindow dataWindow, 
            IDocumentWriter writer, int totalRows)
        {
            var footer = new DocumentFooter
            {
                TotalRows = totalRows,
                CompletedDate = DateTime.Now,
                Summary = $"Export completed with {totalRows} rows"
            };

            await writer.WriteFooterAsync(footer);
        }

        /// <summary>
        /// Monitors memory usage and triggers garbage collection if needed
        /// </summary>
        private async Task CheckMemoryUsageAsync()
        {
            var currentMemory = GC.GetTotalMemory(false);
            
            if (currentMemory > MaxMemoryThreshold)
            {
                // Force garbage collection and wait for finalizers
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // Add a small delay to allow memory to be reclaimed
                await Task.Delay(10);
            }
        }

        /// <summary>
        /// Creates an appropriate document writer for the output stream
        /// This would be extended to support different formats
        /// </summary>
        private IDocumentWriter CreateDocumentWriter(Stream outputStream)
        {
            // This is a placeholder - would be implemented with specific format writers
            return new NullDocumentWriter();
        }
    }

    /// <summary>
    /// Null implementation of IDocumentWriter for testing
    /// </summary>
    internal class NullDocumentWriter : IDocumentWriter
    {
        public Task WriteHeaderAsync(DocumentHeader header) => Task.CompletedTask;
        public Task WriteRowBatchAsync(IEnumerable<DocumentRow> rows) => Task.CompletedTask;
        public Task WriteFooterAsync(DocumentFooter footer) => Task.CompletedTask;
        public Task FlushAsync() => Task.CompletedTask;
        public void Dispose() { }
    }
}