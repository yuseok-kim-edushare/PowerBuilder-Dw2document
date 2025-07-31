using System;
using System.Collections.Generic;
using System.IO;
using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Abstractions;
using yuseok.kim.dw2docs.Common.Utils;

namespace yuseok.kim.dw2docs
{
    /// <summary>
    /// Builder class for creating documents from virtual grid data with robust error handling
    /// </summary>
    public class DocumentBuilder
    {
        private VirtualGrid? _grid;
        private string? _outputPath;
        private readonly Dictionary<string, object> _options;

        public DocumentBuilder()
        {
            _options = new Dictionary<string, object>();
        }

        /// <summary>
        /// Set the virtual grid data source for document generation
        /// </summary>
        /// <param name="grid">The virtual grid containing data to export</param>
        /// <returns>This DocumentBuilder instance for method chaining</returns>
        public DocumentBuilder WithGrid(VirtualGrid grid)
        {
            try
            {
                
                if (grid == null)
                {
                    throw new ArgumentNullException(nameof(grid), "Virtual grid cannot be null");
                }

                // Additional validation for grid contents
                if (grid.CellRepository == null)
                {
                    throw new ArgumentException("Virtual grid must have a valid cell repository", nameof(grid));
                }

                _grid = grid;
                return this;
            }
            catch (Exception ex)
            {
                FileLogger.LogToFile($"[DocumentBuilder.WithGrid] Error setting grid: {ex.Message}", ex);
                throw; // Re-throw to maintain original behavior
            }
        }

        /// <summary>
        /// Set the output path for the generated document
        /// </summary>
        /// <param name="outputPath">File path where document will be saved</param>
        /// <returns>This DocumentBuilder instance for method chaining</returns>
        public DocumentBuilder WithOutputPath(string outputPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(outputPath))
                {
                    throw new ArgumentException("Output path cannot be null or empty", nameof(outputPath));
                }

                // Added validation for directory accessibility
                var directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory))
                {
                    try
                    {
                        // Added directory existence check and creation
                        if (!Directory.Exists(directory))
                        {
                            Directory.CreateDirectory(directory);
                        }

                        // Added write permission check
                        var testFile = Path.Combine(directory, $"test_write_{Guid.NewGuid()}.tmp");
                        try
                        {
                            File.WriteAllText(testFile, "test");
                            File.Delete(testFile);
                        }
                        catch (UnauthorizedAccessException)
                        {
                            throw new UnauthorizedAccessException($"No write permission for directory: {directory}");
                        }
                    }
                    catch (Exception ex) when (!(ex is UnauthorizedAccessException))
                    {
                        // Added I/O error handling
                        throw new IOException($"Cannot access or create directory: {directory}", ex);
                    }
                }

                _outputPath = outputPath;
                return this;
            }
            catch (Exception ex)
            {
                FileLogger.LogToFile($"[DocumentBuilder.WithOutputPath] Error setting output path: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// Add configuration options for document generation
        /// </summary>
        /// <param name="key">Option key</param>
        /// <param name="value">Option value</param>
        /// <returns>This DocumentBuilder instance for method chaining</returns>
        public DocumentBuilder WithOption(string key, object value)
        {
            try
            {
                // Added null check for option key
                if (string.IsNullOrWhiteSpace(key))
                {
                    throw new ArgumentException("Option key cannot be null or empty", nameof(key));
                }

                // Added null check for option value
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value), "Option value cannot be null");
                }

                _options[key] = value;
                return this;
            }
            catch (Exception ex)
            {
                FileLogger.LogToFile($"[DocumentBuilder.WithOption] Error setting option {key}: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// Build and generate the document using the configured settings
        /// </summary>
        /// <param name="writer">The document writer to use for generation</param>
        /// <returns>Success message or error details</returns>
        public string Build(AbstractVirtualGridWriter writer)
        {
            try
            {
                // Added comprehensive validation before building
                ValidateConfiguration();

                // Added null check for writer parameter
                if (writer == null)
                {
                    throw new ArgumentNullException(nameof(writer), "Document writer cannot be null");
                }

                // Added grid validation
                if (_grid == null)
                {
                    throw new InvalidOperationException("Virtual grid must be set before building document");
                }

                // Added output path validation
                if (string.IsNullOrWhiteSpace(_outputPath))
                {
                    throw new InvalidOperationException("Output path must be set before building document");
                }

                try
                {
                    // Attempt to generate the document
                    bool success = writer.WriteEntireGrid(_outputPath, out string? error);
                    
                    if (!success)
                    {
                        return $"Error: Document generation failed - {error ?? "Unknown error"}";
                    }

                    // Added file existence validation
                    if (!File.Exists(_outputPath))
                    {
                        return "Error: Document was not created successfully - file not found after generation";
                    }

                    // Added file size validation
                    var fileInfo = new FileInfo(_outputPath);
                    if (fileInfo.Length == 0)
                    {
                        return "Error: Generated document is empty";
                    }

                    return $"Success: Document created at {_outputPath} (Size: {fileInfo.Length} bytes)";
                }
                catch (UnauthorizedAccessException ex)
                {
                    // Added specific handling for access denied errors
                    var errorMsg = $"Access denied while creating document at {_outputPath}: {ex.Message}";
                    FileLogger.LogToFile($"[DocumentBuilder.Build] {errorMsg}", ex);
                    return $"Error: {errorMsg}";
                }
                catch (IOException ex)
                {
                    // Added specific handling for I/O errors
                    var errorMsg = $"I/O error while creating document: {ex.Message}";
                    FileLogger.LogToFile($"[DocumentBuilder.Build] {errorMsg}", ex);
                    return $"Error: {errorMsg}";
                }
                catch (OutOfMemoryException ex)
                {
                    // Added handling for memory issues with large documents
                    var errorMsg = "Insufficient memory to generate document - document may be too large";
                    FileLogger.LogToFile($"[DocumentBuilder.Build] {errorMsg}: {ex.Message}", ex);
                    return $"Error: {errorMsg}";
                }
            }
            catch (Exception ex)
            {
                // Added catch-all exception handler with logging
                var errorMsg = $"Unexpected error during document building: {ex.Message}";
                FileLogger.LogToFile($"[DocumentBuilder.Build] {errorMsg}\n{ex.StackTrace}", ex);
                return $"Error: {errorMsg}";
            }
            finally
            {
                // Added cleanup in finally block
                try
                {
                    if (writer is IDisposable disposable)
                    {
                        disposable.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    // Added logging for disposal errors
                    FileLogger.LogToFile($"[DocumentBuilder.Build] Error disposing writer: {ex.Message}", ex);
                }
            }
        }

        /// <summary>
        /// Validate the current configuration before attempting to build
        /// </summary>
        private void ValidateConfiguration()
        {
            try
            {
                // Added validation for required fields
                var errors = new List<string>();

                if (_grid == null)
                {
                    errors.Add("Virtual grid is not set");
                }
                else
                {
                    // Added grid content validation
                    if (_grid.CellRepository == null)
                    {
                        errors.Add("Virtual grid has no cell repository");
                    }
                    else if (_grid.CellRepository.Cells.Count == 0)
                    {
                        errors.Add("Virtual grid contains no cells");
                    }

                    if (_grid.Rows == null || _grid.Rows.Count == 0)
                    {
                        errors.Add("Virtual grid contains no rows");
                    }

                    if (_grid.Columns == null || _grid.Columns.Count == 0)
                    {
                        errors.Add("Virtual grid contains no columns");
                    }
                }

                if (string.IsNullOrWhiteSpace(_outputPath))
                {
                    errors.Add("Output path is not set");
                }

                if (errors.Count > 0)
                {
                    throw new InvalidOperationException($"Configuration validation failed: {string.Join(", ", errors)}");
                }
            }
            catch (Exception ex)
            {
                FileLogger.LogToFile($"[DocumentBuilder.ValidateConfiguration] Validation error: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// Reset the builder to its initial state
        /// </summary>
        public void Reset()
        {
            try
            {
                _grid = null;
                _outputPath = null;
                _options.Clear();
            }
            catch (Exception ex)
            {
                // Added error handling for reset operation
                FileLogger.LogToFile($"[DocumentBuilder.Reset] Error during reset: {ex.Message}", ex);
                throw;
            }
        }
    }
}