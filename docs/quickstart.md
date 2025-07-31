# PowerBuilder DataWindow to Document - Quick Start Guide

This guide provides minimal code snippets and instructions for using the PowerBuilder DataWindow to Document converter with the Phase 1 improvements.

## Overview

The library now includes enhanced error handling and new serialization capabilities through the `DwSerializer` and `DocumentBuilder` classes, in addition to the existing `DatawindowExporter` functionality.

## Prerequisites

- .NET 8.0, .NET 6.0, or .NET Framework 4.8.1
- Windows environment (for PowerBuilder integration)
- PowerBuilder application (for data source)

## Build Instructions

### 1. Build the Library

```bash
# Clone the repository
git clone https://github.com/yuseok-kim-edushare/PowerBuilder-Dw2document.git
cd PowerBuilder-Dw2document

# Build for specific target framework
dotnet build src/yuseok.kim.dw2docs.csproj --framework net8.0-windows
# OR
dotnet build src/yuseok.kim.dw2docs.csproj --framework net481
# OR  
dotnet build src/yuseok.kim.dw2docs.csproj --framework net6.0-windows

# Build all target frameworks
dotnet build src/yuseok.kim.dw2docs.csproj
```

### 2. Publish for Distribution

```bash
# Publish with ILRepack bundling
dotnet publish src/yuseok.kim.dw2docs.csproj -r windows -f net481
# OR
dotnet publish src/yuseok.kim.dw2docs.csproj -r windows -f net8.0-windows

# The bundled DLL will be available at:
# {repository_root}/bin/{target_framework}/yuseok.kim.dw2docs.dll
```

## Usage Examples

### Method 1: Using DwSerializer (Phase 1 Addition)

```csharp
using yuseok.kim.dw2docs;
using yuseok.kim.dw2docs.Common.VirtualGrid;

// Create serializer instance
var serializer = new DwSerializer();

// Export to DOCX with enhanced error handling
VirtualGrid grid = CreateVirtualGridFromData(); // Your data source
string result = serializer.ExportToDocx(grid, @"C:\Reports\MyReport.docx");

if (result.StartsWith("Success"))
{
    Console.WriteLine("Document created successfully!");
}
else
{
    Console.WriteLine($"Export failed: {result}");
}

// PDF export (currently stubbed for future implementation)
string pdfResult = serializer.ExportToPdf(grid, @"C:\Reports\MyReport.pdf");
Console.WriteLine(pdfResult); // Will show "not yet implemented" message
```

### Method 2: Using DocumentBuilder (Phase 1 Addition)

```csharp
using yuseok.kim.dw2docs;
using yuseok.kim.dw2docs.Docx.VirtualGridWriter.DocxWriter;

// Create document with builder pattern and enhanced error handling
var builder = new DocumentBuilder();
var docxWriterBuilder = new VirtualGridDocxWriterBuilder { WriteToPath = @"C:\Reports\output.docx" };

try
{
    VirtualGrid grid = CreateVirtualGridFromData(); // Your data source
    var writer = docxWriterBuilder.Build(grid, out string? error);
    
    if (writer != null)
    {
        string result = builder
            .WithGrid(grid)
            .WithOutputPath(@"C:\Reports\MyReport.docx")
            .WithOption("PageSize", "A4")
            .Build(writer);
            
        Console.WriteLine(result);
    }
    else
    {
        Console.WriteLine($"Failed to create writer: {error}");
    }
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
}
```

### Method 3: Using DatawindowExporter (Existing Method)

```csharp
using yuseok.kim.dw2docs.Interop;

// Create exporter instance
var exporter = new DatawindowExporter();

// Export to Excel
string jsonData = GetDatawindowDataAsJson(); // Your JSON data
string excelResult = exporter.ExportToExcel(jsonData, @"C:\Reports\MyReport.xlsx", "DataSheet");
Console.WriteLine(excelResult);

// Export to Word
string wordResult = exporter.ExportToWord(jsonData, @"C:\Reports\MyReport.docx");
Console.WriteLine(wordResult);
```

## PowerBuilder Integration

### 1. Import the User Object

1. In PowerBuilder IDE, right-click your target library
2. Select "Import" 
3. Navigate to `dw_exporter.sru` file
4. Click Open to import

### 2. Add .NET Assembly Reference

1. Open Project Properties
2. Go to ".NET Assemblies" tab
3. Click "Add" and browse to the generated DLL:
   - `{repository_root}\bin\net481\yuseok.kim.dw2docs.dll` (for PB 2019+)
   - `{repository_root}\bin\net6.0-windows\yuseok.kim.dw2docs.dll` (for newer PB)

### 3. PowerBuilder Usage Example

```powerbuilder
// Create and initialize the exporter
dw_exporter lnv_exporter
lnv_exporter = CREATE dw_exporter

if not lnv_exporter.of_initialize() then
    MessageBox("Error", "Failed to initialize DataWindow exporter")
    return
end if

// Export DataWindow to Excel
string ls_xlsx_path = "C:\Reports\MyReport.xlsx"
string ls_result = lnv_exporter.of_export_to_excel(dw_1, ls_xlsx_path, "MySheet")
MessageBox("Export Result", ls_result)

// Export DataWindow to Word  
string ls_docx_path = "C:\Reports\MyReport.docx"
ls_result = lnv_exporter.of_export_to_word(dw_1, ls_docx_path)
MessageBox("Export Result", ls_result)

// Clean up
DESTROY lnv_exporter
```

## Phase 1 Improvements Summary

### Task 1: DwSerializer.ExportToDocx() Enhancements
- ✅ Added comprehensive null checks and parameter validation
- ✅ Enhanced error handling with specific exception types
- ✅ Added file and directory validation before export
- ✅ Included TODO comments for style mapping improvements
- ✅ Added XML documentation for all parameters

### Task 2: DwSerializer.ExportToPdf() Implementation
- ✅ Created stub method with full parameter validation
- ✅ Added placeholder for PDF writer implementation
- ✅ Included image embedding validation stub
- ✅ Added comprehensive TODO comments for future implementation
- ✅ Basic XML documentation

### Task 3: DocumentBuilder Exception Handling
- ✅ Created DocumentBuilder class with robust error handling
- ✅ Added try-catch blocks for all major operations
- ✅ Implemented validation for null references and I/O errors
- ✅ Added specific exception handling for access denied, memory, and file issues
- ✅ Comprehensive logging for debugging and error tracking

### Task 4: Documentation and Examples
- ✅ Created this quickstart guide with minimal working examples
- ✅ Included build and run instructions
- ✅ Provided usage examples for all new classes
- ✅ PowerBuilder integration instructions

## Error Handling Improvements

The Phase 1 enhancements include comprehensive error handling:

- **Null Reference Protection**: All methods now validate input parameters
- **File I/O Error Handling**: Specific handling for access denied, file not found, and disk space issues
- **Memory Management**: Protection against out-of-memory conditions with large documents
- **Resource Cleanup**: Proper disposal of resources in finally blocks
- **Detailed Logging**: All errors are logged to `C:\temp\Dw2Doc_ExcelError.log` for debugging

## Troubleshooting

### Common Issues

1. **"Class not found" errors**
   - Verify the .NET assembly is properly referenced in PowerBuilder
   - Check that the DLL is in the application's path

2. **"Access denied" errors** 
   - Ensure write permissions for the output directory
   - Check if the target file is not already open in another application

3. **PDF export shows "not implemented"**
   - This is expected behavior in Phase 1 - PDF functionality is stubbed for future implementation

4. **Build failures**
   - Ensure you have the correct .NET SDK installed
   - Check that all NuGet packages are restored properly

### Log Files

Check these locations for debugging information:
- `C:\temp\Dw2Doc_ExcelError.log` - Main error log
- `{project_root}\src\ilrepack.log` - Build process log

## Next Steps

Phase 1 provides the foundation for enhanced document generation. Future phases will include:
- Complete PDF export implementation
- Advanced style mapping improvements  
- Image embedding functionality
- Performance optimizations
- Additional document format support

For issues or questions, refer to the repository's issue tracker or the comprehensive logging output.