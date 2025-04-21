# Powerbuilder Datawindow to Document(Xlsx, Docx, ...?)
This Project is .NET project that implement **Converter Powerbuilder Datawindow to MS-Office document and so on**
This project inspired from Appeon's [Dw2Doc example Project](https://github.com/Appeon/PowerBuilder-Dw2Doc-Example)

## Why Re-construct Architecture?

- We want use in production environment
- but appeon's original repository has multiple csproj and complex user-object structure
- SO, to simplifying distribution and manaing and avoiding dll collision from other modules, Single Integrated DLL is each project result is better choice

## Introduce
- We use Virtual Grid Idea to convert datawindow to other format, this is appeon's idea
- this project target .net8 and .net6 and .net481 for compatability among powerbuilder versions
- We use Polysharp to comaptability among dotnet versions
- We use NPOI to handle Office Open XML
- We use ILRepack to bundling dll

## Acknowledgements
1. **Appeon's primary idea repository encourage our challenge**
2. **[NPOI](https://github.com/nissl-lab/npoi)**
3. **[ILRepack](https://github.com/gluck/il-repack)**
4. **Dotnet Foundation and Microsoft**
5. **[PolySharp](https://github.com/Sergio0694/PolySharp)**

## Setup Instructions

### Building the .NET Library

1. Build the yuseok.kim.dw2docs project with dotnet publish:
   ```
   dotnet publish PowerBuilder-Dw2document.slnx -r windows -f <your target>
   ```
   (Possible target : net481, net6.0-windows, net8.0-windows)


### PowerBuilder Setup

1. Import the PowerBuilder object (dw_exporter.sru) into your PowerBuilder application:
   - Open your PowerBuilder application
   - Right-click on your target library in the System Tree
   - Select "Import"
   - Navigate to the location of dw_exporter.sru
   - Click Open

2. Add the .NET assembly reference to your PowerBuilder project:
   - Open your PowerBuilder project
   - Select Project → Project Properties
   - Go to ".NET Assemblies" tab
   - Click "Add"
   - Browse to the location of yuseok.kim.dw2docs.dll
   - Select the assembly and click "Open"
   - Click "OK" to save the project properties

3. Make sure the .NET assembly is accessible:
   - The yuseok.kim.dw2docs.dll file should be in your application's path
   - Either place it in a known system location or in your application's directory

## Usage

```PowerBuilder
// Create and initialize the exporter
dw_exporter lnv_exporter
lnv_exporter = CREATE dw_exporter
if not lnv_exporter.of_initialize() then
    MessageBox("Error", "Failed to initialize DataWindow exporter")
    return
end if

// Export a DataWindow to Excel
string ls_xlsx_path = "C:\Reports\MyReport.xlsx"
string ls_result
ls_result = lnv_exporter.of_export_to_excel(dw_1, ls_xlsx_path, "MySheet")

// Export a DataWindow to Word
string ls_docx_path = "C:\Reports\MyReport.docx"
ls_result = lnv_exporter.of_export_to_word(dw_1, ls_docx_path)

// Don't forget to destroy the object when done
DESTROY lnv_exporter
```

## Troubleshooting

If you encounter issues with the PowerBuilder/.NET integration:

1. Verify the .NET assembly is properly built and accessible
2. Check that the assembly is added to your PowerBuilder project
3. Make sure the class name in of_initialize() matches exactly the one in the .NET code
4. Check PowerBuilder's System Error log for .NET-related errors
5. Try using the PowerBuilder .NET Assembly Browser to verify the assembly is loaded correctly
6. If you get "Class not found" errors, make sure the namespace and class name match exactly
7. For "Method not found" errors, check parameter types and counts in both PowerBuilder and .NET 