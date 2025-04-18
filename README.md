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