using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Abstractions;
using System.IO;

namespace yuseok.kim.dw2docs.Docx.VirtualGridWriter.DocxWriter
{
    public class VirtualGridDocxWriterBuilder : IVirtualGridWriterBuilder
    {
        public string? WriteToPath { get; set; }
        
        public AbstractVirtualGridWriter? Build(VirtualGrid grid, out string? error)
        {
            error = null;
            
            if (string.IsNullOrEmpty(WriteToPath))
            {
                error = "Must specify a path";
                return null;
            }
            
            return new VirtualGridDocxWriter(grid, WriteToPath);
        }

        public AbstractVirtualGridWriter? BuildFromTemplate(VirtualGrid grid, string path, out string? error)
        {
            throw new NotImplementedException();
        }
    }
}
