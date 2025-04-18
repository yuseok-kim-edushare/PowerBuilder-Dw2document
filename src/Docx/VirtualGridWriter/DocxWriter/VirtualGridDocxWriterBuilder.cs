using yuseok.kim.dw2docs.Common.VirtualGrid;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Abstractions;

namespace yuseok.kim.dw2docs.Docx.VirtualGridWriter.DocxWriter
{
    internal class VirtualGridDocxWriterBuilder : IVirtualGridWriterBuilder
    {
        public AbstractVirtualGridWriter? Build(VirtualGrid grid, out string? error)
        {
            error = null;
            return new VirtualGridDocxWriter(grid);
        }

        public AbstractVirtualGridWriter? BuildFromTemplate(VirtualGrid grid, string path, out string? error)
        {
            throw new NotImplementedException();
        }
    }
}
