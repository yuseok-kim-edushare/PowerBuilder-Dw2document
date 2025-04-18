using yuseok.kim.dw2docs.Common.VirtualGridWriter.Models;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Common;
using System.Collections.Concurrent;

namespace yuseok.kim.dw2docs.Docx.VirtualGridWriter.Renderers;

public class RendererLocator<TContext>
{
    private readonly ConcurrentDictionary<Type, Func<IDwObject, IRenderer<TContext>>> _rendererFactories = new();
    private readonly TContext _context;

    public RendererLocator(TContext context)
    {
        _context = context;
    }

    public void RegisterRenderer<TModel>(Func<TModel, RendererLocator<TContext>, IRenderer<TContext>> factory)
        where TModel : IDwObject
    {
        _rendererFactories[typeof(TModel)] = (model) => factory((TModel)model, this);
    }

    public IRenderer<TContext> GetRenderer<TModel>(TModel model)
        where TModel : IDwObject
    {
        if (_rendererFactories.TryGetValue(typeof(TModel), out var factory))
        {
            return factory(model);
        }

        throw new InvalidOperationException($"No renderer registered for model type {typeof(TModel).Name}");
    }
}
