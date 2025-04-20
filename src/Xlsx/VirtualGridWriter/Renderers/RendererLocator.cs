using yuseok.kim.dw2docs.Common.DwObjects;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Common;
using System.Collections.Concurrent;

namespace yuseok.kim.dw2docs.Xlsx.VirtualGridWriter.Renderers
{
    public class RendererLocator
    {
        private readonly ConcurrentDictionary<Type, ObjectRendererBase> _renderers = new();

        public RendererLocator()
        {
        }

        public void RegisterRenderer(Type objectType, ObjectRendererBase renderer)
        {
            _renderers[objectType] = renderer;
        }

        public ObjectRendererBase? Find(Type objectType)
        {
            if (_renderers.TryGetValue(objectType, out var renderer))
            {
                return renderer;
            }

            // Try to find a renderer for a parent type
            foreach (var kvp in _renderers)
            {
                if (objectType.IsSubclassOf(kvp.Key))
                {
                    return kvp.Value;
                }
            }

            return null;
        }
    }
} 