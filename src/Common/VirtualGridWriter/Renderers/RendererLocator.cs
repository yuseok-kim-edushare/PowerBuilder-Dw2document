using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Attributes;
using yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers.Common;
using System.Reflection;

namespace yuseok.kim.dw2docs.Common.VirtualGridWriter.Renderers
{
    internal static class RendererLocator
    {
        private static readonly IDictionary<Type, ObjectRendererBase> Cache;

        static RendererLocator()
        {
            Cache = new Dictionary<Type, ObjectRendererBase>();
        }

        public static ObjectRendererBase? Find(Type attributeTypeToRender)
        {
            if (Cache.ContainsKey(attributeTypeToRender))
            {
                return Cache[attributeTypeToRender];
            }

            var renderers = Assembly.GetExecutingAssembly()
                .GetTypes()
                .Where(t => t.IsClass && !t.IsAbstract && typeof(ObjectRendererBase).IsAssignableFrom(t)
                            && t.GetCustomAttributes<RendererForAttribute>(true).Any());

            foreach (var rendererType in renderers)
            {
                var attributes = rendererType.GetCustomAttributes<RendererForAttribute>(true);
                foreach (var attrib in attributes)
                {
                    if (Cache.ContainsKey(attrib.TargetType))
                    {
                        continue;
                    }

                    var @obj = Activator.CreateInstance(attrib.OwningType) as ObjectRendererBase;
                    if (@obj is null)
                    {
                        Console.WriteLine($"Warning: Failed to create or cast renderer instance for {attrib.OwningType.FullName}");
                        continue;
                    }
                    Cache[attrib.TargetType] = @obj;
                }
            }

            return Cache.TryGetValue(attributeTypeToRender, out var renderer) ? renderer : null;
        }
    }
}
