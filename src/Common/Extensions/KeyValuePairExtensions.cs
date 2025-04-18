using System.Collections.Generic;

namespace yuseok.kim.dw2docs.Common.Extensions
{
    /// <summary>
    /// Extension methods for KeyValuePair to support deconstruction in .NET Framework 4.8.1
    /// </summary>
    public static class KeyValuePairExtensions
    {
        // Extension method to allow deconstruction of KeyValuePair<TKey, TValue>
        public static void Deconstruct<TKey, TValue>(this KeyValuePair<TKey, TValue> kvp, out TKey key, out TValue value)
        {
            key = kvp.Key;
            value = kvp.Value;
        }
    }
} 