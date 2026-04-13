#if NET462
using System.Collections.Generic;

namespace System.Collections.Generic
{
    internal static class NetStandardPolyfill
    {
        public static bool TryAdd<TKey, TValue>(this Dictionary<TKey, TValue> dict, TKey key, TValue value)
        {
            if (!dict.ContainsKey(key))
            {
                dict.Add(key, value);
                return true;
            }
            return false;
        }

        public static List<T> Append<T>(this List<T> list, T item)
        {
            list.Add(item);
            return list;
        }
    }
}
#endif
