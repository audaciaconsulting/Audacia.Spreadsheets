using System.Collections.Generic;

namespace Audacia.Spreadsheets.Extensions
{
    public static class ListExtensions
    {
        public static void AddRange<T>(this IList<T> source, IEnumerable<T> range) where T : class
        {
            foreach (var item in range)
            {
                source.Add(item);
            }
        }
    }
}