using System.Collections.Generic;
using System.Linq;

namespace Audacia.Spreadsheets.Extensions
{
    public static class TableColumns
    {
        /// <summary>
        /// Creates a dictionary that maps column header to column index.
        /// </summary>
        /// <param name="columns">Worksheet table columns</param>
        public static IDictionary<string, int> ToDictionary(this IEnumerable<TableColumn> columns)
        {
            return columns
                .Select((col, index) => (index, col.Name.Trim()))
                .ToDictionary(c => c.Item2, c => c.index);
        }
    }
}