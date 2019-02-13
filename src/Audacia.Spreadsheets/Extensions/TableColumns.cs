using System.Collections.Generic;
using System.Linq;

namespace Audacia.Spreadsheets.Extensions
{
    public static class TableColumns
    {
        public static IDictionary<string, int> ToDictionary(this IEnumerable<TableColumn> columns)
        {
            return columns
                .Select((col, index) => (index, col.Name.Trim()))
                .ToDictionary(c => c.Item2, c => c.Item1);
        }
    }
}