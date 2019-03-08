using System.Collections.Generic;
using System.Linq;

namespace Audacia.Spreadsheets.Extensions
{
    public static class TableRows
    {
        /// <summary>
        /// Returns an array of values for the provided rows.
        /// </summary>
        public static string[][] GetFields(this IEnumerable<TableRow> rows)
        {
            return rows.Select(GetFields).ToArray();
        }
        
        /// <summary>
        /// Returns an array of cell values for the provided row.
        /// </summary>
        public static string[] GetFields(this TableRow row)
        {
            return row.Cells.Select(TableCells.GetValue).ToArray();
        }

        /// <summary>
        /// Returns true if the row contains no data.
        /// </summary>
        public static bool IsEmpty(this TableRow row)
        {
            return row.Cells.Select(TableCells.GetValue).All(string.IsNullOrEmpty);
        }
    }
}