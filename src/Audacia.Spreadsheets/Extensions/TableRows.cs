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
            return rows
                .Select(row => row.GetFields())
                .ToArray();
        }
        
        /// <summary>
        /// Returns an array of cell values for the provided row.
        /// </summary>
        public static string[] GetFields(this TableRow row)
        {
            return row.Cells
                .Select(cell => cell.Value.ToString().Trim())
                .ToArray();
        }
    }
}