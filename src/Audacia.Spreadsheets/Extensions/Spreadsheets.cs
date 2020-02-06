using System;
using System.Collections.Generic;

namespace Audacia.Spreadsheets.Extensions
{
    public static class Spreadsheets
    {
        internal static IEnumerable<Table> GetTables(this Spreadsheet spreadsheet)
        {
            foreach (var ws in spreadsheet.Worksheets)
            {
                if (ws is Worksheet singleTableWorksheet)
                {
                    yield return singleTableWorksheet.Table;
                }
                else if (ws is MultiTableWorksheet multiTableWorksheet)
                {
                    foreach (var table in multiTableWorksheet.Tables)
                    {
                        yield return table;
                    }                
                }
                else
                {
                    throw new NotSupportedException($"Type of {ws?.GetType()} is not supported by GetTables().");
                }
            }
        }
    }
}