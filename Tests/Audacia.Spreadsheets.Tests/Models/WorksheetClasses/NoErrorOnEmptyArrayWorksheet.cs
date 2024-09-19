using System.Collections.Generic;
using System.Linq;

namespace Audacia.Spreadsheets.Tests.Models.WorksheetClasses
{
    public class NoErrorOnEmptyArrayWorksheet : Worksheet
    {
        private static List<string> ColumnHeaders { get; set; } = new List<string>() { "test1", "test2", "test3" };

        public static List<string> GetColumnHeaders => ColumnHeaders;

        public NoErrorOnEmptyArrayWorksheet()
        {
            SheetName = "Test Sheet";
            ShowGridLines = true;
            Table = new Table(true)
            {
                Columns = TableColumns.ToList()
            };
        }

        public static IEnumerable<TableColumn> TableColumns => ColumnHeaders.Select(column => new TableColumn(column));
    }
}