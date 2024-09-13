using Audacia.Spreadsheets.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Audacia.Spreadsheets.Tests
{
    public class NoErrorOnEmptyArraySheet : Worksheet
    {
        public static List<string> ColumnHeaders = new List<string>() { "test1", "test2", "test3" };
        public NoErrorOnEmptyArraySheet()
        {
            SheetName = "Test Sheet";
            ShowGridLines = true;
            Table = new Table(true)
            {
                Columns = TableColumns.ToList(),
            };
        }

        public static IEnumerable<TableColumn> TableColumns => ColumnHeaders.Select(column => new TableColumn(column));
    }
   
    public class OverallConversionTests
    {
        [Fact]
        public void No_Error_On_Empty_Array()
        {
            var worksheet = new NoErrorOnEmptyArraySheet();

            var importer = new WorksheetImporter<NoErrorOnEmptyArraySheet>();
            var spreadsheet = Spreadsheet.FromWorksheets(worksheet);
            var bytes = spreadsheet.Export();
            var testsheet = Spreadsheet.FromBytes(bytes);
            var columns = testsheet.Worksheets[0].GetTable().Columns.ConvertAll(c => c.Name);
            Assert.True(columns.All(c => NoErrorOnEmptyArraySheet.TableColumns.Any(tc => tc.Name == c)));
        }
    }
}
