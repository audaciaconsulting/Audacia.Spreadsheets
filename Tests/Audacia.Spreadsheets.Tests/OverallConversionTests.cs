using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Audacia.Spreadsheets.Extensions;
using Xunit;
using Audacia.Spreadsheets.Tests.Models.WorksheetClasses;

namespace Audacia.Spreadsheets.Tests
{
    public class OverallConversionTests
    {
        [Fact]
        public void NoErrorOnEmptyArray()
        {
            var worksheet = new NoErrorOnEmptyArrayWorksheet();

            var importer = new WorksheetImporter<NoErrorOnEmptyArrayWorksheet>();
            var spreadsheet = Spreadsheet.FromWorksheets(worksheet);
            var bytes = spreadsheet.Export();
            var testsheet = Spreadsheet.FromBytes(bytes);
            var columns = testsheet.Worksheets[0].GetTable().Columns.ConvertAll(c => c.Name);
            Assert.True(columns.All(c => NoErrorOnEmptyArrayWorksheet.TableColumns.Any(tc => tc.Name == c)));
        }
    }
}
