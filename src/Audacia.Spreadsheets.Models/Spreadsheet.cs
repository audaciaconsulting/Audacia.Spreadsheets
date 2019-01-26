using System.Collections.Generic;

namespace Audacia.Spreadsheets.Models
{
    public class Spreadsheet
    {
        public IList<Worksheet> Worksheets { get; } = new List<Worksheet>();

        public static Spreadsheet FromWorksheets(IEnumerable<Worksheet> worksheets)
        {
            var spreadsheet = new Spreadsheet();
            foreach (var worksheet in worksheets)
            {
                spreadsheet.Worksheets.Add(worksheet);
            }

            return spreadsheet;
        }

        // TODO: Re-implement this here and remove/deprecate OpenXmlReportGenerator
        //public byte[] Bytes => new OpenXmlReportGenerator().GenerateSearchResultReport(this);
    }
}
