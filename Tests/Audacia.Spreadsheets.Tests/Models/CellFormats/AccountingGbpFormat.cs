using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class AccountingGbpFormat
    {
        [CellFormat(CellFormat.AccountingGBP)]
        public decimal Value { get; set; }
    }
}
