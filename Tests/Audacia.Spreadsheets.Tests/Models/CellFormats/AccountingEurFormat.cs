using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class AccountingEurFormat
    {
        [CellFormat(CellFormat.AccountingEUR)]
        public decimal Value { get; set; }
    }
}
