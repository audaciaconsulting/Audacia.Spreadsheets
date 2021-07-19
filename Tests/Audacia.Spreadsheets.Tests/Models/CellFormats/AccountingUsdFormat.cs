using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class AccountingUsdFormat
    {
        [CellFormat(CellFormat.AccountingUSD)]
        public decimal Value { get; set; }
    }
}
