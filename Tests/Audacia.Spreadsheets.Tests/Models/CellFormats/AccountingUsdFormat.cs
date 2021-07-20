using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class AccountingUsdFormat
    {
        [CellFormat(CellFormat.AccountingUSD)]
        public decimal Value { get; set; }

        public static implicit operator AccountingUsdFormat(decimal d)
        {
            return new AccountingUsdFormat
            {
                Value = d
            };
        }
    }
}
