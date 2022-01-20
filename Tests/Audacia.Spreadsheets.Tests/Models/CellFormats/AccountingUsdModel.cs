using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class AccountingUsdModel
    {
        [CellFormat(CellFormat.AccountingUSD)]
        public decimal Value { get; set; }

        public static implicit operator AccountingUsdModel(decimal d)
        {
            return new AccountingUsdModel
            {
                Value = d
            };
        }
    }
}
