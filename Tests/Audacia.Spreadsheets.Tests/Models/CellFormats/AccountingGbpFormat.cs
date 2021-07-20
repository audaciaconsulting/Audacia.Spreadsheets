using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class AccountingGbpFormat
    {
        [CellFormat(CellFormat.AccountingGBP)]
        public decimal Value { get; set; }

        public static implicit operator AccountingGbpFormat(decimal d)
        {
            return new AccountingGbpFormat
            {
                Value = d
            };
        }
    }
}
