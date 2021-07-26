using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class AccountingGbpModel
    {
        [CellFormat(CellFormat.AccountingGBP)]
        public decimal Value { get; set; }

        public static implicit operator AccountingGbpModel(decimal d)
        {
            return new AccountingGbpModel
            {
                Value = d
            };
        }
    }
}
