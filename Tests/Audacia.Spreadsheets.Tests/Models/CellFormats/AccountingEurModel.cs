using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class AccountingEurModel
    {
        [CellFormat(CellFormat.AccountingEUR)]
        public decimal Value { get; set; }

        public static implicit operator AccountingEurModel(decimal d)
        {
            return new AccountingEurModel
            {
                Value = d
            };
        }
    }
}
