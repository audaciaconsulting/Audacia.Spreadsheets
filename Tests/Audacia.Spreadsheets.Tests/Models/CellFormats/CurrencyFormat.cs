using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class CurrencyFormat
    {
        [CellFormat(CellFormat.Currency)]
        public decimal Value { get; set; }

        public static implicit operator CurrencyFormat(decimal d)
        {
            return new CurrencyFormat
            {
                Value = d
            };
        }
    }
}
