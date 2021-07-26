using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class CurrencyModel
    {
        [CellFormat(CellFormat.Currency)]
        public decimal Value { get; set; }

        public static implicit operator CurrencyModel(decimal d)
        {
            return new CurrencyModel
            {
                Value = d
            };
        }
    }
}
