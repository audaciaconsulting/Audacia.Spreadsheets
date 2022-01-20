using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class DateVariantModel
    {
        [CellFormat(CellFormat.DateVariant)]
        public DateTime Value { get; set; }

        public static implicit operator DateVariantModel(DateTime dt)
        {
            return new DateVariantModel
            {
                Value = dt
            };
        }
    }
}
