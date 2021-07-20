using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class DateVariantFormat
    {
        [CellFormat(CellFormat.DateVariant)]
        public DateTime Value { get; set; }

        public static implicit operator DateVariantFormat(DateTime dt)
        {
            return new DateVariantFormat
            {
                Value = dt
            };
        }
    }
}
