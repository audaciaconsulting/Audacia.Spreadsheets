using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class DateTimeFormat
    {
        [CellFormat(CellFormat.DateTime)]
        public DateTime Value { get; set; }

        public static implicit operator DateTimeFormat(DateTime dt)
        {
            return new DateTimeFormat
            {
                Value = dt
            };
        }
    }
}
