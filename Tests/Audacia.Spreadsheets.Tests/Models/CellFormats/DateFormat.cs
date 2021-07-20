using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class DateFormat
    {
        [CellFormat(CellFormat.Date)]
        public DateTime Value { get; set; }

        public static implicit operator DateFormat(DateTime dt)
        {
            return new DateFormat
            {
                Value = dt
            };
        }
    }
}
