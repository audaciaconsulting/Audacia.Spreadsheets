using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class DateTimeModel
    {
        [CellFormat(CellFormat.DateTime)]
        public DateTime Value { get; set; }

        public static implicit operator DateTimeModel(DateTime dt)
        {
            return new DateTimeModel
            {
                Value = dt
            };
        }
    }
}
