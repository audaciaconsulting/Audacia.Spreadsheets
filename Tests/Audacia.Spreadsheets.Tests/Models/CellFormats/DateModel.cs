using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class DateModel
    {
        [CellFormat(CellFormat.Date)]
        public DateTime Value { get; set; }

        public static implicit operator DateModel(DateTime dt)
        {
            return new DateModel
            {
                Value = dt
            };
        }
    }
}
