using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TimeSpanHoursFormat
    {
        [CellFormat(CellFormat.TimeSpanHours)]
        public TimeSpan Value { get; set; }

        public static implicit operator TimeSpanHoursFormat(TimeSpan t)
        {
            return new TimeSpanHoursFormat
            {
                Value = t
            };
        }
    }
}
