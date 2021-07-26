using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TimeSpanHoursModel
    {
        [CellFormat(CellFormat.TimeSpanHours)]
        public TimeSpan Value { get; set; }

        public static implicit operator TimeSpanHoursModel(TimeSpan t)
        {
            return new TimeSpanHoursModel
            {
                Value = t
            };
        }
    }
}
