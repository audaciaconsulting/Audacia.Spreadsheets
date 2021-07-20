using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TimeSpanMinutesModel
    {
        [CellFormat(CellFormat.TimeSpanMinutes)]
        public TimeSpan Value { get; set; }

        public static implicit operator TimeSpanMinutesModel(TimeSpan t)
        {
            return new TimeSpanMinutesModel
            {
                Value = t
            };
        }
    }
}
