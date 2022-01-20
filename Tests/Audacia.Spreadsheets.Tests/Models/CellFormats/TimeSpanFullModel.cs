using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TimeSpanFullModel
    {
        [CellFormat(CellFormat.TimeSpanFull)]
        public TimeSpan Value { get; set; }

        public static implicit operator TimeSpanFullModel(TimeSpan t)
        {
            return new TimeSpanFullModel
            {
                Value = t
            };
        }
    }
}
