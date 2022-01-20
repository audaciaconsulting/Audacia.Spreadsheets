using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TimeWithSecondsModel
    {
        [CellFormat(CellFormat.TimeWithSeconds)]
        public TimeSpan Value { get; set; }

        public static implicit operator TimeWithSecondsModel(TimeSpan t)
        {
            return new TimeWithSecondsModel
            {
                Value = t
            };
        }
    }
}
