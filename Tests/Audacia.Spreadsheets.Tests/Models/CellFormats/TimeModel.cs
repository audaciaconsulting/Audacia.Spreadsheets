using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TimeModel
    {
        [CellFormat(CellFormat.Time)]
        public TimeSpan Value { get; set; }

        public static implicit operator TimeModel(TimeSpan t)
        {
            return new TimeModel
            {
                Value = t
            };
        }
    }
}
