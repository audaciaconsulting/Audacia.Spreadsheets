using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TimeSpanMinutesFormat
    {
        [CellFormat(CellFormat.TimeSpanMinutes)]
        public TimeSpan Value { get; set; }
    }
}
