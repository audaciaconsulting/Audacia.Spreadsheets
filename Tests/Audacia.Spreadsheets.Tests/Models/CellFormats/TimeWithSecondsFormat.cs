using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TimeWithSecondsFormat
    {
        [CellFormat(CellFormat.TimeWithSeconds)]
        public TimeSpan Value { get; set; }
    }
}
