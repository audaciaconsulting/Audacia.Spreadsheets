using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TimeSpanFullFormat
    {
        [CellFormat(CellFormat.TimeSpanFull)]
        public TimeSpan Value { get; set; }
    }
}
