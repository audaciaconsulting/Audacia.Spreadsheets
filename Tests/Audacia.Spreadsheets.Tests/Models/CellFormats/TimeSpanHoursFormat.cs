using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TimeSpanHoursFormat
    {
        [CellFormat(CellFormat.TimeSpanHours)]
        public TimeSpan Value { get; set; }
    }
}
