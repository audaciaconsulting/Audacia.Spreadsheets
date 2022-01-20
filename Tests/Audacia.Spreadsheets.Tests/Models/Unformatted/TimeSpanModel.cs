using System;

namespace Audacia.Spreadsheets.Tests.Models.Unformatted
{
    public class TimeSpanModel
    {
        public TimeSpan Value { get; set; }

        public static implicit operator TimeSpanModel(TimeSpan t)
        {
            return new TimeSpanModel
            {
                Value = t
            };
        }
    }
}
