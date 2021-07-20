using System;

namespace Audacia.Spreadsheets.Tests.Models.Unformatted
{
    public class DateTimeModel
    {
        public DateTime Value { get; set; }

        public static implicit operator DateTimeModel(DateTime dt)
        {
            return new DateTimeModel
            {
                Value = dt
            };
        }
    }
}
