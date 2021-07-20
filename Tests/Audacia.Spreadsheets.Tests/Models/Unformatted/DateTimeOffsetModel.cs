using System;

namespace Audacia.Spreadsheets.Tests.Models.Unformatted
{
    public class DateTimeOffsetModel
    {
        public DateTimeOffset Value { get; set; }

        public static implicit operator DateTimeOffsetModel(DateTimeOffset dt)
        {
            return new DateTimeOffsetModel
            {
                Value = dt
            };
        }
    }
}
