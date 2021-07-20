﻿using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class PercentageFormat
    {
        [CellFormat(CellFormat.Percentage)]
        public decimal Value { get; set; }

        public static implicit operator PercentageFormat(decimal d)
        {
            return new PercentageFormat
            {
                Value = d
            };
        }
    }
}
