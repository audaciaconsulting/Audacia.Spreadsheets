﻿using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TimeFormat
    {
        [CellFormat(CellFormat.Time)]
        public TimeSpan Value { get; set; }
    }
}
