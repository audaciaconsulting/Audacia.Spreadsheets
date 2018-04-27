using System;
using System.Collections.Generic;
using System.Text;

namespace Audacia.Spreadsheets.Models.WorksheetData
{
    public class SpreadsheetHeaderStyle
    {
        public string TextColour { get; set; }
        public string FillColour { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool HasStrike { get; set; }
        public double FontSize { get; set; } = 10d;
        public string FontName { get; set; } = "Calibri";
    }
}
