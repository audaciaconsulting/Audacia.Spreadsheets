using Audacia.Spreadsheets.Models;
using DocumentFormat.OpenXml;

namespace Audacia.Spreadsheets.Export
{
    public class SpreadsheetCellStyle
    {
        public UInt32Value Index { get; set; }
        
        public bool BorderTop { get; set; }
        public bool BorderRight { get; set; }
        public bool BorderBottom { get; set; }
        public bool BorderLeft { get; set; }

        public uint BackgroundColour { get; set; }
        public uint TextColour { get; set; }

        public CellFormatType Format { get; set; }
        public bool HasWordWrap { get; set; }
    }
}
