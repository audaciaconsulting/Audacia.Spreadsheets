using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class CellStyle
    {
        public UInt32Value Index { get; set; } = null!;

        public bool BorderTop { get; set; }
        
        public bool BorderRight { get; set; }
        
        public bool BorderBottom { get; set; }
        
        public bool BorderLeft { get; set; }

        public uint BackgroundColour { get; set; }
        
        public uint TextColour { get; set; }

        public CellFormat Format { get; set; }

        public HorizontalAlignmentValues AlignHorizontal { get; set; }

        public VerticalAlignmentValues AlignVertical { get; set; }

        public bool HasWordWrap { get; set; }

        public bool IsEditable { get; set; }
    }
}
