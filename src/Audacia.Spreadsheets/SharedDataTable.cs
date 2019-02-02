using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    /// <summary>
    /// Global data object that can be used by any worksheet.
    /// </summary>
    public class SharedDataTable
    {
        public List<CellStyle> CellFormats { get; set; } = new List<CellStyle>();

        public DefinedNames DefinedNames { get; set; } = new DefinedNames();

        public Stylesheet Stylesheet { get; set; }
        
        public Dictionary<string, uint> FillColours { get; set; }

        public Dictionary<string, uint> TextColours { get; set; } 

        public Dictionary<string, uint> Fonts { get; set; }
    }
}