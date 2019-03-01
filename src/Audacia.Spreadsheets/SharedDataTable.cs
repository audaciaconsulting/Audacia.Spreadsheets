using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    /// <summary>
    /// Global data object that can be used by any worksheet.
    /// </summary>
    public class SharedDataTable
    {
        public List<CellStyle> CellFormats { get; } = new List<CellStyle>();

        public DefinedNames DefinedNames { get; } = new DefinedNames();

        public Stylesheet Stylesheet { get; set; }

        public Dictionary<string, uint> FillColours { get; } = new Dictionary<string, uint>();

        public Dictionary<string, uint> TextColours { get; } = new Dictionary<string, uint>();

        public Dictionary<string, uint> Fonts { get; } = new Dictionary<string, uint>();
    }
}