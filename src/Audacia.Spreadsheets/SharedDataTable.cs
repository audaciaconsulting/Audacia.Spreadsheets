using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    /// <summary>
    /// Global data object that can be used by any worksheet.
    /// </summary>
    public class SharedDataTable
    {
        public IList<CellStyle> CellFormats { get; } = new List<CellStyle>();

        public DefinedNames DefinedNames { get; } = new DefinedNames();

        public Stylesheet Stylesheet { get; set; }

        public IDictionary<string, uint> FillColours { get; } = new Dictionary<string, uint>();

        public IDictionary<string, uint> TextColours { get; } = new Dictionary<string, uint>();

        public IDictionary<string, uint> Fonts { get; } = new Dictionary<string, uint>();
    }
}