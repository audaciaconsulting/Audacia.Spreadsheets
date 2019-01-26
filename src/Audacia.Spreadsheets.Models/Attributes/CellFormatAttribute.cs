using System;

namespace Audacia.Spreadsheets.Models.Attributes
{
    public class CellFormatAttribute : Attribute
    {
        public CellFormatAttribute(CellFormatType type)
        {
            CellFormatType = type;
        }
        
        public CellFormatType CellFormatType { get; set; }
    }
}
