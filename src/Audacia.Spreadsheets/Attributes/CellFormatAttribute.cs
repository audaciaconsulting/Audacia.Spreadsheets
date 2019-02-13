using System;

namespace Audacia.Spreadsheets.Attributes
{
    public class CellFormatAttribute : Attribute
    {
        public CellFormatAttribute(CellFormat type)
        {
            CellFormat = type;
        }
        
        public CellFormat CellFormat { get; set; }
    }
}
