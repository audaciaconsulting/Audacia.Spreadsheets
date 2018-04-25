using System;
using Audacia.Spreadsheets.Models.Enums;

namespace Audacia.Spreadsheets.Models.Attributes
{
    public class CellFormatAttribute : Attribute
    {
        public CellFormatType CellFormatType { get; set; }

        public CellFormatAttribute(CellFormatType type)
        {
            CellFormatType = type;
        }
    }
}
