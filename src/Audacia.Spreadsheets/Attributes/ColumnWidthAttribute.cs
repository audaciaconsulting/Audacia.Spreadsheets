using System;

namespace Audacia.Spreadsheets.Attributes
{
    public class ColumnWidthAttribute : Attribute
    {
        public ColumnWidthAttribute() { }

        public ColumnWidthAttribute(int value) => Width = value;
        
        public int Width { get; set; }
    }
}