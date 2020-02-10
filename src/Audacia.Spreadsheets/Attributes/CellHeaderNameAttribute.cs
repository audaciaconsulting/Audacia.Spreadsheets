using System;

namespace Audacia.Spreadsheets.Attributes
{
    public class CellHeaderNameAttribute : Attribute
    {
        public string Name { get; set; }
    }
}