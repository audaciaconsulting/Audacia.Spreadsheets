using System;

namespace Audacia.Spreadsheets.Models.Attributes
{
    public class CellBackgroundColourAttribute : Attribute
    {
        public string ReferenceField { get; set; }
        public string Colour { get; set; }

        public CellBackgroundColourAttribute(string colour = null)
        {
            Colour = colour;
        }
    }
}
