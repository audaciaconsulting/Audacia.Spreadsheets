using System;

namespace Audacia.Spreadsheets.Models.Attributes
{
    public class CellTextColourAttribute : Attribute
    {
        public string ReferenceField { get; set; }
        public string Colour { get; set; }

        public CellTextColourAttribute(string colour = null)
        {
            Colour = colour;
        }
    }
}
