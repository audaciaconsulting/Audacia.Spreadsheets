using System;

namespace Audacia.Spreadsheets.Models.Attributes
{
    public class CellBackgroundColourAttribute : Attribute
    {
        public CellBackgroundColourAttribute() { }

        public CellBackgroundColourAttribute(string colour) => Colour = colour;

        public string ReferenceField { get; set; }

        public string Colour { get; set; }
    }
}
