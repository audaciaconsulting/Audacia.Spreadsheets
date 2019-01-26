using System;

namespace Audacia.Spreadsheets.Models.Attributes
{
    public class CellTextColourAttribute : Attribute
    {
        public CellTextColourAttribute() { }
        
        public CellTextColourAttribute(string colour) => Colour = colour;
                
        public string ReferenceField { get; set; }
        
        public string Colour { get; set; }
    }
}
