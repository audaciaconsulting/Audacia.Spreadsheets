using System;
using System.ComponentModel.DataAnnotations;
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Attributes;

namespace Demo.Entities
{
    public class Book
    {
        public string Name { get; set; }
        
        public string Author { get; set; }
        
        [Display(Name = "ISBN Number")]
        public string IsbnNumber { get; set; }
        
        public DateTime Published { get; set; }
        
        [CellFormat(CellFormatType.Currency)]
        public decimal Price { get; set; }
    }
}