using System;
using System.ComponentModel.DataAnnotations;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Demo.Models
{
    public class Book
    {
        public string Name { get; set; }
        
        public string Author { get; set; }
        
        [Display(Name = "ISBN Number")]
        public string IsbnNumber { get; set; }

        public DateTime Published { get; set; }
        
        [CellFormat(CellFormat.Currency)]
        public decimal Price { get; set; }
        
        public TimeSpan Timespan { get; set; }

        public override string ToString()
        {
            return $"{Name}, {Author}, {Published:d}, {Price:C}, {IsbnNumber}";
        }
    }
}