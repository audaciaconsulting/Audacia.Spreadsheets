using System;

namespace Audacia.Spreadsheets.Demo.Models
{
    public class Book
    {
        public string Name { get; set; }
        
        public string Author { get; set; }

        public string IsbnNumber { get; set; }

        public DateTime Published { get; set; }

        public decimal Price { get; set; }

        public override string ToString()
        {
            return $"{Name}, {Author}, {Published:d}, {Price:C}, {IsbnNumber}";
        }
    }
}