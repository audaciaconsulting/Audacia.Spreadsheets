namespace Audacia.Spreadsheets.Demo.Models
{
    public class Stock
    {
        public string PartNumber { get; set; }

        public string Manufacturer { get; set; }

        public string Condition { get; set; }

        public int StockLevel { get; set; }

        public override string ToString()
        {
            return $"{PartNumber}, {Manufacturer}, {Condition}, {StockLevel}";
        }
    }
}
