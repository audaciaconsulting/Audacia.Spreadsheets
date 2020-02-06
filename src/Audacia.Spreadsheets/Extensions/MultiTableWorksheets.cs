using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Audacia.Core.Extensions;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Extensions
{
    /// <summary>
    /// These extensions exist solely for the Gymfinity project.
    /// Please try to avoid from using them, they're not amazing.
    /// </summary>
    public static class MultiTableWorksheets
    {
        private static string PascalToTitleCase(this string input)
        {
            return Regex.Replace(input, "([a-z](?=[A-Z]|[0-9])|[A-Z](?=[A-Z][a-z]|[0-9])|[0-9](?=[^0-9]))", "$1 ");
        }

        public static IList<TableColumn> GetColumns<T>(params string[] propertiesToIgnore)
        {
            var columns = new List<TableColumn>();
            var properties = typeof(T).GetProps();
            
            foreach (var property in properties)
            {
                var formatAttribute =
                    (CellFormatAttribute[])property.GetCustomAttributes(typeof(CellFormatAttribute), false);
                var nameAttribute =
                    (CellHeaderNameAttribute[])property.GetCustomAttributes(typeof(CellHeaderNameAttribute), false);
                
                var name = nameAttribute.FirstOrDefault()?.Name ?? property.GetDataAnnotationDisplayName();
                var column = new TableColumn(name.PascalToTitleCase());
                
                if (formatAttribute.Any())
                {
                    column.Format = formatAttribute.First().CellFormat;
                }
                
                columns.Add(column);
            }
            
            return columns;
        }
        
        public static IEnumerable<TableRow> GetRows<T>(this IEnumerable<T> dataList, params string[] propertiesToIgnore)
        {
            var properties = typeof(T).GetProps();
            
            return dataList.Select(entry =>
            {
                var cells = properties
                    .Select(prop => new TableCell(prop.GetValue(entry)))
                    .ToList();
                return TableRow.FromCells(cells, null);
            });
        }
        
        public static Table ToSpreadsheetTable<T>(this List<T> dataList, IList<string> propertiesToIgnore)
        {
            var table = new Table(true);
            var columns = GetColumns<T>(propertiesToIgnore.ToArray());
            table.Columns.AddRange(columns);
            table.Rows = dataList.GetRows(propertiesToIgnore.ToArray());
            return table;
        }
    }
}