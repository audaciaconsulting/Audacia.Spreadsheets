using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Audacia.Core.Extensions;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Extensions
{
    public static class Tables
    {
        /// <summary>
        /// Reads all properties on the provided class and returns a list of Table Columns.
        /// </summary>
        public static List<TableColumn> GetColumns<TEntity>(params string[] ignoreProperties)
            where TEntity : class
        {
            var tableColumns = new List<TableColumn>();
            var properties = typeof(TEntity).GetProps(ignoreProperties).ToArray();
            
            foreach (var property in properties)
            {
                var cellFormat = property.GetCustomAttributes<CellFormatAttribute>(false).FirstOrDefault();
                var cellHeader = property.GetCustomAttributes<CellHeaderNameAttribute>(false).FirstOrDefault();
                var hideCellHeader = property.GetCustomAttributes<HideHeaderAttribute>(false).Any();
                var displaySubtotal = property.GetCustomAttributes<SubtotalHeaderAttribute>(false).Any() &&
                                      property.PropertyType.IsNumeric();
                var backgroundColour = property.GetCustomAttributes<CellBackgroundColourAttribute>(false).FirstOrDefault();
                var textColour = property.GetCustomAttributes<CellTextColourAttribute>(false).FirstOrDefault();
                var columnWidth = property.GetCustomAttributes<ColumnWidthAttribute>(false).FirstOrDefault();
                
                // The column name is either the cell header, the Display attribute or the property name.
                // Please leave the cell header logic alone, this is specific to gymfinity :(
                var columnName =  cellHeader?.Name?.PascalToTitleCase()
                                  ?? property.GetDataAnnotationDisplayName();

                var column = new TableColumn
                {
                    PropertyInfo = property,
                    Name = hideCellHeader ? string.Empty : columnName,
                    DisplaySubtotal = displaySubtotal,
                    CellBackgroundFormat = backgroundColour,
                    CellTextFormat = textColour,
                    Width = columnWidth?.Width
                };

                if (cellFormat != null)
                {
                    column.Format = cellFormat.CellFormat;
                }

                tableColumns.Add(column);
            }
            
            return tableColumns;
        }
        
        /// <summary>
        /// Converts the provided enumerable into an enumerable of table rows.
        /// </summary>
        public static IEnumerable<TableRow> GetRows<TEntity>(IEnumerable<TEntity> source, IReadOnlyCollection<TableColumn> columns)
            where TEntity : class
        {
            return source.Select(entry =>
            {
                var cells = columns.Select(column =>
                {
                    var cell = new TableCell
                    {
                        Value = column.PropertyInfo.GetValue(entry) ?? string.Empty,
                        FillColour = column.CellBackgroundFormat?.Colour,
                        TextColour = column.CellTextFormat?.Colour
                    };

                    // Get FillColour from property
                    if (!string.IsNullOrEmpty(column.CellBackgroundFormat?.ReferenceField))
                    {
                        var valueOfMatchingProperty = typeof(TEntity)
                            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .Where(property => string.Equals(property.Name, column.CellBackgroundFormat.ReferenceField))
                            .Select(property => property.GetValue(entry, null) as string)
                            .FirstOrDefault();

                        cell.FillColour = valueOfMatchingProperty;
                    }
                    
                    // Get TextColour from property
                    if (!string.IsNullOrEmpty(column.CellTextFormat?.ReferenceField))
                    {
                        var valueOfMatchingProperty = typeof(TEntity)
                            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .Where(property => string.Equals(property.Name, column.CellTextFormat.ReferenceField))
                            .Select(property => property.GetValue(entry, null) as string)
                            .FirstOrDefault();

                        cell.TextColour = valueOfMatchingProperty;
                    }

                    return cell;
                });
                
                return TableRow.FromCells(cells, null);
            });
        }

        /// <summary>
        /// Creates a spreadsheet table from an enumerable.
        /// </summary>
        public static Table ToTable<TEntity>(this IEnumerable<TEntity> source,
            bool includeHeaders = true,
            TableHeaderStyle headerStyle = null,
            params string[] ignoreProperties)
            where TEntity : class
        {
            var columns = GetColumns<TEntity>(ignoreProperties);
            return new Table(includeHeaders)
            {
                HeaderStyle = headerStyle ?? new TableHeaderStyle(),
                Columns = columns,
                Rows = GetRows(source, columns)
            };
        }
    }
}