using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using Audacia.Core.Extensions;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Extensions
{
    public static class Worksheets
    {
        private static IEnumerable<PropertyInfo> GetProps(this Type classType, params string[] ignoreProperties)
        {
            return classType
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => !ignoreProperties.Contains(p.Name))
                .Where(p =>
                {
                    var underlyingType = p.PropertyType.GetUnderlyingTypeIfNullable();
                    return underlyingType.IsValueType || underlyingType == typeof(string);
                })
                .ToArray();
        }

        private static IEnumerable<PropertyInfo> GetBaseProps(this Type classType, params string[] ignoreProperties)
        {
            var baseType = classType.BaseType;
            if (baseType != null && baseType.IsClass)
            {
                var parentProperties = GetBaseProps(baseType, ignoreProperties).ToArray();
                var parentPropertyNames = parentProperties.Select(p => p.Name).ToArray();
                var childProperties = GetProps(classType, ignoreProperties)
                    .Where(p => !parentPropertyNames.Contains(p.Name)).ToArray();

                return parentProperties.Union(childProperties);
            }

            return GetProps(classType, ignoreProperties);
        }

        /// <summary>
        /// Creates a Table from an enumerable
        /// </summary>
        [Obsolete("For legacy project use only. Please create a report model by inheriting from Worksheet or Table instead.")]
        public static Table ToTable<T>(this ICollection<T> data, bool includeHeaders,
            TableHeaderStyle headerStyle = null, params string[] ignoreProperties)
        {
            var table = new Table
            {
                IncludeHeaders = includeHeaders,
                HeaderStyle = headerStyle ?? new TableHeaderStyle()
            };
            
            var properties = typeof(T).GetBaseProps(ignoreProperties).ToArray();
            
            foreach (var prop in properties)
            {
                var hideColumn = prop.GetCustomAttributes<IgnoreDataMemberAttribute>().Any()
                                || prop.GetCustomAttributes<IdColumnAttribute>().Any();
                
                if (hideColumn) { continue; }

                var hideHeader = prop.GetCustomAttributes<HideHeaderAttribute>(false).Any();
                var displaySubtotal = prop.GetCustomAttributes<SubtotalHeaderAttribute>(false).Any() 
                                   && prop.PropertyType.IsNumeric();
                var backgroundColour = prop.GetCustomAttributes<CellBackgroundColourAttribute>(false).FirstOrDefault();
                var textColour = prop.GetCustomAttributes<CellTextColourAttribute>(false).FirstOrDefault();
                var format = prop.GetCustomAttributes<CellFormatAttribute>(false).FirstOrDefault();
                
                var column = new TableColumn
                {
                    Name = hideHeader ? string.Empty : prop.GetDataAnnotationDisplayName(),
                    DisplaySubtotal = displaySubtotal,
                    CellBackgroundFormat = backgroundColour,
                    CellTextFormat = textColour
                };

                if (format != default(CellFormatAttribute))
                {
                    column.Format = format.CellFormat;
                }

                table.Columns.Add(column);
            }

            foreach (var item in data)
            {
                var row = new TableRow();
                var values = properties.Select(p => p.GetValue(item, null)).ToList();

                for (var index = 0; index < values.Count; index++)
                {
                    var column = table.Columns[index];
                    var cellValue = values[index];

                    var cell = new TableCell
                    {
                        Value = cellValue ?? string.Empty,
                        FillColour = column.CellBackgroundFormat?.Colour,
                        TextColour = column.CellTextFormat?.Colour
                    };

                    // Get FillColour from property
                    if (!string.IsNullOrEmpty(column.CellBackgroundFormat?.ReferenceField))
                    {
                        var valueOfMatchingProperty = typeof(T)
                            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .Where(prop => string.Equals(prop.Name, column.CellBackgroundFormat.ReferenceField))
                            .Select(prop => prop.GetValue(item, null) as string)
                            .FirstOrDefault();

                        cell.FillColour = valueOfMatchingProperty;
                    }
                    
                    // Get TextColour from property
                    if (!string.IsNullOrEmpty(column.CellTextFormat?.ReferenceField))
                    {
                        var valueOfMatchingProperty = typeof(T)
                            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .Where(prop => string.Equals(prop.Name, column.CellTextFormat.ReferenceField))
                            .Select(prop => prop.GetValue(item, null) as string)
                            .FirstOrDefault();

                        cell.TextColour = valueOfMatchingProperty;
                    }


                    row.Cells.Add(cell);
                }

                table.Rows.Add(row);
            }

            return table;
        }

        /// <summary>
        /// Creates a Worksheet from an enumerable
        /// </summary>
        [Obsolete("For legacy project use only. Please create a report model by inheriting from Worksheet or Table instead.")]
        public static Worksheet ToWorksheet<T>(this ICollection<T> data, 
            string sheetName = null, 
            bool includeHeaders = true,
            TableHeaderStyle headerStyle = null,
            params string[] ignoreProperties)
        {
            var table = data.ToTable(includeHeaders, headerStyle, ignoreProperties);

            var freezePane = default(FreezePane);
            if (includeHeaders)
            {
                freezePane = new FreezePane();
                if (table.Columns.Any(c => c.DisplaySubtotal))
                {
                    freezePane.StartingCell = "A3";
                    freezePane.FrozenRows = 2;
                }
            }

            return new Worksheet
            {
                SheetName = sheetName,
                FreezePane = freezePane,
                Tables = new List<Table> { table }
            };
        }
    }
}
