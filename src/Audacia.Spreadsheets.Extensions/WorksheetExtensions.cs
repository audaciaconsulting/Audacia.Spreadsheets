using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using Audacia.Core.Extensions;
using Audacia.Spreadsheets.Models.Attributes;
using Audacia.Spreadsheets.Models.WorksheetData;

namespace Audacia.Spreadsheets.Extensions
{
    public static class WorksheetExtensions
    {
        private static IEnumerable<PropertyInfo> GetPublicProperties(this Type classType, params string[] ignoreProperties)
        {
            return classType
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Select(prop => new
                {
                    Property = prop,
                    Type = prop.PropertyType,
                    UnderlyingType = prop.PropertyType.GetUnderlyingTypeIfNullable()
                })
                .Where(item =>
                {
                    if (ignoreProperties.Contains(item.Property.Name))
                    {
                        return false;
                    }

                    var ignoreDataMemberAttribute = item.Property.GetCustomAttributes<IgnoreDataMemberAttribute>(false);

                    if (ignoreDataMemberAttribute.Any())
                    {
                        return false;
                    }
                    
                    // Ignore any properties which are not value types or strings
                    if (!item.UnderlyingType.IsValueType && item.UnderlyingType != typeof(string))
                    {
                        return false;
                    }

                    return true;
                })
                .Select(a => a.Property)
                .ToArray();
        }
        
        /// <summary>
        /// Creates a Data Table from an enumerable
        /// </summary>
        public static TableWrapperModel ToDataTableModel<T>(this IEnumerable<T> data, params string[] ignoreProperties)
        {
            if (ignoreProperties == null)
            {
                ignoreProperties = new string[0];
            }

            var dataTable = new TableWrapperModel();

            var propertiesWithTypes = typeof(T).GetPublicProperties(ignoreProperties).ToList();

            var columns = new List<TableColumnModel>();

            foreach (var item in propertiesWithTypes)
            {
                var column = new TableColumnModel
                {
                    Name = item.GetDataAnnotationDisplayName(),
                    IsIdColumn = ((IdColumnAttribute[])item
                            .GetCustomAttributes(typeof(IdColumnAttribute), false))
                            .FirstOrDefault() != null,

                    CellBackgroundFormat = ((CellBackgroundColourAttribute[])item
                        .GetCustomAttributes(typeof(CellBackgroundColourAttribute), false))
                        .FirstOrDefault(),

                    CellTextFormat = ((CellTextColourAttribute[])item
                        .GetCustomAttributes(typeof(CellTextColourAttribute), false))
                        .FirstOrDefault()
                };

                var cellFormatAttribute = (CellFormatAttribute[])item
                    .GetCustomAttributes(typeof(CellFormatAttribute), false);

                if (cellFormatAttribute.Any())
                {
                    column.Format = cellFormatAttribute.First().CellFormatType;
                }

                var hideHeaderAttribute =
                    (HideHeaderAttribute[])item.GetCustomAttributes(typeof(HideHeaderAttribute), false);

                if (hideHeaderAttribute.Any())
                {
                    column.HideHeader = true;
                }

                columns.Add(column);
            }

            dataTable.Columns = columns.Where(c => !c.IsIdColumn);

            var rows = new List<TableRowModel>();

            foreach (var entity in data)
            {
                int? id = null;
                var values = propertiesWithTypes.Select(
                    item => item.GetValue(entity, null)).ToList();

                var cells = new List<TableCellModel>();

                var index = 0;
                foreach (var v in values)
                {
                    var cell = new TableCellModel()
                    {
                        Value = v ?? string.Empty
                    };

                    if (columns.ElementAt(index).CellBackgroundFormat != null)
                    {
                        if (!string.IsNullOrWhiteSpace(columns.ElementAt(index).CellBackgroundFormat.Colour))
                        {
                            cell.FillColour = columns.ElementAt(index).CellBackgroundFormat.Colour;
                        }
                        else if (!string.IsNullOrWhiteSpace(columns.ElementAt(index).CellBackgroundFormat.ReferenceField))
                        {
                            var property =
                                typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance).SingleOrDefault(
                                    prop => prop.Name == columns.ElementAt(index).CellBackgroundFormat.ReferenceField);

                            var propValue = property?.GetValue(entity, null) as string;
                            cell.FillColour = propValue;
                        }
                    }

                    if (columns.ElementAt(index).CellTextFormat != null)
                    {
                        if (!string.IsNullOrWhiteSpace(columns.ElementAt(index).CellTextFormat.Colour))
                        {
                            cell.TextColour = columns.ElementAt(index).CellTextFormat.Colour;
                        }
                        else if (!string.IsNullOrWhiteSpace(columns.ElementAt(index).CellTextFormat.ReferenceField))
                        {
                            var property =
                                typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance).SingleOrDefault(
                                    prop => prop.Name == columns.ElementAt(index).CellTextFormat.ReferenceField);

                            var propValue = property?.GetValue(entity, null) as string;
                            cell.TextColour = propValue;
                        }
                    }

                    // Don't add ID column, just add it to our model
                    if (columns.ElementAt(index).IsIdColumn)
                    {
                        id = v as int?;
                    }
                    else
                    {
                        cells.Add(cell);
                    }

                    index++;
                }

                var row = new TableRowModel()
                {
                    Id = id,
                    Cells = cells
                };

                rows.Add(row);
            }

            dataTable.Rows = rows;

            return dataTable;
        }
        
        /// <summary>
        /// Creates a Worksheet from an enumerable
        /// </summary>
        public static WorksheetModel ToWorkSheet<T>(this IEnumerable<T> data, int sheetIndex, string sheetName, bool includeHeaders,
            SpreadsheetHeaderStyle headerStyle = null, params string[] ignoreProperties)
        {
            return new WorksheetModel
            {
                SheetIndex = sheetIndex,
                SheetName = sheetName,
                Tables = new List<TableModel>
                {
                    new TableModel
                    {
                        HeaderStyle = headerStyle,
                        IncludeHeaders = includeHeaders,
                        Data = data.ToDataTableModel(ignoreProperties)
                    }
                }
            };
        }
    }
}
