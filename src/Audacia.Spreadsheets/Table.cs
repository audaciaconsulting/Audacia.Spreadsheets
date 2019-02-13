using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Core.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class Table
    {
        public Table() { }

        public Table(bool includeHeaders) => IncludeHeaders = includeHeaders;
        
        public string StartingCellRef { get; set; } = "A1";
        
        public TableHeaderStyle HeaderStyle { get; set; }

        public bool IncludeHeaders { get; set; }
        
        public IList<TableColumn> Columns { get; } = new List<TableColumn>();

        public IList<TableRow> Rows { get; } = new List<TableRow>();
        
        public void Write(SharedDataTable sharedData, OpenXmlWriter writer)
        {
            var rowReference = new CellReference(StartingCellRef);

            // Write Subtotals above headers
            if (IncludeHeaders && Columns.Any(c => c.DisplaySubtotal))
            {
                var subtotalCellRef = rowReference.Clone();
                writer.WriteStartElement(new Row());

                foreach (var column in Columns)
                {
                    var isFirstColumn = column == Columns.ElementAt(0);
                    var isLastColumn = column == Columns.ElementAt(Columns.Count - 1);
                    column.WriteSubtotal(subtotalCellRef, isFirstColumn, isLastColumn, Rows.Count, sharedData, writer);
                    subtotalCellRef.NextColumn();
                }

                writer.WriteEndElement();
                rowReference.NextRow();
            }

            // Write headers above data
            if (IncludeHeaders)
            {
                var headerCellRef = rowReference.Clone();
                writer.WriteStartElement(new Row());

                foreach (var column in Columns)
                {
                    var isFirstColumn = column == Columns.ElementAt(0);
                    var isLastColumn = column == Columns.ElementAt(Columns.Count - 1);
                    column.Write(HeaderStyle, headerCellRef, isFirstColumn,isLastColumn, sharedData, writer);
                    headerCellRef.NextColumn();
                }

                writer.WriteEndElement();
                rowReference.NextRow();
            }

            // Write data
            foreach (var row in Rows)
            {
                row.Write(rowReference.Clone(), Columns, sharedData, writer);
                rowReference.NextRow();
            }
        }

        
        public static Dictionary<int, int> GetMaxCharacterWidth(Table model)
        {
            //iterate over all cells getting a max char value for each column
            var maxColWidth = new Dictionary<int, int>();

            // Create Cells for Data
            var columnHeaderWithData = model.Rows.ToList();

            // Create Cells for Headers
            var rowCells = model.Columns.Select(c => new TableCell(c.Name));
            var row = TableRow.FromCells(rowCells, 0);
            
            columnHeaderWithData.Add(row);
            
            // Create Cells for Rollups
            if (model.Columns.Any(c => c.DisplaySubtotal))
            {
                var rollupCells = model.Columns
                    .Select((col, index) => new {col, index})
                    .Where(t => t.col.DisplaySubtotal)
                    .Select(t => model.Rows
                        .Where(r => r.Cells.Count > t.index)
                        .Select(r =>
                        {
                            var value = r.Cells.ElementAt(t.index).Value;
                            var isNumeric = value.GetType().IsNumeric();
                            return isNumeric ? Convert.ToDecimal(value) : 0;
                        })
                        .DefaultIfEmpty(0)
                        .Sum(v => v))
                    .Select(value => new TableCell
                    {
                        // Format as currency because the number value alone just isn't long enough
                        Value = $"{value:C}"
                    });
                
                var rollupRow = TableRow.FromCells(rollupCells, 0);
                columnHeaderWithData.Add(rollupRow);
            }

            // Find the max cell width of each column
            foreach (var r in columnHeaderWithData)
            {
                var cells = r.Cells.ToArray();

                //using cell index as my column
                for (var i = 0; i < cells.Length; i++)
                {
                    var cell = cells[i];
                    var cellValue = cell.Value?.ToString() ?? string.Empty;
                    var cellTextLength = cellValue.Length;

                    if (maxColWidth.ContainsKey(i))
                    {
                        var current = maxColWidth[i];
                        if (cellTextLength > current)
                        {
                            maxColWidth[i] = cellTextLength;
                        }
                    }
                    else
                    {
                        maxColWidth.Add(i, cellTextLength);
                    }
                }
            }

            return maxColWidth;
        }
    }
}
