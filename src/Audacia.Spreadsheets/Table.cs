﻿using System;
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

        public List<TableColumn> Columns { get; } = new List<TableColumn>();

        public List<TableRow> Rows { get; } = new List<TableRow>();

        public void Write(SharedDataTable sharedData, OpenXmlWriter writer)
        {
            writer.WriteStartElement(new SheetData());
            
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
                    column.Write(HeaderStyle, headerCellRef, isFirstColumn, isLastColumn, sharedData, writer);
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
            
            writer.WriteEndElement(); // Sheet Data
        }


        public static int GetMaxCharacterWidth(Table table, int columnIndex)
        {
            var column = table.Columns[columnIndex];

            if (column.Width.HasValue)
            {
                return column.Width.Value;
            }

            var cells = table.Rows.Select(r => r.Cells.Count > columnIndex ? r.Cells[columnIndex] : null).Where(c => c != null).ToList();

            if (table.IncludeHeaders)
            {
                cells.Add(new TableCell(column.Name));
            }

            // Create Cells for Rollups
            if (column.DisplaySubtotal)
            {
                var total = table.Rows
                        .Where(r => r.Cells.Count > columnIndex)
                        .Select(r =>
                        {
                            var value = r.Cells.ElementAt(columnIndex).Value;
                            var isNumeric = value.GetType().IsNumeric();
                            return isNumeric ? Convert.ToDecimal(value) : 0;
                        })
                        .DefaultIfEmpty(0)
                        .Sum(v => v);
                cells.Add(new TableCell
                {
                    // Format as currency because the number value alone just isn't long enough
                    Value = $"{total:C}"
                });
            }

            // Find the max cell width of supplied column           
            var current = 0;
            for (var i = 0; i < cells.Count; i++)
            {
                var cell = cells[i];
                var cellValue = cell.Value?.ToString() ?? string.Empty;
                var cellTextLength = cellValue.Length;

                if (cellTextLength > current)
                {
                    current = cellTextLength;
                }
            }

            //  75 is chosen as a maximum length to prevent the column becoming too monsterous
            if (current > 75)
            {
                current = 75;
            }

            return current;
        }
    }
}