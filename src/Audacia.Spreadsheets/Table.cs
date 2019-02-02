using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class Table
    {
        public string StartingCellRef { get; set; } = "A1";
        
        public TableHeaderStyle HeaderStyle { get; set; }

        public bool IncludeHeaders { get; set; }
        
        public IList<WorksheetTableColumn> Columns { get; } = new List<WorksheetTableColumn>();

        public IList<WorksheetTableRow> Rows { get; } = new List<WorksheetTableRow>();
        
        public void Write(SharedData sharedData, OpenXmlWriter writer)
        {
            var stylesheet = sharedData.Stylesheet;
            var cellFormats = sharedData.CellFormats;
            var fillColours = sharedData.FillColours;
            var textColours = sharedData.TextColours;
            var fonts = sharedData.Fonts;

            var cellReference = new CellReference(StartingCellRef);
            var cellReferenceRowIndex = StartingCellRef.GetReferenceRowIndex();
            var cellReferenceColumnIndex = StartingCellRef.GetReferenceColumnIndex();

            // Write Subtotals above headers
            if (IncludeHeaders && Columns.Any(c => c.DisplaySubtotal))
            {
                var subtotalCellRef = cellReference.Clone();
                writer.WriteStartElement(new Row());

                foreach (var column in Columns)
                {
                    var isFirstColumn = column == Columns.ElementAt(0);
                    var isLastColumn = column == Columns.ElementAt(Columns.Count - 1);
                    column.WriteSubtotal(subtotalCellRef, isFirstColumn, isLastColumn, Rows.Count, sharedData, writer);
                    subtotalCellRef.NextColumn();
                }

                writer.WriteEndElement();
                cellReference.NextRow();
            }

            // Write headers above data
            if (IncludeHeaders)
            {
                var headerCellRef = cellReference.Clone();
                writer.WriteStartElement(new Row());

                foreach (var column in Columns)
                {
                    var isFirstColumn = column == Columns.ElementAt(0);
                    var isLastColumn = column == Columns.ElementAt(Columns.Count - 1);
                    column.Write(HeaderStyle, headerCellRef, isFirstColumn,isLastColumn, sharedData, writer);
                    headerCellRef.NextColumn();
                }

                writer.WriteEndElement();
                cellReference.NextRow();
            }

            // Write data
            foreach (var row in Rows)
            {
                writer.WriteStartElement(new Row());
                var columnIndex = 0;
                foreach (var column in Columns)
                {
                    var cellModel = row.Cells.ElementAt(columnIndex);
                    var value = cellModel.Value;

                    var cellStyle = new CellStyle
                    {
                        TextColour = 0U,
                        BackgroundColour = 0U,
                        BorderBottom = true,
                        BorderTop = true,
                        BorderLeft = true,
                        BorderRight = true,
                        Format = value is DateTime ? CellFormatType.Date : column.Format,
                        HasWordWrap = value is string
                    };

                    if (!string.IsNullOrWhiteSpace(cellModel.FillColour))
                    {
                        cellStyle.BackgroundColour = fillColours[cellModel.FillColour];
                    }

                    if (!string.IsNullOrWhiteSpace(cellModel.TextColour))
                    {
                        cellStyle.TextColour = textColours[cellModel.TextColour];
                    }

                    var styleIndex = GetOrCreateCellFormat(cellStyle, cellFormats, stylesheet).Index;

                    var dataTypeAndValue = GetDataTypeAndFormattedValue(value);

                    WriteCell(writer, styleIndex, $"{cellReferenceColumnIndex}{cellReferenceRowIndex}",
                        dataTypeAndValue.Item1, dataTypeAndValue.Item2, cellModel.IsFormula);

                    cellReferenceColumnIndex = (cellReferenceColumnIndex.GetColumnNumber() + 1)
                        .GetExcelColumnName();

                    columnIndex++;
                }
                cellReferenceColumnIndex = StartingCellRef.GetReferenceColumnIndex();
                cellReferenceRowIndex++;
                writer.WriteEndElement();
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
            var row = WorksheetTableRow.FromCells(rowCells, 0);
            
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
                            // TODO JP: do this properly later
                            var isNumeric = value.ToString().IsNumeric();
                            return (decimal)(isNumeric ? value : 0);
                        })
                        .DefaultIfEmpty(0)
                        .Sum(v => v))
                    .Select(value => new TableCell
                    {
                        // Format as currency because the number value alone just isn't long enough
                        Value = $"{value:C}"
                    });
                
                var rollupRow = WorksheetTableRow.FromCells(rollupCells, 0);
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
