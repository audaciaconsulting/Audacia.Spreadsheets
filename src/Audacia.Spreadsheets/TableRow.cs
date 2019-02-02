using System.Collections.Generic;
using Audacia.Spreadsheets.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class TableRow
    {
        public int? Id { get; set; }
        public IList<TableCell> Cells { get; } = new List<TableCell>();

        public void Write(CellReference cellReference, IList<TableColumn> columns, SharedDataTable sharedData, OpenXmlWriter writer)
        {
            writer.WriteStartElement(new Row());
                
            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var column = columns[columnIndex];
                var cell = Cells[columnIndex];
                var value = cell.Value;

                var cellStyle = new CellStyle
                {
                    TextColour = 0U,
                    BackgroundColour = 0U,
                    BorderBottom = true,
                    BorderTop = true,
                    BorderLeft = true,
                    BorderRight = true,
                    Format = column.Format,
                    HasWordWrap = value is string && !cell.IsFormula
                };

                if (!string.IsNullOrWhiteSpace(cell.FillColour))
                {
                    cellStyle.BackgroundColour = sharedData.FillColours[cell.FillColour];
                }

                if (!string.IsNullOrWhiteSpace(cell.TextColour))
                {
                    cellStyle.TextColour = sharedData.TextColours[cell.TextColour];
                }

                var styleIndex = sharedData.GetOrCreateCellFormat(cellStyle).Index;

                cell.Write(styleIndex, cellReference, writer);

                cellReference.NextColumn();
            }

            writer.WriteEndElement();
        }

        public static TableRow FromCells(IEnumerable<TableCell> cells, int? id)
        {
            var row = new TableRow
            {
                Id = id
            };
            
            foreach (var cell in cells)
            {
                row.Cells.Add(cell);
            }
            
            return row;
        }
    }
}
