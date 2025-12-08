using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class MultiTableWorksheet : WorksheetBase
    {
        public bool IncludeGapBetweenTables { get; set; } = true;

        public IList<Table> Tables { get; set; } = new List<Table>();

        public IEnumerable<string> MergeCells { get; set; } = new List<string>();

        protected override void WriteSheetData(SharedDataTable sharedData, OpenXmlWriter writer)
        {
            CellReference? prevCellTableEnd = null;
            foreach (var table in Tables)
            {
                // Extract the logic for handling the gap between tables
                if (prevCellTableEnd != null)
                {
                    HandleGapBetweenTables(writer, prevCellTableEnd);
                    table.StartingCellRef = prevCellTableEnd;
                }

                // Write the table data and return the last cell ref for the next table
                prevCellTableEnd = table.Write(sharedData, writer);
            }
        }

        private void HandleGapBetweenTables(OpenXmlWriter writer, CellReference prevCellTableEnd)
        {
            if (IncludeGapBetweenTables)
            {
                // Add a row inbetween to separate the tables
                var newRow = new Row();
                writer.WriteStartElement(newRow);
                writer.WriteEndElement();

                // Set the next table to start 1 cell below the current table
                prevCellTableEnd.NextRow();
            }
        }
    }
}