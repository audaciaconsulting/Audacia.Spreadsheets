using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace Audacia.Spreadsheets
{
    public class MultiTableWorksheet : WorksheetBase
    {
        public IList<Table> Tables { get; set; }
        
        protected override void WriteSheetData(SharedDataTable sharedData, OpenXmlWriter writer)
        {
            CellReference prevCellTableEnd = null;
            foreach (var table in Tables)
            {
                // Move the next table down by the size of the current table
                if (prevCellTableEnd != null)
                {
                    // Add a row inbetween to separate the tables
                    writer.WriteStartElement(new Row());
                    writer.WriteEndElement();
                    
                    // Set the next table to start 1 cell below the current table
                    prevCellTableEnd.NextRow();
                    table.StartingCellRef = prevCellTableEnd;
                }

                // Write the table data and return the last cell ref for the next table
                prevCellTableEnd = table.Write(sharedData, writer);
            }
        }
    }
}