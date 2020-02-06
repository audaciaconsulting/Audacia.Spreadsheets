using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace Audacia.Spreadsheets
{
    public class MultiTableWorksheet : WorksheetBase
    {
        public IReadOnlyCollection<Table> Tables { get; set; }
        
        protected override void WriteSheet(SharedDataTable sharedData, OpenXmlWriter writer)
        {
            // Sheet view, Columns, and Sheet Data should only ever be written once per worksheet
            AddSheetView(writer);
            AddColumns(Tables, writer);

            writer.WriteStartElement(new SheetData());

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
            
            // Close SheetData tag
            writer.WriteEndElement();
            
            // We don't currently support autofilters for multi-table worksheets, but if we did it would go here

            DataValidations dataValidations = new DataValidations();
            
            // Add Static Data Validation
            if (StaticDataValidations != null && StaticDataValidations.Any())
            {
                foreach (var val in StaticDataValidations)
                {
                    val.Write(dataValidations);
                }
            }
            
            // Add Dynamic Data Validation
            if (DependentDataValidations != null && DependentDataValidations.Any())
            {
                foreach (var val in DependentDataValidations)
                {
                    val.Write(dataValidations);
                }
            }

            //  Only add validation if dataValidations has Descendants
            if (dataValidations.Descendants<DataValidation>().Any())
            {
                writer.WriteElement(dataValidations);
            }
        }
    }
}