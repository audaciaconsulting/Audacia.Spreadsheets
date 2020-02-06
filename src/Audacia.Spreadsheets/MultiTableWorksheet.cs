using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace Audacia.Spreadsheets
{
    public class MultiTableWorksheet : WorksheetBase
    {
        public IEnumerable<Table> Tables { get; set; }
        
        public override void WriteSheetContent(SharedDataTable sharedData, OpenXmlWriter writer)
        {
            AddSheetView(writer);

            foreach (var table in Tables)
            {
                AddColumns(table, writer);
            
                table.Write(sharedData, writer);
            }
            
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