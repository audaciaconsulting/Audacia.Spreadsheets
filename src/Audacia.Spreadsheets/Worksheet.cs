using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace Audacia.Spreadsheets
{
    public class Worksheet : WorksheetBase
    {
        public Table Table { get; set; }
        
        public override void WriteSheetContent(SharedDataTable sharedData, OpenXmlWriter writer)
        {
            AddSheetView(writer);
            AddColumns(Table, writer);
            
            Table.Write(sharedData, writer);
            
            if (HasAutofilter)
            {
                AddAutoFilter(Table, sharedData.DefinedNames, writer);
            }
            
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
