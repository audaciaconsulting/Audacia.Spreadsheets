using System.Collections.Generic;
using System.IO;
using System.Linq;
using Audacia.Core.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace Audacia.Spreadsheets
{
    public class Spreadsheet
    {
        public IList<Worksheet> Worksheets { get; } = new List<Worksheet>();

        public static Spreadsheet FromWorksheets(IEnumerable<Worksheet> worksheets)
        {
            var spreadsheet = new Spreadsheet();
            foreach (var worksheet in worksheets)
            {
                spreadsheet.Worksheets.Add(worksheet);
            }

            return spreadsheet;
        }

        public byte[] Export()
        {
            using (var stream = new MemoryStream())
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var cellFormats = new List<CellStyle>();
                var workbookPart = document.AddWorkbookPart();
                var workbook = workbookPart.Workbook = new Workbook();
                var sheets = workbook.AppendChild(new Sheets());
                var definedNames = workbook.AppendChild(new DefinedNames());
                workbook.CalculationProperties = new CalculationProperties();

                // Shared string table
                var sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringTablePart.SharedStringTable = new SharedStringTable();
                sharedStringTablePart.SharedStringTable.Save();

                // Stylesheet
                var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = new StylesheetBuilder(Worksheets)
                    .GetDefaultStyles(out var fillColours, out var textColours, out var fonts);
                workbookStylesPart.Stylesheet.Save();

                foreach (var worksheet in Worksheets)
                {
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                    var sheetName = !string.IsNullOrWhiteSpace(worksheet.SheetName)
                        ? worksheet.SheetName
                        : Worksheets.IndexOf(worksheet).ToString();

                    var sheetId = sheets.Elements<Sheet>().Any()
                        ? sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1
                        : 1;

                    var sheet = new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = sheetId,
                        Name = sheetName.Truncate(30, string.Empty),
                        State = SheetStateValues.Visible
                    };
                    sheets.Append(sheet);

                    var writer = OpenXmlWriter.Create(worksheetPart);

                    foreach (var table in worksheet.Tables)
                    {
                        writer.WriteStartElement(new OpenXmlWorksheet());

                        SpreadsheetBuilderHelper.AddSheetView(writer, table.FreezeTopRows);
                        SpreadsheetBuilderHelper.AddColumns(writer, table);

                        writer.WriteStartElement(new SheetData());

                        SpreadsheetBuilderHelper.Insert(table, workbookStylesPart.Stylesheet, cellFormats, fillColours,
                            textColours, fonts, worksheetPart, writer);

                        writer.WriteEndElement(); // Sheet Data

                        // Auto Filter all data on worksheet
                        if (SpreadsheetBuilderHelper.TryGetAutoFilter(sheetName, table, definedNames, out var filter))
                        {
                            writer.WriteElement(filter);
                        }

                        writer.WriteEndElement(); // Worksheet
                    }

                    writer.Close();

                    if (worksheet.WorksheetProtection != null)
                    {
                        SpreadsheetBuilderHelper.AddProtection(worksheetPart, worksheet.WorksheetProtection);
                    }
                }

                document.Close();
                return stream.ToArray();
            }
        }
    }
}
