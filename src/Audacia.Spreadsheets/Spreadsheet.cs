using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class Spreadsheet
    {
        public IList<Worksheet> Worksheets { get; } = new List<Worksheet>();

        /// <summary>
        /// Writes the spreadsheet to a stream as an Excel Workbook (*.xlsx).
        /// </summary>
        public void Write(Stream stream)
        {
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var sharedData = new StylesheetBuilder(Worksheets).Build();
                
                var workbookPart = document.AddWorkbookPart();
                var workbook = workbookPart.Workbook = new Workbook();
                var sheets = workbook.AppendChild(new Sheets());
                workbook.AppendChild(sharedData.DefinedNames);
                workbook.CalculationProperties = new CalculationProperties();

                // Shared string table
                var sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringTablePart.SharedStringTable = new SharedStringTable();
                sharedStringTablePart.SharedStringTable.Save();

                // Stylesheet
                var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = sharedData.Stylesheet;
                workbookStylesPart.Stylesheet.Save();

                for (var index = 0; index < Worksheets.Count; index++)
                {
                    var sheetNumber = index + 1;
                    var worksheet = Worksheets[index];
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                    // Sanitize worksheet name
                    const int maxSheetNameLength = 30;
                    if (string.IsNullOrWhiteSpace(worksheet.SheetName))
                    {
                        worksheet.SheetName = $"Sheet {sheetNumber}";
                    }
                    else if (worksheet.SheetName.Length > maxSheetNameLength)
                    {
                        worksheet.SheetName = worksheet.SheetName.Substring(0, maxSheetNameLength).Trim();
                    }

                    var sheet = new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = Convert.ToUInt32(sheetNumber),
                        State = SheetStateValues.Visible,
                        Name = worksheet.SheetName
                    };
                    
                    sheets.Append(sheet);
                    worksheet.Write(worksheetPart, sharedData);
                }

                document.Close();
            }
        }

        /// <summary>
        /// Writes the spreadsheet to a byte array as an Excel Workbook (*.xlsx).
        /// </summary>
        public byte[] Export()
        {
            using (var stream = new MemoryStream())
            {
                Write(stream);
                return stream.ToArray();
            }
        }
        
        public static Spreadsheet FromWorksheets(params Worksheet[] worksheets)
        {
            var spreadsheet = new Spreadsheet();
            if (worksheets != null)
            {
                spreadsheet.Worksheets.AddRange(worksheets);
            }
            
            return spreadsheet;
        }

        public static Spreadsheet FromStream(Stream stream, bool includeHeaders = true)
        {
            using (var spreadSheet = SpreadsheetDocument.Open(stream, false))
            {
                var worksheets = spreadSheet.WorkbookPart.Workbook.Descendants<Sheet>()
                    .Select(sheet => Worksheet.FromOpenXml(sheet, spreadSheet, includeHeaders))
                    .ToArray();

                return FromWorksheets(worksheets);
            }
        }
    }
}
