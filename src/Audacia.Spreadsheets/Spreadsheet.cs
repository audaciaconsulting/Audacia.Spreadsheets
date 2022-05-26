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
        public List<WorksheetBase> Worksheets { get; } = new List<WorksheetBase>();

        /// <summary>
        /// Must be defined after worksheets have been defined or the ranges will be moved by the addition of tables
        /// </summary>
        public List<NamedRangeModel> NamedRanges { get; } = new List<NamedRangeModel>();
        
        /// <summary>
        /// Writes the spreadsheet to a stream as an Excel Workbook (*.xlsx).
        /// </summary>
        public void Write(Stream stream)
        {
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var tables = Worksheets.GetTables();
                var sharedData = new StylesheetBuilder(tables).Build();
                
                var workbookPart = document.AddWorkbookPart();
                var workbook = workbookPart.Workbook = new Workbook();
                var sheets = workbook.AppendChild(new Sheets());
                
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
                    using (var writer = OpenXmlWriter.Create(worksheetPart))
                    {
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
                            State = worksheet.Visibility,
                            Name = worksheet.SheetName
                        };
                        
                        sheets.Append(sheet);

                        worksheet.Write(sharedData, writer); 
                        
                        // Close the openxml writer for this worksheet part
                        writer.Close();
                    }
                }
                var definedNames = new DefinedNames();

                if (NamedRanges != null && NamedRanges.Any())
                {
                    //  Adds DefinedNames To Workbook
                    foreach (var namedRange in NamedRanges)
                    {
                        var definedNameToPush = new DefinedName
                        {
                            Name = namedRange.Name,
                            Text = $"\'{namedRange.SheetName}\'!{namedRange.StartCell}:{namedRange.EndCell}"
                        };
                        definedNames.Append(definedNameToPush);
                    }
                    workbookPart.Workbook.DefinedNames = definedNames;
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

        /// <summary>
        /// Writes the spreadsheet to the specified filepath as an Excel Workbook (*.xlsx).
        /// </summary>
        public void Export(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentNullException(filePath);

            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                throw new DirectoryNotFoundException(directory);

            using (var fileStream = new FileStream(filePath, FileMode.OpenOrCreate))
            {
                Write(fileStream);
                fileStream.Close();
            }
        }

        /// <summary>
        /// Reads the spreadsheet from the provided file bytes, supports Excel Workbook (*.xlsx).
        /// </summary>
        /// <param name="bytes">Spreadsheet file bytes</param>
        /// <param name="includeHeaders">Declare if column header row is included on the spreadsheet</param>
        /// <param name="hasSubtotals">Declare if subtotal row exists above column header row</param>
        /// <param name="ignoreSheets">Declare if sheets should be skipped, filters by name</param>
        public static Spreadsheet FromBytes(byte[] bytes, bool includeHeaders = true, bool hasSubtotals = false, IEnumerable<string> ignoreSheets = default)
        {
            using (var ms = new MemoryStream(bytes))
            {
                return FromStream(ms, includeHeaders, hasSubtotals, ignoreSheets);
            }
        }

        /// <summary>
        /// Reads the spreadsheet from the provided file location, supports Excel Workbook (*.xlsx).
        /// </summary>
        /// <param name="filePath">Path to spreadsheet file</param>
        /// <param name="includeHeaders">Declare if column header row is included on the spreadsheet</param>
        /// <param name="hasSubtotals">Declare if subtotal row exists above column header row</param>
        /// <param name="ignoreSheets">Declare if sheets should be skipped, filters by name</param>
        public static Spreadsheet FromFilePath(string filePath, bool includeHeaders = true, bool hasSubtotals = false, IEnumerable<string> ignoreSheets = default)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                return FromStream(fs, includeHeaders, hasSubtotals, ignoreSheets);
            }
        }

        /// <summary>
        /// Creates a spreadsheet containing the provided worksheets.
        /// </summary>
        /// <param name="worksheets">Worksheets to include</param> 
        public static Spreadsheet FromWorksheets(params WorksheetBase[] worksheets)
        {
            var spreadsheet = new Spreadsheet();
            if (worksheets != null)
            {
                spreadsheet.Worksheets.AddRange(worksheets);
            }
            
            return spreadsheet;
        }

        /// <summary>
        /// Reads the spreadsheet from the provided stream, supports Excel Workbook (*.xlsx).
        /// </summary>
        /// <param name="bytes">Spreadsheet file bytes</param>
        /// <param name="includeHeaders">Declare if column header row is included on the spreadsheet</param>
        /// <param name="hasSubtotals">Declare if subtotal row exists above column header row</param>
        /// <param name="ignoreSheets">Declare if sheets should be skipped, filters by name</param>
        public static Spreadsheet FromStream(Stream stream, bool includeHeaders = true, bool hasSubtotals = false, IEnumerable<string> ignoreSheets = default)
        {
            using (var spreadSheet = SpreadsheetDocument.Open(stream, false))
            {
                var descendants = spreadSheet.WorkbookPart.Workbook.Descendants<Sheet>();
                if (ignoreSheets != null && ignoreSheets.Any())
                {
                    descendants = descendants.Where(d => !ignoreSheets.Contains(d.Name?.ToString()));
                }
                
                var worksheets = descendants
                    .Select(sheet => Worksheet.FromOpenXml(sheet, spreadSheet, includeHeaders, hasSubtotals))
                    .ToArray();

                return FromWorksheets(worksheets);
            }
        }
    }
}
