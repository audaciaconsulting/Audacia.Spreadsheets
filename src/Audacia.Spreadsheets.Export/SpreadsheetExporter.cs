// ReSharper disable PossiblyMistakenUseOfParamsMethod
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Audacia.Spreadsheets.Models.WorksheetData;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets.Export
{
    public class SpreadsheetExporter : ISpreadsheetExporter
    {
        public byte[] ExportSpreadsheetBytes(SpreadsheetModel model)
        {
            using (var stream = new MemoryStream())
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var cellFormats = new List<SpreadsheetCellStyle>();

                var allTables = model.Worksheets.SelectMany(w => w.Tables).ToList();

                var distinctHeaderStyles =
                    allTables
                        .Where(t => t.HeaderStyle != null)
                        .Select(t => t.HeaderStyle)
                        .Distinct();

                var distinctBackgroundColours =
                    allTables
                    .SelectMany(dt => dt.Data.Rows)
                    .SelectMany(r => r.Cells.Select(c => c.FillColour)).Where(c => !string.IsNullOrWhiteSpace(c))
                    .Distinct();

                var distinctTextColours =
                    allTables
                    .SelectMany(dt => dt.Data.Rows)
                    .SelectMany(r => r.Cells.Select(c => c.TextColour)).Where(c => !string.IsNullOrWhiteSpace(c))
                    .Distinct();

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
                workbookStylesPart.Stylesheet = GetDefaultStyles(distinctBackgroundColours, distinctTextColours, distinctHeaderStyles,
                    out var fillColours, out var textColours, out var fonts);
                workbookStylesPart.Stylesheet.Save();

                foreach (var worksheetModel in model.Worksheets.OrderBy(w => w.SheetIndex))
                {
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                    var sheetName = !string.IsNullOrWhiteSpace(worksheetModel.SheetName)
                        ? worksheetModel.SheetName
                        : worksheetModel.SheetIndex.ToString();

                    var sheetId = sheets.Elements<Sheet>().Any()
                        ? sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1
                        : 1;

                    sheets.Append(new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = sheetId,
                        Name = sheetName,
                        State = SheetStateValues.Visible
                    });

                    var writer = OpenXmlWriter.Create(worksheetPart);

                    foreach (var table in worksheetModel.Tables)
                    {
                        writer.WriteStartElement(new Worksheet());

                        SpreadsheetBuilderHelper.AddSheetView(writer);
                        SpreadsheetBuilderHelper.AddColumns(writer, table);

                        writer.WriteStartElement(new SheetData());

                        SpreadsheetBuilderHelper.Insert(table, workbookStylesPart.Stylesheet, cellFormats, fillColours,
                            textColours, fonts, worksheetPart, writer);

                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }

                    writer.Close();
                }

                document.Close();
                return stream.ToArray();
            }
        }

        private static Stylesheet GetDefaultStyles(IEnumerable<string> backgroundColours, IEnumerable<string> textColours, IEnumerable<SpreadsheetHeaderStyle> headerStyles,
            out Dictionary<string, uint> backgroundColoursDictionary, out Dictionary<string, uint> textColoursDictionary, out Dictionary<string, uint> fontsDictionary)
        {
            var stylesheet = new Stylesheet();

            var numberingFormats1 = new NumberingFormats { Count = 1U };
            var numberingFormat1 = new NumberingFormat { NumberFormatId = 165U, FormatCode = "\"£\"#,##0.00" };

            numberingFormats1.Append(numberingFormat1);
            stylesheet.Append(numberingFormats1);

            // Add fonts
            // Standard
            var fonts = new Fonts(new Font
            {
                Bold = new Bold { Val = false },
                Italic = new Italic { Val = false },
                Strike = new Strike { Val = false },
                Underline = new Underline { Val = UnderlineValues.None },
                FontSize = new FontSize { Val = 10D },
                Color = new Color { Rgb = "00000000" },
                FontName = new FontName { Val = "Calibri" }
            }, new Font
            {
                Bold = new Bold { Val = true },
                Italic = new Italic { Val = false },
                Strike = new Strike { Val = false },
                Underline = new Underline { Val = UnderlineValues.None },
                FontSize = new FontSize { Val = 11D },
                Color = new Color { Rgb = "FF000000" },
                FontName = new FontName { Val = "Calibri" }
            });

            textColoursDictionary = new Dictionary<string, uint>();
            fontsDictionary = new Dictionary<string, uint>();

            var index = 0;

            foreach (var colour in textColours)
            {
                fonts.Append(new Font
                {
                    Bold = new Bold { Val = false },
                    Italic = new Italic { Val = false },
                    Strike = new Strike { Val = false },
                    Underline = new Underline { Val = UnderlineValues.None },
                    FontSize = new FontSize { Val = 10D },
                    Color = new Color { Rgb = "FF" + colour },
                    FontName = new FontName { Val = "Calibri" }
                });
                textColoursDictionary.Add(colour, 2U + (uint)index);
                index++;
            }

            foreach (var headerStyle in headerStyles)
            {
                var headerStyleKey = $"{ headerStyle.FontName }:{ headerStyle.TextColour}";

                if (fontsDictionary.ContainsKey(headerStyleKey)) continue;

                fonts.Append(new Font
                {
                    Bold = new Bold { Val = headerStyle.IsBold },
                    Italic = new Italic { Val = headerStyle.IsItalic },
                    Strike = new Strike { Val = headerStyle.HasStrike },
                    Underline = new Underline { Val = UnderlineValues.None },
                    FontSize = new FontSize { Val = headerStyle.FontSize },
                    Color = new Color { Rgb = "FF" + headerStyle.TextColour },
                    FontName = new FontName { Val = headerStyle.FontName }
                });
                fontsDictionary.Add(headerStyleKey, 2U + (uint)index++);
            }

            // Add fills
            var fills = new Fills();
            var fill = new Fill();
            var patternFill = new PatternFill { PatternType = PatternValues.Solid };
            var foregroundColor = new ForegroundColor { Rgb = "FF79A7E3" };

            patternFill.Append(foregroundColor);
            fill.Append(patternFill);

            fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } });      // none
            fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } });   // grey
            fills.Append(fill);                                                                                 // header


            backgroundColoursDictionary = new Dictionary<string, uint>();
            index = 0;
            foreach (var colour in backgroundColours)
            {
                fills.Append(new Fill
                {
                    PatternFill = new PatternFill()
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor { Rgb = "FF" + colour }
                    }
                });
                backgroundColoursDictionary.Add(colour, 3U + (uint)index);
                index++;
            }

            foreach (var headerStyle in headerStyles)
            {
                if (backgroundColoursDictionary.ContainsKey(headerStyle.FillColour)) continue;

                fills.Append(new Fill
                {
                    PatternFill = new PatternFill()
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor { Rgb = "FF" + headerStyle.FillColour }
                    }
                });
                backgroundColoursDictionary.Add(headerStyle.FillColour, 3U + (uint)index++);
            }

            stylesheet.Append(fonts);
            stylesheet.Append(fills);

            // Add borders
            var borders = new Borders();

            var border1 = new Border();
            var border2 = new Border { TopBorder = new TopBorder { Style = BorderStyleValues.Thin } };
            var border3 = new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin } };
            var border4 = new Border { TopBorder = new TopBorder { Style = BorderStyleValues.Thin }, RightBorder = new RightBorder { Style = BorderStyleValues.Thin } };
            var border5 = new Border { BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin } };
            var border6 = new Border { TopBorder = new TopBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin } };
            var border7 = new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin } };
            var border8 = new Border { TopBorder = new TopBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin }, RightBorder = new RightBorder { Style = BorderStyleValues.Thin } };
            var border9 = new Border { LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin } };
            var border10 = new Border { LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin }, TopBorder = new TopBorder { Style = BorderStyleValues.Thin } };
            var border11 = new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin } };
            var border12 = new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin }, TopBorder = new TopBorder { Style = BorderStyleValues.Thin } };
            var border13 = new Border { BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin } };
            var border14 = new Border { TopBorder = new TopBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin } };
            var border15 = new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin } };
            var border16 = new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin }, TopBorder = new TopBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin } };

            borders.Append(border1);
            borders.Append(border2);
            borders.Append(border3);
            borders.Append(border4);
            borders.Append(border5);
            borders.Append(border6);
            borders.Append(border7);
            borders.Append(border8);
            borders.Append(border9);
            borders.Append(border10);
            borders.Append(border11);
            borders.Append(border12);
            borders.Append(border13);
            borders.Append(border14);
            borders.Append(border15);
            borders.Append(border16);

            stylesheet.Append(borders);

            // blank cell format list
            stylesheet.CellStyleFormats = new CellStyleFormats { Count = 1 };
            stylesheet.CellStyleFormats.AppendChild(new CellFormat());

            // Add formats
            var cellFormats = new CellFormats(new CellFormat());
            stylesheet.Append(cellFormats);

            return stylesheet;
        }
    }
}
