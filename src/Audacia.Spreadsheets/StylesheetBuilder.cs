using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class StylesheetBuilder
    {
        private readonly ICollection<TableHeaderStyle> _distinctHeaderStyles;
        private readonly ICollection<string> _distinctBackgroundColours;
        private readonly ICollection<string> _distinctTextColours;

        public StylesheetBuilder(IEnumerable<Worksheet> worksheets)
        {
            var allTables = worksheets
                .SelectMany(w => w.Tables)
                .ToArray();

            _distinctHeaderStyles = allTables
                .Where(t => t.HeaderStyle != null)
                .Select(t => t.HeaderStyle)
                .Distinct()
                .ToArray();

            _distinctBackgroundColours = allTables
                .SelectMany(dt => dt.Rows)
                .SelectMany(r => r.Cells.Select(c => c.FillColour))
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Distinct()
                .ToArray();

            _distinctTextColours = allTables
                .SelectMany(dt => dt.Rows)
                .SelectMany(r => r.Cells.Select(c => c.TextColour))
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Distinct()
                .ToArray();
        }
        
        public Stylesheet GetDefaultStyles(out Dictionary<string, uint> backgroundColoursDictionary, 
            out Dictionary<string, uint> textColoursDictionary,
            out Dictionary<string, uint> headerFontsDictionary)
        {
            var stylesheet = new Stylesheet();

            var numberingFormats1 = new NumberingFormats { Count = 1U };
            var numberingFormat1 = new NumberingFormat { NumberFormatId = 165U, FormatCode = "\"£\"#,##0.00" };

            numberingFormats1.Append(numberingFormat1);
            stylesheet.Append(numberingFormats1);

            // Add fonts
            var fonts = new Fonts();
            var fills = new Fills();

            textColoursDictionary = GetTextColours(fonts);
            backgroundColoursDictionary = GetFillColours(fills);
            headerFontsDictionary = GetHeaderStyles(backgroundColoursDictionary, fonts, fills);

            stylesheet.Append(fonts);
            stylesheet.Append(fills);

            // Add borders
            var borders = CreateBorders();
            stylesheet.Append(borders);

            // blank cell format list
            stylesheet.CellStyleFormats = new CellStyleFormats { Count = 1 };
            stylesheet.CellStyleFormats.AppendChild(new CellFormat());

            // Add formats
            var cellFormats = new CellFormats(new CellFormat());
            stylesheet.Append(cellFormats);

            return stylesheet;
        }
        
        private Dictionary<string, uint> GetTextColours(Fonts fonts)
        {
            var textColourDictionary = new Dictionary<string, uint>();

            var defaultFonts = new []
            {
                new Font
                {
                    Bold = new Bold { Val = false },
                    Italic = new Italic { Val = false },
                    Strike = new Strike { Val = false },
                    Underline = new Underline { Val = UnderlineValues.None },
                    FontSize = new FontSize { Val = 10D },
                    Color = new Color { Rgb = "00000000" },
                    FontName = new FontName { Val = "Calibri" }
                },
                new Font
                {
                    Bold = new Bold { Val = true },
                    Italic = new Italic { Val = false },
                    Strike = new Strike { Val = false },
                    Underline = new Underline { Val = UnderlineValues.None },
                    FontSize = new FontSize { Val = 11D },
                    Color = new Color { Rgb = "FF000000" },
                    FontName = new FontName { Val = "Calibri" }
                }
            };
            fonts.Append(defaultFonts);

            var index = 2;
            foreach (var colour in _distinctTextColours)
            {
                if (textColourDictionary.ContainsKey(colour)) { continue; }

                var newFont = new Font
                {
                    Bold = new Bold {Val = false},
                    Italic = new Italic {Val = false},
                    Strike = new Strike {Val = false},
                    Underline = new Underline {Val = UnderlineValues.None},
                    FontSize = new FontSize {Val = 10D},
                    Color = new Color {Rgb = "FF" + colour},
                    FontName = new FontName {Val = "Calibri"}
                };
                fonts.Append(newFont);
                textColourDictionary.Add(colour, (uint)index++);
            }

            return textColourDictionary;
        }

        private Dictionary<string, uint> GetFillColours(Fills fills)
        {
            var backgroundColoursDictionary = new Dictionary<string, uint>();

            var defaultFills = new[]
            {
                new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } },                             // none
                new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } },                          // grey
                new Fill { PatternFill = new PatternFill { ForegroundColor = new ForegroundColor { Rgb = "FF79A7E3" } } }    // default header
            };
            fills.Append(defaultFills);
            
            var index = 3;
            foreach (var colour in _distinctBackgroundColours)
            {
                if (backgroundColoursDictionary.ContainsKey(colour)) { continue; }

                var newFill = new Fill
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor { Rgb = "FF" + colour }
                    }
                };
                fills.Append(newFill);
                backgroundColoursDictionary.Add(colour, (uint)index++);
            }

            return backgroundColoursDictionary;
        }

        private Dictionary<string, uint> GetHeaderStyles(Dictionary<string, uint> backgroundColoursDictionary,
            Fonts fonts, Fills fills)
        {
            var fontsDictionary = new Dictionary<string, uint>();
            var headerFontsIndex = fonts.ChildElements.Count;
            var backgroundFillsIndex = fills.ChildElements.Count;

            foreach (var headerStyle in _distinctHeaderStyles)
            {
                var headerStyleKey = $"{headerStyle.FontName}:{headerStyle.TextColour}";

                if (!fontsDictionary.ContainsKey(headerStyleKey))
                {
                    var newFont = new Font
                    {
                        Bold = new Bold {Val = headerStyle.IsBold},
                        Italic = new Italic {Val = headerStyle.IsItalic},
                        Strike = new Strike {Val = headerStyle.HasStrike},
                        Underline = new Underline {Val = UnderlineValues.None},
                        FontSize = new FontSize {Val = headerStyle.FontSize},
                        Color = new Color {Rgb = "FF" + headerStyle.TextColour},
                        FontName = new FontName {Val = headerStyle.FontName}
                    };
                    fonts.Append(newFont);
                    fontsDictionary.Add(headerStyleKey, (uint)headerFontsIndex++);
                }

                if (backgroundColoursDictionary.ContainsKey(headerStyle.FillColour)) { continue; }

                var newFill = new Fill
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor {Rgb = "FF" + headerStyle.FillColour}
                    }
                };
                fills.Append(newFill);
                backgroundColoursDictionary.Add(headerStyle.FillColour, (uint)backgroundFillsIndex++);
            }

            return fontsDictionary;
        }

        private Borders CreateBorders()
        {
            var borders = new Borders();
            var borderArray = new[]
            {
                new Border(),
                new Border { TopBorder = new TopBorder { Style = BorderStyleValues.Thin } },
                new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin } },
                new Border { TopBorder = new TopBorder { Style = BorderStyleValues.Thin }, RightBorder = new RightBorder { Style = BorderStyleValues.Thin } },
                new Border { BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin } },
                new Border { TopBorder = new TopBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin } },
                new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin } },
                new Border { TopBorder = new TopBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin }, RightBorder = new RightBorder { Style = BorderStyleValues.Thin } },
                new Border { LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin } },
                new Border { LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin }, TopBorder = new TopBorder { Style = BorderStyleValues.Thin } },
                new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin } },
                new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin }, TopBorder = new TopBorder { Style = BorderStyleValues.Thin } },
                new Border { BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin } },
                new Border { TopBorder = new TopBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin } },
                new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin } },
                new Border { RightBorder = new RightBorder { Style = BorderStyleValues.Thin }, LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin }, TopBorder = new TopBorder { Style = BorderStyleValues.Thin }, BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin } }
            };
            borders.Append(borderArray);
            return borders;
        }
    }
}