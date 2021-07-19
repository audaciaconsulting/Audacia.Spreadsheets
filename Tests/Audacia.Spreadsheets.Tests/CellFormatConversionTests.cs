using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using Audacia.Spreadsheets.Tests.Models.CellFormats;
using Xunit;

namespace Audacia.Spreadsheets.Tests
{
    public class CellFormatConversionTests
    {
        [Fact]
        public void BooleanOneZeroConversions()
        {
            Validate(new BooleanOneZeroFormat[]
            {
                true,
                false
            }, t => t.Value);
        }

        [Fact]
        public void BooleanTrueFalseConversions()
        {
            Validate(new BooleanTrueFalseFormat[]
            {
                true,
                false
            }, t => t.Value);
        }

        [Fact]
        public void BooleanYesNoConversions()
        {
            Validate(new BooleanYesNoFormat[]
            {
                true,
                false
            }, t => t.Value);
        }

        [Fact]
        public void BooleanYnConversions()
        {
            Validate(new BooleanYnFormat[]
            {
                true,
                false
            }, t => t.Value);
        }

        [Fact]
        public void TextConversions()
        {
            Validate(new TextFormat[]
            {
                "Hello World",
                "Wibble",
                "Wobble"
            }, t => t.Value);
        }

        /// <summary>
        /// Validates the ability to convert and parse a formatted value.
        /// </summary>
        /// <typeparam name="T">Row Model</typeparam>
        /// <param name="source">Input Collection</param>
        /// <param name="property">Property to compare</param>
        private static void Validate<T>(IList<T> source, Func<T, object> property)
            where T : class, new()
        {
            var bytes = Spreadsheet.FromWorksheets(source.ToWorksheet()).Export();

            var output = new WorksheetImporter<T>()
                .ParseWorksheet(Spreadsheet.FromBytes(bytes).Worksheets[0])
                .Select(importRow => importRow.IsValid
                    ? property(importRow.Data)
                    : default(T))
                .ToArray();

            var expected = source
                .Select(property)
                .ToArray();

            Assert.Equal(expected, output);
        }
    }
}
