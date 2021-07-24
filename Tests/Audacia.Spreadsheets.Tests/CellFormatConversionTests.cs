using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using Audacia.Spreadsheets.Tests.Models.CellFormats;
using Xunit;

namespace Audacia.Spreadsheets.Tests
{
    /// <summary>
    /// Ensure that <see cref="WorksheetImporter{TRowModel}"/> can parse all types that have been exported with a <see cref="CellFormat"/>.
    /// </summary>
    public class CellFormatConversionTests
    {
        [Fact]
        public void AccountingEurConversions()
        {
            Validate(new AccountingEurModel[]
            {
                0m,
                1m,
                3.5m,
                10.99m,
                314159265.35m
            }, t => t.Value);
        }

        [Fact]
        public void AccountingGbpConversions()
        {
            Validate(new AccountingGbpModel[]
            {
                0m,
                1m,
                3.5m,
                10.99m,
                314159265.35m
            }, t => t.Value);
        }

        [Fact]
        public void AccountingUsdConversions()
        {
            Validate(new AccountingUsdModel[]
            {
                0m,
                1m,
                3.5m,
                10.99m,
                314159265.35m
            }, t => t.Value);
        }

        [Fact]
        public void BooleanOneZeroConversions()
        {
            Validate(new BooleanOneZeroModel[]
            {
                true,
                false
            }, t => t.Value);
        }

        [Fact]
        public void BooleanTrueFalseConversions()
        {
            Validate(new BooleanTrueFalseModel[]
            {
                true,
                false
            }, t => t.Value);
        }

        [Fact]
        public void BooleanYesNoConversions()
        {
            Validate(new BooleanYesNoModel[]
            {
                true,
                false
            }, t => t.Value);
        }

        [Fact]
        public void BooleanYnConversions()
        {
            Validate(new BooleanYnModel[]
            {
                true,
                false
            }, t => t.Value);
        }

        [Fact]
        public void CurrencyConversions()
        {
            Validate(new CurrencyModel[]
            {
                0m,
                1m,
                3.5m,
                10.99m,
                314159265.35m
            }, t => t.Value);
        }

        [Fact]
        public void DateConversions()
        {
            Validate(new DateModel[]
            {
                new DateTime(1970, 1, 1),
                new DateTime(2011, 1, 1),
                new DateTime(2016, 10, 1),
                new DateTime(2018, 3, 2),
                new DateTime(2020, 11, 30),
                new DateTime(2021, 7, 20)
            }, t => t.Value);
        }

        [Fact]
        public void DateTimeConversions()
        {
            Validate(new DateTimeModel[]
            {
                new DateTime(1970, 1, 1, 0, 0, 0),
                new DateTime(2011, 1, 1, 18, 30, 0),
                new DateTime(2016, 10, 1, 8, 35, 5),
                new DateTime(2018, 3, 2, 23, 50, 0),
                new DateTime(2020, 11, 30, 12, 0, 15),
                new DateTime(2021, 7, 20, 20, 43, 23)
            }, t => t.Value);
        }

        [Fact]
        public void DateTimeVariantConversions()
        {
            Validate(new DateVariantModel[]
            {
                new DateTime(1970, 1, 1),
                new DateTime(2011, 1, 1),
                new DateTime(2016, 10, 1),
                new DateTime(2018, 3, 2),
                new DateTime(2020, 11, 30),
                new DateTime(2021, 7, 20)
            }, t => t.Value);
        }

        [Fact]
        public void Decimal2DpConversions()
        {
            Validate(new Decimal2DpModel[]
            {
                -7.59m
                -3m,
                0m,
                1m,
                3.5m,
                10.99m,
                314159265.35m
            }, t => t.Value);
        }

        [Fact]
        public void Decimal2DpWithCommasConversions()
        {
            Validate(new Decimal2DpWithCommasModel[]
            {
                -7.59m
                -3m,
                0m,
                1m,
                3.5m,
                10.99m,
                314159265.35m
            }, t => t.Value);
        }

        [Fact]
        public void FractionLargeConversions()
        {
            Validate(new FractionLargeModel[]
            {
                0.5m,
                0.3333333333333333m,
                0.25m,
                0.2m,
                0.1666666666666667m,
                0.1428571428571429m,
                0.125m,
                0.1111111111111111m,
                0.75m,
                1m
            }, t => t.Value);
        }

        [Fact]
        public void FractionSmallConversions()
        {
            Validate(new FractionSmallModel[]
            {
                0.5m,
                0.3333333333333333m,
                0.25m,
                0.2m,
                0.1666666666666667m,
                0.1428571428571429m,
                0.125m,
                0.1111111111111111m,
                0.75m,
                1m
            }, t => t.Value);
        }

        [Fact]
        public void PercentageConversions()
        {
            Validate(new PercentageModel[]
            {
                0.1m,
                0.25m,
                0.5m,
                0.75m,
                1m
            }, t => t.Value);
        }

        [Fact]
        public void Percentage2DpConversions()
        {
            Validate(new Percentage2DpModel[]
            {
                0.1m,
                0.25m,
                0.5m,
                0.75m,
                1m
            }, t => t.Value);
        }

        [Fact]
        public void ScientificConversions()
        {
            Validate(new ScientificModel[]
            {
                0.5m,
                0.3333333333333333m,
                0.25m,
                0.2m,
                0.1666666666666667m,
                0.1428571428571429m,
                0.125m,
                0.1111111111111111m,
                0.75m,
                1m
            }, t => t.Value);
        }

        [Fact]
        public void TextConversions()
        {
            Validate(new TextModel[]
            {
                "Hello World",
                "Wibble",
                "Black Sails"
            }, t => t.Value);
        }

        [Fact]
        public void TimeConversions()
        {
            Validate(new TimeModel[]
            {
                new TimeSpan(0, 0, 0),
                new TimeSpan(3, 0, 0),
                new TimeSpan(7, 30, 0),
                new TimeSpan(9, 13, 0),
                new TimeSpan(12, 0, 7),
                new TimeSpan(16, 43, 15),
                new TimeSpan(18, 3, 0),
            }, t => t.Value);
        }

        [Fact]
        public void TimeSpanFullConversions()
        {
            Validate(new TimeSpanFullModel[]
            {
                new TimeSpan(0, 0, 0),
                new TimeSpan(3, 0, 0),
                new TimeSpan(7, 30, 0),
                new TimeSpan(9, 13, 0),
                new TimeSpan(12, 0, 7),
                new TimeSpan(16, 43, 15),
                new TimeSpan(18, 3, 0),
            }, t => t.Value);
        }

        [Fact]
        public void TimeSpanHoursConversions()
        {
            Validate(new TimeSpanHoursModel[]
            {
                new TimeSpan(0, 0, 0),
                new TimeSpan(3, 0, 0),
                new TimeSpan(7, 30, 0),
                new TimeSpan(9, 13, 0),
                new TimeSpan(12, 0, 7),
                new TimeSpan(16, 43, 15),
                new TimeSpan(18, 3, 0),
            }, t => t.Value);
        }

        [Fact]
        public void TimeSpanMinutesConversions()
        {
            Validate(new TimeSpanMinutesModel[]
            {
                new TimeSpan(0, 0, 0),
                new TimeSpan(3, 0, 0),
                new TimeSpan(7, 30, 0),
                new TimeSpan(9, 13, 0),
                new TimeSpan(12, 0, 7),
                new TimeSpan(16, 43, 15),
                new TimeSpan(18, 3, 0),
            }, t => t.Value);
        }

        [Fact]
        public void TimeWithSecondsConversions()
        {
            Validate(new TimeWithSecondsModel[]
            {
                new TimeSpan(0, 0, 0),
                new TimeSpan(3, 0, 2),
                new TimeSpan(7, 30, 36),
                new TimeSpan(9, 13, 24),
                new TimeSpan(12, 0, 7),
                new TimeSpan(16, 43, 15),
                new TimeSpan(18, 3, 0),
            }, t => t.Value);
        }

        /// <summary>
        /// Validates the ability to convert and parse a formatted value.
        /// </summary>
        /// <typeparam name="T">Row Model</typeparam>
        /// <param name="source">Input Collection</param>
        /// <param name="propertyFunc">Property to compare</param>
        private static void Validate<T>(IList<T> source, Func<T, object> propertyFunc)
            where T : class, new()
        {
            // Export row models into spreadsheet file
            var bytes = Spreadsheet.FromWorksheets(source.ToWorksheet()).Export();

            // Read and parse spreadsheet into row models
            var output = new WorksheetImporter<T>()
                .ParseWorksheet(Spreadsheet.FromBytes(bytes).Worksheets[0])
                .ToArray();

            var actual = output
                .Select(importRow => importRow.IsValid
                    ? propertyFunc(importRow.Data)
                    : default(T))
                .ToArray();

            // Select out expected values to compare against
            var expected = source
                .Select(propertyFunc)
                .ToArray();

            // Assert parsed collection matches the expected collection
            Assert.Equal(expected, actual);

            // Ensure that a parsing failure isn't ignored
            Assert.True(output.All(t => t.IsValid));
        }
    }
}
