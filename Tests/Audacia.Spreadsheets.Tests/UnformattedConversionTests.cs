using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using Audacia.Spreadsheets.Tests.Models.Unformatted;
using Xunit;

namespace Audacia.Spreadsheets.Tests
{
    /// <summary>
    /// Ensure that <see cref="WorksheetImporter{TRowModel}"/> can parse all types that have been exported without a <see cref="CellFormat"/>.
    /// </summary>
    public class UnformattedConversionTests
    {
        [Fact]
        public void BooleanConversions()
        {
            Validate(new BooleanModel[]
            {
                true,
                false
            }, t => t.Value);
        }

        [Fact]
        public void DateTimeConversions()
        {
            Validate(new DateTimeModel[]
            {
                new DateTime(1970, 1, 1, 0, 0, 0),
                new DateTime(2011, 1, 1, 18, 30, 1),
                new DateTime(2016, 10, 1, 8, 35, 5),
                new DateTime(2018, 3, 2, 23, 50, 0),
                new DateTime(2020, 11, 30, 12, 0, 15),
                new DateTime(2021, 7, 20, 20, 43, 23)
            }, t => t.Value);
        }

        [Fact]
        public void DateTimeOffsetConversions()
        {
            Validate(new DateTimeOffsetModel[]
            {
                new DateTimeOffset(new DateTime(1970, 1, 1, 0, 0, 0)),
                new DateTimeOffset(new DateTime(2011, 1, 1, 18, 30, 1)),
                new DateTimeOffset(new DateTime(2016, 10, 1, 8, 35, 5)),
                new DateTimeOffset(new DateTime(2018, 3, 2, 23, 50, 0)),
                new DateTimeOffset(new DateTime(2020, 11, 30, 12, 0, 15)),
                new DateTimeOffset(new DateTime(2021, 7, 20, 20, 43, 23))
            }, t => t.Value);
        }

        [Fact]
        public void DecimalConversions()
        {
            Validate(new DecimalModel[]
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
        public void DoubleConversions()
        {
            Validate(new DoubleModel[]
            {
                -7.59d
                -3d,
                0d,
                1d,
                3.5d,
                10.99d,
                314159265.35d
            }, t => t.Value);
        }

        [Fact]
        public void EnumConversions()
        {
            Validate(new EnumModel[]
            {
                EnumModel.Shape.Hexagon,
                EnumModel.Shape.Pentagon,
                EnumModel.Shape.Square,
                EnumModel.Shape.Triangle,
                EnumModel.Shape.Circle,
            }, t => t.Value);
        }

        [Fact]
        public void FloatConversions()
        {
            Validate(new FloatModel[]
            {
                -7.59f
                -3f,
                0f,
                1f,
                3.5f,
                10.99f,
                314159265.35f
            }, t => t.Value);
        }

        [Fact]
        public void IntegerConversions()
        {
            Validate(new IntegerModel[]
            {
                -200,
                -36,
                0,
                1,
                18,
                256,
                1024
            }, t => t.Value);
        }

        [Fact]
        public void StringConversions()
        {
            Validate(new StringModel[]
            {
                "Hello World",
                "Wibble",
                "Black Sails"
            }, t => t.Value);
        }

        [Fact]
        public void TimeSpanConversions()
        {
            Validate(new TimeSpanModel[]
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

        /// <summary>
        /// Validates the ability to convert and parse a formatted value.
        /// </summary>
        /// <typeparam name="T">Row Model</typeparam>
        /// <param name="source">Input Collection</param>
        /// <param name="property">Property to compare</param>
        private static void Validate<T>(IList<T> source, Func<T, object> property)
            where T : class, new()
        {
            // Export row models into spreadsheet file
            var bytes = Spreadsheet.FromWorksheets(source.ToWorksheet()).Export();

            // Read and parse spreadsheet into row models
            var output = new WorksheetImporter<T>()
                .ParseWorksheet(Spreadsheet.FromBytes(bytes).Worksheets[0])
                .Select(importRow => importRow.IsValid
                    ? property(importRow.Data)
                    : default(T))
                .ToArray();

            // Select out expected values to compare against
            var expected = source
                .Select(property)
                .ToArray();

            // Assert parsed collection matches the expected collection
            Assert.Equal(expected, output);
        }
    }
}
