using System;
using System.Text.RegularExpressions;

namespace Audacia.Spreadsheets.Extensions
{
    public static class CellReferences
    {
        /// <summary>
        /// Returns the row number from a cell reference string.
        /// ie; "A1" would be "1"
        /// </summary>
        public static uint GetRowNumber(this string cellReference)
        {
            var regex = new Regex(@"^(?<ColumnIndex>[A-Z]+)(?<RowIndex>\d+)");
            var match = regex.Match(cellReference);
            if (!match.Success || !match.Groups["RowIndex"].Success)
                return 0;
            return uint.Parse(match.Groups["RowIndex"].Value);
        }

        /// <summary>
        /// Returns the column letter from a cell reference string.
        /// ie; "A1" would be "A"
        /// </summary>
        public static string GetColumnLetter(this string cellReference)
        {
            var regex = new Regex(@"^(?<ColumnIndex>[A-Z]+)(?<RowIndex>\d+)");
            var match = regex.Match(cellReference);
            if (!match.Success || !match.Groups["ColumnIndex"].Success)
                return "A";
            return match.Groups["ColumnIndex"].Value;
        }

        /// <summary>
        /// Returns the column number from the given column letter.
        /// ie; "A" would be "1"
        ///
        /// Do not use on a cell reference string.
        /// Should be chained from .GetColumnLetter()
        /// </summary>
        public static int ToColumnNumber(this string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
                throw new ArgumentNullException(nameof(columnName));

            columnName = columnName.ToUpperInvariant();

            var sum = 0;

            for (var i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }

        /// <summary>
        /// Returns the column letter from a column number.
        /// ie; "1" would be "A"
        /// </summary>
        public static string ToColumnLetter(this int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}
