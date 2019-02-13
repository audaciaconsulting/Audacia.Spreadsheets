using System;
using System.Linq;

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
            var numberChars = cellReference.Where(char.IsNumber).ToArray();
            var rowNumber = new string(numberChars);
            if (uint.TryParse(rowNumber, out var result))
            {
                return result;
            }
            
            return 0;
        }

        /// <summary>
        /// Returns the column letter from a cell reference string.
        /// ie; "A1" would be "A"
        /// </summary>
        public static string GetColumnLetter(this string cellReference)
        {
            var columnChars = cellReference.Where(char.IsLetter).ToArray();
            if (columnChars.Any())
            {
                return new string(columnChars);
            }

            return "A";
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
