using System;
using System.Text.RegularExpressions;

namespace Audacia.Spreadsheets.Extensions
{
    public static class CellReferenceHelper
    {
        public static uint GetReferenceRowIndex(this string cellReference)
        {
            var regex = new Regex(@"^(?<ColumnIndex>[A-Z]+)(?<RowIndex>\d+)");
            var match = regex.Match(cellReference);
            if (!match.Success || !match.Groups["RowIndex"].Success)
                return 0;
            return uint.Parse(match.Groups["RowIndex"].Value);
        }

        public static string GetReferenceColumnIndex(this string cellReference)
        {
            var regex = new Regex(@"^(?<ColumnIndex>[A-Z]+)(?<RowIndex>\d+)");
            var match = regex.Match(cellReference);
            if (!match.Success || !match.Groups["ColumnIndex"].Success)
                return "A";
            return match.Groups["ColumnIndex"].Value;
        }

        public static int GetColumnNumber(this string columnName)
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

        public static string GetExcelColumnName(this int columnNumber)
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
