using System.Text.RegularExpressions;

namespace Audacia.Spreadsheets.Extensions
{
    internal static class Strings
    {
        public static string PascalToTitleCase(this string input)
        {
            // This regex is specific to how gymfinity use the CellHeaderAttribute
            return Regex.Replace(input, "([a-z](?=[A-Z]|[0-9])|[A-Z](?=[A-Z][a-z]|[0-9])|[0-9](?=[^0-9]))", "$1 ");
        }
    }
}