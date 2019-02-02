using System.Linq;

namespace Audacia.Spreadsheets.Extensions
{
    public static class Strings
    {
        public static bool IsNumeric(this string input)
        {
            return !string.IsNullOrWhiteSpace(input) 
                   && input.ToCharArray()
                           .All(e => char.IsDigit(e)
                              || e == '.'
                              || e == '-');
        }
    }
}