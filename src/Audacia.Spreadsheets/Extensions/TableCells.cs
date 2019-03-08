namespace Audacia.Spreadsheets.Extensions
{
    public static class TableCells
    {
        public static string GetValue(this TableCell cell)
        {
            return cell.Value?.ToString().Trim();
        }
    }
}