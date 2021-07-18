using Audacia.Spreadsheets.Demo.Tasks;

namespace Audacia.Spreadsheets.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Creating custom importers & exporters
            var bookExample = new BooksTask();
            var bookFileBytes = bookExample.Export();
            var books = bookExample.Import(bookFileBytes);

            // Creating a generic datasheet export
            var accountExample = new AccountsTask();
            var accountFileBytes = accountExample.Export();
            var accounts = accountExample.Import(accountFileBytes);

            // Creating an export with multiple tables
        }
    }
}
