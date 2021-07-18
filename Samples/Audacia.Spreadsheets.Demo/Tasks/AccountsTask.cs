using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Demo.Models;
using Audacia.Spreadsheets.Extensions;


namespace Audacia.Spreadsheets.Demo.Tasks
{
    /// <summary>
    /// Example of exporting and importing a dataset with generic logic.
    /// </summary>
    public class AccountsTask
    {
        private Account[] Dataset { get; } = new[]
        {
            new Account
            { 
                UserId = 1,
                Username = "USER1",
                Type = Account.AccountType.Administrator,
                StartDate = new DateTime(2016, 7, 1),
                WorkingHours = TimeSpan.FromHours(7.5),
                HourlyRate = 8.69m,
                MinTimeoutInMins = 27.34d,
                Age = 28.1f,
                Enabled = true,
                Created = DateTimeOffset.Now
            },
            new Account
            {
                UserId = 2,
                Username = "USER2",
                Type = Account.AccountType.Moderator,
                StartDate = new DateTime(2018, 2, 6),
                WorkingHours = TimeSpan.FromHours(7.5),
                HourlyRate = 6.75m,
                MinTimeoutInMins = 4d,
                Age = 26.5f,
                Enabled = true,
                Created = DateTimeOffset.Now
            },
            new Account
            {
                UserId = 3,
                Username = "USER3",
                Type = Account.AccountType.Guest,
                StartDate = new DateTime(2021, 1, 1),
                WorkingHours = TimeSpan.FromHours(7.5),
                HourlyRate = 0m,
                MinTimeoutInMins = 60d,
                Age = 45f,
                Enabled = false,
                Created = DateTimeOffset.Now
            }
        };

        public byte[] Export()
        {
            Console.WriteLine("\r\nAccountsTask: Export() Started");

            // Print out data that will be exported
            Console.WriteLine("Printing dataset.");
            foreach (var a in Dataset)
            {
                Console.WriteLine(a);
            }

            // Create an exportable worksheet
            Console.WriteLine("Creating an exportable worksheet.");
            var worksheet = Dataset.ToWorksheet();

            // Add all your worksheets into a spreadsheet
            Console.WriteLine("Creating a spreadsheet.");
            var spreadsheet = Spreadsheet.FromWorksheets(worksheet);

            // Export your spreadsheet
            // .Export() can write to a byte[] or a filepath, to write to a stream using .Write()
            Console.WriteLine("Exporting spreadsheet to a byte array.");
            var bytes = spreadsheet.Export();

            Console.WriteLine("AccountsTask: Export() Completed\r\n");
            return bytes;
        }

        public ICollection<Account> Import(byte[] fileBytes)
        {
            Console.WriteLine("\r\nAccountsTask: Import() Started");

            // Read in spreadsheet, supports .FromBytes() and .FromStream()
            Console.WriteLine("Reading spreadsheet from bytes.");
            var spreadsheet = Spreadsheet.FromBytes(fileBytes);

            // Parse first worksheet on the spreadsheet
            Console.WriteLine("Converting from spreadsheet back to collection of \"Accounts\"");
            var accountImporter = new WorksheetImporter<Account>();
            var accounts = accountImporter
                .ParseWorksheet(spreadsheet.Worksheets[0])
                .Where(r => r.IsValid)
                .Select(r => r.Data)
                .ToList();

            // Print out data that was imported
            Console.WriteLine("Printing dataset.");
            foreach (var a in accounts)
            {
                Console.WriteLine(a);
            }

            Console.WriteLine("AccountsTask: Import() Completed\r\n");
            return accounts;
        }
    }
}