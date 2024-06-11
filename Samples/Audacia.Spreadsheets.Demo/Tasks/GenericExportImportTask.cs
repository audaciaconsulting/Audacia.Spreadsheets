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
    public class GenericExportImportTask
    {
        private Account[] Dataset { get; } = 
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
            Console.WriteLine("\r\nAccounts: Export() Started");

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

            Console.WriteLine("Accounts: Export() Completed\r\n");
            return bytes;
        }

        public ICollection<Account> Import(byte[] fileBytes)
        {
            Console.WriteLine("\r\nAccounts: Import() Started");

            // Read in spreadsheet, supports .FromBytes() and .FromStream()
            Console.WriteLine("Reading spreadsheet from bytes.");
            var spreadsheet = Spreadsheet.FromBytes(fileBytes);

            // Parse first worksheet on the spreadsheet
            Console.WriteLine("Converting from spreadsheet back to collection of \"Accounts\"");
            var accountImporter = new WorksheetImporter<Account>();
            var importRows = accountImporter
                .ParseWorksheet(spreadsheet.Worksheets[0])
                .ToArray();

            // Print validation errors from failure to parse data
            if (importRows.Any(r => !r.IsValid)) 
            {
                Console.WriteLine("Import Failed: see below errors");
                foreach (var invalidRow in importRows.Where(r => !r.IsValid)) 
                {
                    Console.WriteLine("Row {0}:", invalidRow.RowId);

                    foreach (var fieldError in invalidRow.ImportErrors)
                    {
                        Console.WriteLine("\t{0}", fieldError.GetMessage());
                    }
                }
            }

            var accounts = importRows
                .Where(r => r.IsValid)
                .Select(r => r.Data)
                .ToList();

            // Print out data that was imported
            Console.WriteLine("Printing dataset.");
            foreach (var a in accounts)
            {
                Console.WriteLine(a);
            }

            Console.WriteLine("Accounts: Import() Completed\r\n");
            return accounts;
        }
    }
}