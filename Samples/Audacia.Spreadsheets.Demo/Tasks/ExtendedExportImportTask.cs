using System;
using System.Collections.Generic;
using Audacia.Spreadsheets.Extensions;
using Audacia.Spreadsheets.Demo.Models;
using Audacia.Spreadsheets.Demo.Importers;
using System.Linq;

namespace Audacia.Spreadsheets.Demo.Tasks
{
    /// <summary>
    /// Example of exporting a dataset with generic logic, and importing with custom parsing and custom validation.
    /// </summary>
    public class ExtendedExportImportTask
    {
        private Appointment[] Dataset { get; } = new[]
        {
            new Appointment { Reference = "001", StartDateTime = new DateTime(2021, 07, 25, 08, 01, 00), DurationInMinutes = 60, EmployeeName = "Scott Wolter",   CustomerName = "John Smith"     },
            new Appointment { Reference = "002", StartDateTime = new DateTime(2021, 08, 07, 14, 30, 00), DurationInMinutes = 30, EmployeeName = "Chris Macort",   CustomerName = "Harry Mclean"   },
            new Appointment { Reference = "003", StartDateTime = new DateTime(2021, 09, 08, 10, 15, 00), DurationInMinutes = 15, EmployeeName = "Barry Clifford", CustomerName = "Shane Sullivan" },
            new Appointment { Reference = "004", StartDateTime = new DateTime(2021, 10, 10, 16, 45, 00), DurationInMinutes = 20, EmployeeName = "Scott Wolter",   CustomerName = "Tony Parrot"    }
        };

        public byte[] Export()
        {
            Console.WriteLine("\r\nAppointments: Export() Started");

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

            Console.WriteLine("Appointments: Export() Completed\r\n");
            return bytes;
        }

        public ICollection<Appointment> Import(byte[] fileBytes)
        {
            Console.WriteLine("\r\nAppointments: Import() Started");

            // Read in spreadsheet, supports .FromBytes() and .FromStream()
            Console.WriteLine("Reading spreadsheet from bytes.");
            var spreadsheet = Spreadsheet.FromBytes(fileBytes);

            // Parse first worksheet on the spreadsheet
            Console.WriteLine("Converting from spreadsheet back to collection of \"nAppointments\"");
            var accountImporter = new AppointmentImporter();
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

            var appointments = importRows
                .Where(r => r.IsValid)
                .Select(r => r.Data)
                .ToList();

            // Print out data that was imported
            Console.WriteLine("Printing dataset.");
            foreach (var a in appointments)
            {
                Console.WriteLine(a);
            }

            Console.WriteLine("Appointments: Import() Completed\r\n");
            return appointments;
        }
    }
}
