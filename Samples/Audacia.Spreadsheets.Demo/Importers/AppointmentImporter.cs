using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Demo.Models;
using Audacia.Spreadsheets.Validation;

namespace Audacia.Spreadsheets.Demo.Importers
{
    public class AppointmentImporter : WorksheetImporter<Appointment>
    {
        private readonly DateTime MinDate = new DateTime(2021, 1, 1);
        private HashSet<int> UniqueChecksums = new HashSet<int>();

        protected override IEnumerable<IImportError> ParseRow(out Appointment model)
        {
            // Optionally replace the row parser code
            model = null;
            var importErrors = new List<IImportError>();

            if (!TryGetDateTime(x => x.StartDateTime, out var start))
            {
                importErrors.Add(new FieldParseError(GetRowNumber(), GetColumnHeader(x => x.StartDateTime)));
            }

            if (!TryGetInteger(x => x.DurationInMinutes, out var duration))
            {
                importErrors.Add(new FieldParseError(GetRowNumber(), GetColumnHeader(x => x.DurationInMinutes)));
            }

            if (!TryGetString(x => x.EmployeeName, out var employee))
            {
                importErrors.Add(new FieldParseError(GetRowNumber(), GetColumnHeader(x => x.EmployeeName)));
            }

            if (!TryGetString(x => x.CustomerName, out var customer))
            {
                importErrors.Add(new FieldParseError(GetRowNumber(), GetColumnHeader(x => x.CustomerName)));
            }

            if (!importErrors.Any())
            {
                model = new Appointment
                {
                    StartDateTime = start,
                    DurationInMinutes = duration,
                    EmployeeName = employee,
                    CustomerName = customer
                };
            }

            return importErrors;
        }

        protected override IEnumerable<IImportError> ValidateRow(Appointment row)
        {
            // Optionally add additional validation
            var importErrors = new List<IImportError>();

            if (row.StartDateTime < MinDate)
            {
                importErrors.Add(new FieldValidationError(GetRowNumber(), new[]
                {
                    new ValidationResult(GetColumnHeader(x => x.StartDateTime), "Must be in the year 2021 or after.")
                }));
            }

            var checksum = row.Reference.GetHashCode();

            if (!UniqueChecksums.Add(checksum))
            {
                importErrors.Add(new DuplicateKeyError(GetRowNumber(), GetColumnHeader(x => x.Reference), row.Reference));
            }
            else if (AlreadyExistsInSystem(row)) 
            {
                importErrors.Add(new RecordExistsError(GetRowNumber(), GetColumnHeader(x => x.Reference), row.Reference));
            }

            return importErrors;
        }

        private bool AlreadyExistsInSystem(Appointment row)
        {
            // Do a database lookup for the imported appointment
            return false;
        }
    }
}
