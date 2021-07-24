using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Validation;

namespace Audacia.Spreadsheets
{
    public class ImportRow<TRowModel>
    {
        /// <summary>
        /// The row number from the spreadsheet, defaults to zero for a column parsing errors.
        /// </summary>
        public int RowId { get; set; }

        /// <summary>
        /// The parsed row data as a <see cref="{TRowModel}"/>.
        /// </summary>
        public TRowModel Data { get; set; }

        /// <summary>
        /// An array of validation errors from attempting to parse all expected cells.
        /// </summary>
        public IReadOnlyCollection<IImportError> ImportErrors { get; set; }

        /// <summary>
        /// Returns <see cref="true"/> if there are no <see cref="ImportErrors"/>.
        /// </summary>
        public bool IsValid => !ImportErrors.Any();
    }
}
