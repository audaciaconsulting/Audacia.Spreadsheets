using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Validation;

namespace Audacia.Spreadsheets
{
    public class ImportRow<TRowModel>
    {
        public TRowModel Data { get; set; }

        public IReadOnlyCollection<IImportError> ImportErrors { get; set; }

        public bool IsValid => !ImportErrors.Any();
    }
}
