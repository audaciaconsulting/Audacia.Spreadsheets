using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace Audacia.Spreadsheets
{
    public class StaticDropdown
    {
        public bool AllowBlanks { get; set; }
        public int StartingRow { get; set; } = 2;
        public int EndingRow { get; set; } = 1048576;
        public string Column { get; set; }
        public List<string> Options { get; set; }

        public void Write(DataValidations dataValidations)
        {
            DataValidation dataValidation = new DataValidation()
            {
                Type = DataValidationValues.List,
                AllowBlank = AllowBlanks,
                SequenceOfReferences = new ListValue<StringValue>() { InnerText = $"{Column}{StartingRow}:{Column}{EndingRow}" }
            };
            Formula1 formula = new Formula1
            {
                Text = $"\"{string.Join(",", Options)}\""
            };

            dataValidation.Append(formula);
            dataValidations.Append(dataValidation);

        }
    }
}
