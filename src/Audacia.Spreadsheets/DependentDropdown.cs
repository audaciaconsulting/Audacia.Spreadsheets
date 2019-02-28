using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace Audacia.Spreadsheets
{
    public class DependentDropdown
    {
        public bool AllowBlanks { get; set; }
        public string DependentColumn { get; set; }
        public string Column { get; set; }

        public void Write(DataValidations dataValidations)
        {
            DataValidation dataValidation = new DataValidation()
            {
                Type = DataValidationValues.List,
                AllowBlank = AllowBlanks,
                SequenceOfReferences = new ListValue<StringValue>() { InnerText = $"{Column}2:{Column}1048576" }
            };
            Formula1 formula = new Formula1
            {
                Text = $"=INDIRECT(SUBSTITUTE(${DependentColumn}2, \" \", \"_\"))"
            };

            dataValidation.Append(formula);
            dataValidations.Append(dataValidation);

        }
    }
}
