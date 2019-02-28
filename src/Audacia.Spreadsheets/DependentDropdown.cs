using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Audacia.Spreadsheets.Extensions;
using System;
using System.Collections.Generic;
using System.Text;

namespace Audacia.Spreadsheets
{
    public class DependentDropdown
    {
        public bool AllowBlanks { get; set; } = true;
        public string DependentColumn { get; set; } = "";
        public string Column { get; set; }

        /// <summary>
        /// Will create a dropdown which will look for a Named Range with the same name as the value of 'DependentColumn'
        /// </summary>
        /// <param name="dataValidations"></param>
        public void Write(DataValidations dataValidations)
        {
            //  Sets DependentColumn to previous column if none is specified.
            if (string.IsNullOrEmpty(DependentColumn))
            {
                DependentColumn = Column.PreviousColumn();
            }

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
