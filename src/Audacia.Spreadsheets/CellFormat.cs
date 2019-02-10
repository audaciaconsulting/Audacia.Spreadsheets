namespace Audacia.Spreadsheets
{
    // TODO: support more number formats, see: https://github.com/closedxml/closedxml/wiki/NumberFormatId-Lookup-Table
    public enum CellFormat
    {
        Text = 100,
        Date = 200,
        Currency = 300,
        
        // Custom Formats not provided by OpenXml
        
        /// <summary>
        /// Formats a boolean as a number.
        /// Values 1 or 0.
        /// </summary>
        BooleanOneZero = 1000,
        
        /// <summary>
        /// Formats a boolean as a string.
        /// Values TRUE or FALSE
        /// </summary>
        BooleanTrueFalse = 1100,
        
        /// <summary>
        /// Formats a boolean as a string.
        /// Values Yes or No.
        /// </summary>
        BooleanYesNo = 1200,
        
        /// <summary>
        /// Formats a boolean as a string.
        /// Values Y or N.
        /// </summary>
        BooleanYN = 1300
    }
}
