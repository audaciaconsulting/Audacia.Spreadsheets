﻿namespace Audacia.Spreadsheets
{
    /// <summary>
    /// Format codes for cell values.
    /// 
    /// Please note:
    /// Values below 1000 are supported by OpenXMl.
    /// The values above 1000 (i.e. the boolean values) are not supported by OpenXml,
    /// we have added custom formatting code to convert them to string values at the point of writing the cell.
    /// </summary>
    public enum CellFormat : uint
    {
        // Cell formats provided by OpenXMl.
        
        /// <summary>
        /// Default cell format.
        /// </summary>
        Text = 0U,
        
        /// <summary>
        /// Formats number as an integer.
        /// Format: 0
        /// </summary>
        Integer = 1U,
            
        /// <summary>
        /// Formats number as a decimal.
        /// Format: 0.00
        /// </summary>
        Decimal2Dp = 2U,

        /// <summary>
        /// Formats number as an integer.
        /// Adds commas for numbers over 1000.
        /// Format: #,##0
        /// </summary>
        IntegerWithCommas = 3U,
    
        /// <summary>
        /// Formats number as a decimal.
        /// Adds commas for numbers over 1000.
        /// Format: #,##0.00
        /// </summary>
        Decimal2DpWithCommas = 4U,

        /// <summary>
        /// Formats decimal as a percentage.
        /// Expects values from 0.00 - 1.00.
        /// Format: 0%
        /// </summary>
        Percentage = 9U,
                
        /// <summary>
        /// Formats decimal as a percentage to 2DP.
        /// Expects values from 0.00 - 1.00.
        /// Format: 0.00%
        /// </summary>
        Percentage2Dp = 10U,
        
        /// <summary>
        /// Formats decimal as standard form.
        /// Format: 0.00E+00
        /// </summary>
        Scientific = 11U,

        /// <summary>
        /// Formats decimal as a fraction.
        /// Format: # ?/?
        /// </summary>
        FractionSmall = 12U,

        /// <summary>
        /// Formats decimal as a fraction.
        /// Format: # ??/??
        /// </summary>
        FractionLarge = 13U,
        
        /// <summary>
        /// Formats number/date as a Date.
        /// Format: d/m/yyyy
        /// </summary>
        Date = 14U,
        
        /// <summary>
        /// Formats timespan to show hours and minutes.
        /// Format: H:mm
        /// </summary>
        TimeSpanHours = 20U,

        /// <summary>
        /// Formats cell as timespan.
        /// Format: H:mm:ss
        /// </summary>
        TimeSpanFull = 21U,

        /// <summary>
        /// Formats number/date as a DateTime.
        /// Format: d/mm/yyyy H:mm
        /// </summary>
        DateTime = 22U,

        /// <summary>
        /// Formats number as GBP.
        /// Format £ #,##0.00
        /// </summary>
        Currency = 44U,

        /// <summary>
        /// Formats timespan to show minutes and seconds.
        /// Format: mm:ss
        /// </summary>
        TimeSpanMinutes = 45U,

        /// <summary>
        /// Formats number as GBP.
        /// Format £ #,##0.00
        /// </summary>
        AccountingGBP = 164U,
        
        /// <summary>
        /// Formats number as USD.
        /// Format $ #,##0.00
        /// </summary>
        AccountingUSD = 165U,

        /// <summary>
        /// Formats number as Euros.
        /// Format € #,##0.00
        /// </summary>
        AccountingEUR = 166U,


        // Below are custom formats not provided by OpenXml.
        
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