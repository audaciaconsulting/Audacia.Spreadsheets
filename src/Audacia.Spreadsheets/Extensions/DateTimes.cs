// ReSharper disable InconsistentNaming
using System;

namespace Audacia.Spreadsheets.Extensions
{
    /// <summary>
    /// For more information on date systems in Excel, please read the following article.
    /// https://support.microsoft.com/en-za/help/214330/differences-between-the-1900-and-the-1904-date-system-in-excel
    /// </summary>
    public static class DateTimes
    {
        /* For when we get to implementing 1904 dates
        /// <summary>1904 Date System, Epoch is 1st January 1904.</summary>
        public static readonly DateTime EpochJan1904 = new DateTime(1904, 1, 1);
        */
        
        /// <summary>1900 Date System, Epoch is 30th December 1899.</summary>
        public static readonly DateTime EpochJan1900 = new DateTime(1899, 12, 30);
        
        /// <summary>
        /// Converts an OADate to a DateTime
        /// .Net Standard doesn't support OADate (used by OpenXml)
        /// Ref: http://stackoverflow.com/a/13922172/1336654
        /// </summary>
        /// <param name="d">Ticks since 30th December 1899</param>
        /// <returns>A DateTime</returns>
        public static DateTime FromOADatePrecise(this double d)
        {
            if (!(d >= 0))
                throw new ArgumentOutOfRangeException(); // NaN or negative d not supported

            return EpochJan1900 + TimeSpan.FromTicks(Convert.ToInt64(d * TimeSpan.TicksPerDay));
        }

        /// <summary>
        /// Converts a DateTime to an OADate
        /// .Net Standard doesn't support OADate (used by OpenXml)
        /// Ref: http://stackoverflow.com/a/13922172/1336654
        /// </summary>
        /// <param name="dt"></param>
        /// <returns>An OADate (Ticks since 30th December 1899)</returns>
        public static double ToOADatePrecise(this DateTime dt)
        {
            if (dt < EpochJan1900)
                throw new ArgumentOutOfRangeException();

            return Convert.ToDouble((dt - EpochJan1900).Ticks) / TimeSpan.TicksPerDay;
        }
    }
}
