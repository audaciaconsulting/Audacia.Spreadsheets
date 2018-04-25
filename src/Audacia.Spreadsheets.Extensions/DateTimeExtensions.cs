// ReSharper disable InconsistentNaming
using System;

namespace Audacia.Spreadsheets.Extensions
{
    public static class DateTimeExtensions
    {
        private static readonly DateTime OaDateEpoch = new DateTime(1899, 12, 30);
        
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

            return OaDateEpoch + TimeSpan.FromTicks(Convert.ToInt64(d * TimeSpan.TicksPerDay));
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
            if (dt < OaDateEpoch)
                throw new ArgumentOutOfRangeException();

            return Convert.ToDouble((dt - OaDateEpoch).Ticks) / TimeSpan.TicksPerDay;
        }
    }
}
