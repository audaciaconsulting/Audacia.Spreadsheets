using System;

namespace Audacia.Spreadsheets.Extensions
{
    public static class Timespans
    {
        public static double ToOADatePrecise(this TimeSpan time)
        {
            return Convert.ToDouble(time.Ticks) / TimeSpan.TicksPerDay;
        }
    }
}