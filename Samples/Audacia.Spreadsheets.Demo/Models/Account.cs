using System;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Demo.Models
{
    public class Account
    {
        public enum AccountType
        {
            Guest,
            User,
            Moderator,
            Administrator
        }

        [CellFormat(CellFormat.Integer)]
        public int UserId { get; set; }

        [CellFormat(CellFormat.Text)]
        public string Username { get; set; }

        [CellFormat(CellFormat.Text)]
        public AccountType Type { get; set; }

        [CellFormat(CellFormat.Date)]
        public DateTime StartDate { get; set; }

        [CellFormat(CellFormat.TimeSpanFull)] // TODO JP: Parse using the format on the property
        public TimeSpan WorkingHours { get; set; }

        [CellFormat(CellFormat.Decimal2Dp)]
        public decimal HourlyRate { get; set; }

        public double MinTimeoutInMins { get; set; }

        public float Age { get; set; }

        [CellFormat(CellFormat.BooleanYesNo)]
        public bool Enabled { get; set; }

        [CellFormat(CellFormat.DateTime)]
        public DateTimeOffset Created { get; set; }

        public override string ToString()
        {
            return $"{UserId}, {Username}, {Type}, {StartDate}, {WorkingHours}, {HourlyRate}, {MinTimeoutInMins}, {Age}, {Enabled}, {Created}";
        }
    }
}
