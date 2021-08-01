using System;
using System.ComponentModel;
using System.Runtime.Serialization;
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

        [DisplayName("User ID")]
        [CellFormat(CellFormat.Integer)]
        public int UserId { get; set; }

        [CellFormat(CellFormat.Text)]
        public string Username { get; set; }

        [CellFormat(CellFormat.EnumName)]
        public AccountType Type { get; set; }

        [DisplayName("Start Date")]
        [CellFormat(CellFormat.Date)]
        public DateTime StartDate { get; set; }

        [DisplayName("Working Hours")]
        [CellFormat(CellFormat.TimeSpanFull)]
        public TimeSpan WorkingHours { get; set; }

        [DisplayName("Hourly Rate")]
        [CellFormat(CellFormat.Decimal2Dp)]
        public decimal HourlyRate { get; set; }

        [DisplayName("Minimum Timeout (Mins)")]
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
