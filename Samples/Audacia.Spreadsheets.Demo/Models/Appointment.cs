using System;
using System.ComponentModel;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Demo.Models
{
    public class Appointment
    {
        [DisplayName("Customer Reference")]
        public string Reference { get; set; }

        [DisplayName("Start Date"), CellFormat(CellFormat.Date)]
        public DateTime StartDateTime { get; set; }

        [DisplayName("Start Time")]
        public TimeSpan Time => StartDateTime.TimeOfDay;

        [DisplayName("Duration")]
        public int DurationInMinutes { get; set; }

        [DisplayName("Employee Name")]
        public string EmployeeName { get; set; }

        [DisplayName("Customer Name")]
        public string CustomerName { get; set; }

        public override string ToString()
        {
            return $"{Reference}, {StartDateTime:d}, {Time}, {DurationInMinutes}, {EmployeeName}, {CustomerName}";
        }
    }
}
