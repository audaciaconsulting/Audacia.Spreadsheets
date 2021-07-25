﻿using System;
using System.ComponentModel.DataAnnotations;

namespace Audacia.Spreadsheets.Demo.Models
{
    public class Appointment
    {
        [Display(Name = "Customer Reference")]
        public string Reference { get; set; }

        [Display(Name = "Start Date")]
        public DateTime StartDateTime { get; set; }

        [Display(Name = "Start Time")]
        public TimeSpan Time => StartDateTime.TimeOfDay;

        [Display(Name = "Duration")]
        public int DurationInMinutes { get; set; }

        [Display(Name = "Employee Name")]
        public string EmployeeName { get; set; }

        [Display(Name = "Customer Name")]
        public string CustomerName { get; set; }
    }
}
