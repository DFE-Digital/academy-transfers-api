using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Frontend.Helpers;
using Frontend.Models.Forms;

namespace Frontend.Models
{
    public class TransferDatesViewModel : ProjectViewModel
    {
        public readonly FormErrorsViewModel FormErrors;

        public TransferDatesViewModel()
        {
            FormErrors = new FormErrorsViewModel();
        }

        public DateInputViewModel TransferFirstDiscussed => DateInputForField(Project.Dates.FirstDiscussed);
        public DateInputViewModel TargetDate => DateInputForField(Project.Dates.Target);

        private static DateInputViewModel DateInputForField(string transferDatesFirstDiscussed)
        {
            var splitDate = DatesHelper.DateStringToDayMonthYear(transferDatesFirstDiscussed);
            return new DateInputViewModel {Day = splitDate[0], Month = splitDate[1], Year = splitDate[2]};
        }

        public List<RadioButtonViewModel> PotentialHtbDates(string startDateString)
        {
            var htbDates = DatesHelper.GetFirstWorkingDaysOfTheTheMonthForTheNextYear(startDateString);

            return htbDates.Select(htbDate => new RadioButtonViewModel
            {
                Name = "htbDate", Value = DatesHelper.DateTimeToDateString(htbDate),
                DisplayName = htbDate.ToString("dddd d MMMM yyyy"),
                Checked = Project.Dates.Htb == DatesHelper.DateTimeToDateString(htbDate)
            }).ToList();
        }
    }
}