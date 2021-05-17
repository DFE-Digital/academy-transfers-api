using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Frontend.Helpers
{
    public static class DatesHelper
    {
        public static string DateTimeToDateString(DateTime dateTime)
        {
            return dateTime.ToString("dd/MM/yyyy");
        }

        public static string DayMonthYearToDateString(string day, string month, string year)
        {
            day = string.IsNullOrEmpty(day) ? "" : day.PadLeft(2, '0');
            month = string.IsNullOrEmpty(month) ? "" : month.PadLeft(2, '0');
            year = string.IsNullOrEmpty(year) ? "" : year.PadLeft(4, '0');

            return $"{day}/{month}/{year}";
        }

        public static List<string> DateStringToDayMonthYear(string dateString)
        {
            if (string.IsNullOrEmpty(dateString))
            {
                return new List<string>() {"", "", ""};
            }

            return dateString.Split("/").ToList();
        }

        public static string DateStringToGovUkDate(string dateString)
        {
            var splitDate = dateString.Split("/");
            var date = new DateTime(int.Parse(splitDate[2]), int.Parse(splitDate[1]), int.Parse(splitDate[0]));
            return date.ToString("d MMMM yyyy");
        }

        public static bool IsValidDate(string dateString)
        {
            return DateTime.TryParseExact(dateString, "dd/MM/yyyy", null, DateTimeStyles.None, out _);
        }

        public static List<DateTime> GetFirstWorkingDaysOfTheTheMonthForTheNextYear(string startDateString)
        {
            DateTime.TryParseExact(startDateString, "dd/MM/yyyy", null, DateTimeStyles.None, out var date);
            var dates = new List<DateTime>();

            if (DateIsFirstWorkingDayOfTheMonth(date))
            {
                dates.Add(date);
            }

            while (dates.Count < 12)
            {
                date = date.AddMonths(1);
                date = new DateTime(date.Year, date.Month, 1);

                if (DateIsAWeekend(date))
                {
                    date = GetNextMondayForDate(date);
                }

                dates.Add(date);
            }

            return dates;
        }

        private static DateTime GetNextMondayForDate(DateTime date)
        {
            if (date.DayOfWeek == DayOfWeek.Saturday)
            {
                date = date.AddDays(2);
            }

            if (date.DayOfWeek == DayOfWeek.Sunday)
            {
                date = date.AddDays(1);
            }

            return date;
        }

        private static bool DateIsAWeekend(DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
        }

        private static bool DateIsFirstWorkingDayOfTheMonth(DateTime date)
        {
            return !DateIsAWeekend(date) && date.DayOfWeek == DayOfWeek.Monday && date.Day <= 3 || date.Day == 1;
        }
    }
}