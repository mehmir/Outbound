using System;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutboundCalls
{
    public class PersianDate
    {
        DateTime Date;
        static PersianCalendar pc = new PersianCalendar();

        PersianDate(DateTime date)
        {
            SetDate(date);
        }

        static public PersianDate FromShamsi(int year, int month, int day, int hour = 0, int min = 0)
        {
            var date = pc.ToDateTime(year, month, day, hour, min, 0, 0);
            return new PersianDate(date);
        }

        static public PersianDate FromMiladi(int year, int month, int day, int hour = 0, int min = 0)
        {
            var date = new DateTime(year, month, day, hour, min, 0);
            return new PersianDate(date);
        }

        private void SetDate(DateTime date)
        {
            Date = date;
        }

        public int Day()
        {
            return pc.GetDayOfMonth(Date);
        }

        public int Month()
        {
            return pc.GetMonth(Date);
        }

        public int Year()
        {
            return pc.GetYear(Date);
        }

        override public string ToString()
        {
            return string.Format("{0}/{1}/{2}", Year(), Month(), Day());
        }

        public DayOfWeek DayOfWeek()
        {
            return Date.DayOfWeek;
        }

        public int DayOfWeekIndex()
        {
            return ((int)(Date.DayOfWeek + 2) % 7);
        }

        public static bool operator ==(PersianDate pd1, PersianDate pd2)
        {
            bool result = false;
            if (pd1.Year() == pd2.Year() && pd1.Month() == pd2.Month() && pd1.Day() == pd2.Day())
                result = true;
            return result;
        }

        public static bool operator !=(PersianDate pd1, PersianDate pd2)
        {
            bool result = false;
            if (pd1.Year() != pd2.Year() || pd1.Month() != pd2.Month() || pd1.Day() != pd2.Day())
                result = true;
            return result;
        }
    }
}
