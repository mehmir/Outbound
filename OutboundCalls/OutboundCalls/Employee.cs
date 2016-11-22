using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutboundCalls
{
    public class Employee
    {
        public string Number { get; set; }
        public List<Shift> Shifts { get; set; }

        public ShiftType? GetShiftType(TimeInterval interval, PersianDate date)
        {
            var shift = from s in Shifts
                        where s.PersianDate == date
                        where s.Interval.Include(interval)
                        select s;
            if (shift.Any() && shift.Count() == 1)
                return shift.First().Type;
            return null;
        }
    }
}
