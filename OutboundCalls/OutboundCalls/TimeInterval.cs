using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutboundCalls
{
    public class TimeInterval
    {
        public TimeSpan? Start { get; set; }
        public TimeSpan? End { get; set; }

        public TimeInterval(TimeSpan? start, TimeSpan? end = null)
        {
            Start = start;
            End = end;
        }

        public bool Include(TimeSpan start, TimeSpan end)
        {
            return Start <= start && start <= End && Start <= end && end <= End;
        }

        public bool Include(TimeInterval interval)
        {
            return Include(interval.Start.Value, interval.End.Value);
        }
    }
}
