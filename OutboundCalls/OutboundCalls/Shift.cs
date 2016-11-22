using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutboundCalls
{
    public class Shift
    {
        public TimeInterval Interval { get; set; }
        public ShiftType Type { get; set; }
        public PersianDate PersianDate { get; set; }

        public Shift(TimeInterval interval, PersianDate date, ShiftType type)
        {
            Interval = interval;
            PersianDate = date;
            Type = type;
        }

        public static Color GetColor(ShiftType? type)
        {
            Color c;

            switch (type)
            {
                case ShiftType.D:
                    c = Color.Orange;
                    break;
                case ShiftType.R:
                    c = Color.HotPink;
                    break;
                case ShiftType.None:
                    c = Color.LightGreen;
                    break;
                default:
                    c = Color.Red;
                    break;
            }
            return c;
        }
    }
}
