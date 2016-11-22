using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutboundCalls
{
    public class Call
    {
        public string Number { get; set; }
        public PersianDate Date { get; set; }
        public TimeSpan Time { get; set; }
    }
}
