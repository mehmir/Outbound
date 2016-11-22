using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutboundCalls
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var manager = new Manager
            {
                ReportPath = @"C:\Users\mehdi\Desktop\outbound\report\Report.xlsx",
                ShiftPath = @"C:\Users\mehdi\Desktop\outbound\report\Shift.xlsx",
                CurrentDate = PersianDate.FromShamsi(1395, 8, 25)
            };
            manager.ReadData();
            manager.CalculateCounts();
        }
    }
}
