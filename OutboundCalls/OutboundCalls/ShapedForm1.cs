using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using Telerik.WinControls.UI;
using System.Threading;

namespace OutboundCalls
{
    public partial class ShapedForm1 : Telerik.WinControls.UI.ShapedForm
    {
        PersianCalendar pc = new PersianCalendar();
        public ShapedForm1()
        {
            InitializeComponent();
            var now = DateTime.Now;
            var yesterday = now.AddDays(-1);
            
            var year = pc.GetYear(yesterday);
            var month = pc.GetMonth(yesterday);
            var day = pc.GetDayOfMonth(yesterday);
            LoadYear(year);
            LoadDays(year, month);
            LoadMonth();
            this.month.SelectedIndex = month - 1;
            duration.SelectedIndex = 1;
            this.day.SelectedIndex = day - 1;
        }

        private void execteBtn_Click(object sender, EventArgs e)
        {
            execteBtn.Enabled = false;
            waiting.Visible = true;
            waiting.StartWaiting();

            var thread = new Thread(() => Main());
            thread.Start();
        }

        public void Main()
        {
            string reportPath = string.Empty;
            string shiftPath = string.Empty;
            string outputPth = string.Empty;
            reportSelector.Invoke((MethodInvoker)(() => reportPath = reportSelector.Value));
            shiftSelector.Invoke((MethodInvoker)(() => shiftPath = shiftSelector.Value));
            outputSelector.Invoke((MethodInvoker)(() => outputPth = outputSelector.Value));
            int durationIndex = 0; int year = 0; int month = 0; int day = 0;
            this.duration.Invoke((MethodInvoker)(() => durationIndex = this.duration.SelectedIndex));
            this.year.Invoke((MethodInvoker)(() => year = int.Parse(this.year.SelectedItem.Text)));
            this.month.Invoke((MethodInvoker)(() => month = this.month.SelectedIndex + 1));
            this.day.Invoke((MethodInvoker)(() => day = this.day.SelectedIndex + 1));

            var duration = durationIndex == 0 ? new TimeSpan(0, 15, 0) : new TimeSpan(0, 30, 0);

            var manager = new Manager
            {
                ReportPath = reportPath,
                ShiftPath = shiftPath,
                CurrentDate = PersianDate.FromShamsi(year, month, day),
                OutputPath = outputPth,
                Duration = duration
            };
            manager.ReadData();
            manager.CalculateCounts();

            execteBtn.Invoke((MethodInvoker)(() => execteBtn.Enabled = true));
            waiting.Invoke((MethodInvoker)(() => waiting.Visible = false));
            waiting.Invoke((MethodInvoker)(() => waiting.StopWaiting()));
        }

        #region Initialize
        public void LoadDays(int year, int month)
        {
            var days = pc.GetDaysInMonth(year, month);
            day.Items.Clear();
            for (int i = 1; i <= days; i++)
            {
                RadListDataItem dataItem = new RadListDataItem();
                dataItem.Text = i.ToString();
                day.Items.Add(dataItem);
            }
        }
        public void LoadMonth()
        {
            var list = new List<string>();
            list.Add("فروردین"); list.Add("اردیبهشت"); list.Add("خرداد");
            list.Add("تیر"); list.Add("مرداد"); list.Add("شهریور");
            list.Add("مهر"); list.Add("آبان"); list.Add("آذر");
            list.Add("دی"); list.Add("بهمن"); list.Add("اسفند");

            month.Items.Clear();
            foreach (var month in list)
            {
                RadListDataItem dataItem = new RadListDataItem();
                dataItem.Text = month;
                this.month.Items.Add(dataItem);
            }
        }

        public void LoadYear(int year)
        {
            this.year.Items.Clear();
            for (int i = year - 5; i <= year + 5; i++)
            {
                RadListDataItem dataItem = new RadListDataItem();
                dataItem.Text = i.ToString();
                this.year.Items.Add(dataItem);
                if (i == year)
                    this.year.SelectedItem = dataItem;
            }
        }

        private void month_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
        {
            var index = this.month.SelectedIndex;
            var month = index + 1;
            var year = int.Parse(this.year.SelectedItem.Text);
            LoadDays(year, month);
            this.day.SelectedIndex = 0;
        }
        #endregion
    }
}
