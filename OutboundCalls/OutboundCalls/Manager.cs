using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Drawing;

namespace OutboundCalls
{
    public class Manager
    {
        public string ReportPath { get; set; }
        public string ShiftPath { get; set; }
        public string OutputName { get; set; }
        public string OutputPath { get; set; }
        public PersianDate CurrentDate { get; set; }
        public TimeSpan Duration { get; set; }
        public TimeSpan StartObservation { get; set; }
        public TimeSpan EndObservation { get; set; }

        List<Employee> Employees { get; }
        List<Call> OutboundCalls { get; }
        Dictionary<string, List<OutputCell>> OutboundCountList;

        public Manager()
        {
            OutboundCalls = new List<Call>();
            Employees = new List<Employee>();
            OutboundCountList = new Dictionary<string, List<OutputCell>>();
            OutputName = "OutBoundCalls_{0}.xlsx";
            StartObservation = new TimeSpan(7, 0, 0);
            EndObservation= new TimeSpan(24, 0, 0);
        }

        #region Reading Input Files
        public void ReadData()
        {
            ReadOutboundCalls();
            ReadEmployees();
        }

        private void ReadOutboundCalls()
        {
            FileInfo fileInfo = new FileInfo(ReportPath);
            ExcelWorksheet worksheet;
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                try
                {
                    if (fileInfo.Exists)
                    {
                        worksheet = package.Workbook.Worksheets[1];
                        for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                        {
                            string type = worksheet.Cells[i, 4].Text;
                            string answerTime = worksheet.Cells[i, 6].Text;
                            string duration = worksheet.Cells[i, 7].Text;
                            string calledNumber = worksheet.Cells[i, 3].Text;
                            if (type.ToLower().Contains("outbound") && !answerTime.ToLower().Contains("null") && duration != "0" &&
                                calledNumber.Length > 3 && calledNumber.All(char.IsDigit))
                            {
                                string number = worksheet.Cells[i, 2].Text;
                                var startAnswer = worksheet.Cells[i, 5].Text.ToString().Split(' ');
                                var date = startAnswer[0].Split('/');
                                var time = startAnswer[1].Split(':');
                                var month = int.Parse(date[0]);
                                var day = int.Parse(date[1]);
                                var year = int.Parse(date[2]);
                                var hour = int.Parse(time[0]);
                                var min = int.Parse(time[1]);
                                var callDate = PersianDate.FromMiladi(year < 2000 ? 2000 + year : year, month, day);
                                var callTime = new TimeSpan(hour, min, 0);
                                var call = new Call
                                {
                                    Number = number,
                                    Date = callDate,
                                    Time = callTime
                                };
                                OutboundCalls.Add(call);
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    throw new Exception("Error in reading report File\n" + e.Message);
                }
            }
        }

        private void ReadEmployees()
        {
            FileInfo fileInfo = new FileInfo(ShiftPath);
            ExcelWorksheet worksheet;
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                try
                {
                    if (fileInfo.Exists)
                    {
                        worksheet = package.Workbook.Worksheets[CurrentDate.DayOfWeekIndex()];
                        int from = 5 + 1;
                        int to = 5 + (24 * 2);
                        for (int i = 3; i <= worksheet.Dimension.End.Row; i++)
                        {
                            int row = -1;
                            if (!int.TryParse(worksheet.Cells[i, 1].Text, out row))
                                continue;
                            var number = worksheet.Cells[i, 3].Text;
                            var shifts = CreateShifts(worksheet.Cells[i, from, i, to]);
                            var employee = new Employee
                            {
                                Number = number,
                                Shifts = shifts
                            };
                            Employees.Add(employee);
                        }
                    }
                }
                catch (Exception e)
                {
                    throw new Exception("Error in reading shift File\n" + e.Message);
                }
            }
        }

        private List<Shift> CreateShifts(ExcelRange row)
        {
            var shifts = new List<Shift>();

            int index = 0;
            string last = "";
            foreach (var cell in row)
            {
                var value = cell.Text;
                if (string.IsNullOrWhiteSpace(value))
                    value = "";
                if (value == last)
                {
                    index++;
                    continue;
                }
                switch (value)
                {
                    case "D":
                        AddShift(shifts, index, last, ShiftType.D);
                        break;
                    case "R":
                        AddShift(shifts, index, last, ShiftType.D);
                        break;
                    case "":
                        var lastShift = shifts.Last();
                        lastShift.Interval.End = TimeByIndex(index);
                        break;
                    default:
                        AddShift(shifts, index, last, ShiftType.None);
                        break;
                }
                index++;
                last = value;
            }

            return shifts;
        }

        private TimeSpan TimeByIndex(int index)
        {
            int hour = index / 2;
            int minutes = (index % 2) * 30;
            return new TimeSpan(hour, minutes, 0);
        }

        private void AddShift(List<Shift> shifts, int index, string last, ShiftType type)
        {
            if (last != "")
            {
                var lastShift = shifts.Last();
                lastShift.Interval.End = TimeByIndex(index);
            }
            var interval = new TimeInterval(TimeByIndex(index), TimeByIndex(48));
            var shift = new Shift(interval, CurrentDate, type);
            shifts.Add(shift);
        }
        #endregion

        public void CalculateCounts()
        {
            var end = EndObservation;
            foreach (var employee in Employees)
            {
                var start = StartObservation;
                OutboundCountList.Add(employee.Number, new List<OutputCell>());
                var list = OutboundCountList[employee.Number];

                while (start < end)
                {
                    var calls = from c in OutboundCalls
                                where c.Number == employee.Number
                                where c.Time >= start && c.Time < start.Add(Duration)
                                where c.Date == CurrentDate
                                select c;
                    list.Add(new OutputCell
                    {
                        Count = calls.Count(),
                        Type = employee.GetShiftType(new TimeInterval(start, start.Add(Duration)), CurrentDate)
                    });
                    start = start.Add(Duration);
                }
            }
            WritetoFile();
        }

        public void WritetoFile()
        {
            var name = string.Format(OutputPath + "/" + OutputName, CurrentDate.ToString().Replace("/","-"));
            FileInfo output = new FileInfo(name);
            ExcelWorksheet worksheet;

            if (output.Exists)
                output.Delete();

            using (ExcelPackage package = new ExcelPackage(output))
            {
                try
                {
                    int startI = 3; int startJ = 3;
                    worksheet = package.Workbook.Worksheets.Add("تعداد تماسهای خارجی");
                    var start = StartObservation;
                    var end = EndObservation;
                    int j = startJ;
                    while (start < end)
                    {
                        worksheet.Cells[startI, j].Value = start.ToString(@"hh\:mm");
                        j++;
                        start = start.Add(Duration);
                    }

                    int i = startI + 1;
                    
                    foreach (var employee in Employees)
                    {
                        j = startJ;
                        var list = OutboundCountList[employee.Number];
                        worksheet.Cells[i, j - 1].Value = employee.Number;
                        foreach (var item in list)
                        {
                            if (item.Count != 0)
                            {
                                worksheet.Cells[i, j].Value = item.Count;
                                SetColor(worksheet.Cells[i, j], Shift.GetColor(item.Type));
                            }
                            j++;
                        }
                        i++;
                    }

                    //worksheet.Cells[2, 3, 2, 5].Merge = true;
                    //worksheet.Cells[2, 3, 2, 5].Value = string.Format("روز: {0}", CurrentDate.ToString());
                    SetColor(worksheet.Cells[startI, startJ - 1, startI, j - 1], Color.LightGray);
                    SetColor(worksheet.Cells[startI, startJ - 1, startI + Employees.Count, startJ - 1], Color.LightGray);

                    //SetBorder(worksheet.Cells[2, 8, 2, 11]);
                    //SetSampleCell(worksheet.Cells[2, 8], "D", ShiftType.D);
                    //SetSampleCell(worksheet.Cells[2, 9], "R", ShiftType.R);
                    //SetSampleCell(worksheet.Cells[2, 10], "Other", ShiftType.None);
                    //SetSampleCell(worksheet.Cells[2, 11], "None", null);

                    SetBorder(worksheet.Cells[startI, startJ - 1, startI + Employees.Count, j - 1]);
                    //SetBorder(worksheet.Cells[2, 3, 2, 5]);
                    worksheet.View.RightToLeft = true;
                    worksheet.Cells.Style.ReadingOrder = ExcelReadingOrder.RightToLeft;
                    worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    package.Save();
                }
                catch (Exception e)
                {
                    throw new Exception("Error in writing output File\n" + e.Message);
                }
            }

        }

        private void SetBorder(ExcelRange range)
        {
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        private void SetColor(ExcelRange range, Color color)
        {
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(color);
        }

        private void SetSampleCell(ExcelRange range, string value, ShiftType? type)
        {
            range.Value = value;
            SetColor(range, Shift.GetColor(type));
        }
    }
}