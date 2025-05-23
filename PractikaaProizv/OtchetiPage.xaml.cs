using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace PractikaaProizv
{
    /// <summary>
    /// Логика взаимодействия для OtchetiPage.xaml
    /// </summary>
    public partial class OtchetiPage : Page
    {
        public OtchetiPage()
        {
            InitializeComponent();
        }
        private void Otchet1_Click(object sender, RoutedEventArgs e)
        {
            try {
                Excel.Application app = new Excel.Application()
                {
                    Visible = true,
                    SheetsInNewWorkbook = 1
                };
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                app.DisplayAlerts = false;
                Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.get_Item(1);
                sheet.Name = "Отчет по сотрудникам";
                sheet.Cells[1, 2] = "Фамилия";
                sheet.Cells[1, 3] = "Имя";
                sheet.Cells[1, 4] = "Отдел";
                sheet.Cells[1, 5] = "Должность";
                sheet.Cells.ColumnWidth = 30;
                sheet.Cells[1, 1].ColumnWidth = 10;
                var currentRow = 2;
                var departments = Connect.context.Otdeli.ToList();
                foreach (var department in departments)
                {
                    var employeesInDepartment = Connect.context.Sotrudniki
                        .Where(emp => emp.Dolgnosti.OtdelId == department.IdOtdel)
                        .ToList();

                    if (employeesInDepartment.Any())
                    {
                        sheet.Cells[currentRow, 1] = department.Nazvanie;
                        sheet.Cells[currentRow, 1].Font.Bold = true;
                        sheet.Cells[currentRow, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        sheet.Cells[currentRow, 1].ColumnWidth = 30;
                        currentRow++;
                        sheet.Cells[currentRow, 1] = $"Количество сотрудников: {employeesInDepartment.Count}";
                        sheet.Cells[currentRow, 1].Font.Bold = true;
                        sheet.Cells[currentRow, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        currentRow++;
                        var groupedEmployees = employeesInDepartment
                            .GroupBy(emp => emp.Dolgnosti.Nazvanie)
                            .ToList();
                        foreach (var group in groupedEmployees)
                        {
                            sheet.Cells[currentRow, 5] = group.Key;
                            sheet.Cells[currentRow, 5].Font.Bold = true;
                            currentRow++;
                            foreach (var emp in group)
                            {
                                sheet.Cells[currentRow, 2] = emp.Familia;
                                sheet.Cells[currentRow, 3] = emp.Imya;
                                sheet.Cells[currentRow, 4] = department.Nazvanie;
                                currentRow++;
                            }
                            currentRow++;
                        }
                        currentRow++;
                    }
                }
                string filePath = @"C:\Users\Public\Downloads\otchet_po_sotrudnikam.xlsx";
                try
                {
                    sheet.SaveAs(filePath);
                    MessageBox.Show($"Файл успешно сохранён по адресу: {filePath}", "Успех");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка сохранения файла: {ex.Message}", "Ошибка");
                }
            }
            catch { }
        }
        private void Otchet2_Click(object sender, RoutedEventArgs e)
        {
            GenerateAndShowReport();
        }
        public class EmployeeVisitSummary
        {
            public string FullName { get; set; }
            public Dictionary<DateTime, string> VisitStatuses { get; set; }
            public bool PresentInAnyDay { get; set; }
            public double TotalHours { get; set; } // Новое свойство для подсчета общих часов
        }
        public static List<DateTime> GetLastTenWorkdays()
        {
            var workdays = new List<DateTime>();
            var currentDate = DateTime.Now.Date;

            while (workdays.Count < 10)
            {
                if (currentDate.DayOfWeek != DayOfWeek.Saturday && currentDate.DayOfWeek != DayOfWeek.Sunday)
                {
                    workdays.Insert(0, currentDate);
                }
                currentDate = currentDate.AddDays(-1);
            }
            return workdays;
        }
        public List<EmployeeVisitSummary> PrepareVisitReport()
        {
            var lastTenWorkdays = GetLastTenWorkdays();
            var employees = Connect.context.Sotrudniki.ToList();
            var visitSummaries = new List<EmployeeVisitSummary>();
            foreach (var employee in employees)
            {
                var summary = new EmployeeVisitSummary
                {
                    FullName = employee.Familia,
                    VisitStatuses = new Dictionary<DateTime, string>(),
                    PresentInAnyDay = false,
                    TotalHours = 0
                };
                foreach (var workday in lastTenWorkdays)
                {
                    var ravocheeVremyaEntry = Connect.context.RavocheeVremya.FirstOrDefault(
                        e => e.IdSotr == employee.IdSotr && e.Data == workday);

                    if (ravocheeVremyaEntry != null)
                    {
                        summary.VisitStatuses[workday] = ravocheeVremyaEntry.Status.NameSt;
                        if (ravocheeVremyaEntry.VremyaHachalaRab.HasValue &&
    ravocheeVremyaEntry.VremyaOkonchaniyaRab.HasValue)
                        {
                            var startTime = ravocheeVremyaEntry.VremyaHachalaRab.Value;
                            var endTime = ravocheeVremyaEntry.VremyaOkonchaniyaRab.Value;
                            var timeDifference = endTime - startTime;
                            var hoursWorked = timeDifference.TotalHours;
                            summary.TotalHours += Math.Round(hoursWorked, MidpointRounding.AwayFromZero);

                            if (hoursWorked > 0)
                            {
                                summary.PresentInAnyDay = true;
                            }
                        }
                    }
                    else
                    {
                        summary.VisitStatuses[workday] = "-";
                    }
                }
                visitSummaries.Add(summary);
            }
            return visitSummaries;
        }
        public void ExportToExcel(List<EmployeeVisitSummary> data)
        {
            try {
                Excel.Application app = new Excel.Application();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Cells[1, 1] = "ФИО сотрудника";
                worksheet.Cells[1, 2] = "Даты посещений";
                int columnIndex = 3;
                foreach (var date in GetLastTenWorkdays())
                {
                    worksheet.Cells[1, columnIndex++] = date.ToString("d MMMM yyyy");
                }
                worksheet.Cells[1, columnIndex++] = "Сумма часов за период";
                worksheet.Cells[1, columnIndex++] = "Присутствовал за период";
                int rowIndex = 2;
                foreach (var employee in data)
                {
                    worksheet.Cells[rowIndex, 1] = employee.FullName;
                    worksheet.Cells[rowIndex, 2] = " ";

                    columnIndex = 3;
                    foreach (var entry in employee.VisitStatuses.OrderBy(pair => pair.Key))
                    {
                        worksheet.Cells[rowIndex, columnIndex++] = entry.Value;
                    }
                    worksheet.Cells[rowIndex, columnIndex].Value = employee.TotalHours;
                    worksheet.Cells[rowIndex, columnIndex].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    columnIndex++;

                    worksheet.Cells[rowIndex, columnIndex++] = employee.PresentInAnyDay ? "Да" : "Нет";
                    rowIndex++;
                }
                worksheet.Columns.AutoFit();
                app.Visible = true;
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(app);
            } catch { }
        }
        private void GenerateAndShowReport()
        {
            var data = PrepareVisitReport();
            ExportToExcel(data);
        }
        private void Otchet3_Click(object sender, RoutedEventArgs e)
        {
            try {
                Excel.Application app = new Excel.Application()
                {
                    Visible = true,
                    SheetsInNewWorkbook = 1
                };
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                app.DisplayAlerts = false;
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.Item[1];
                sheet.Name = "Отчеты об отпусках";
                sheet.Cells[1, 1] = "Код отпуска";
                sheet.Cells[1, 2] = "ФИО сотрудника";
                sheet.Cells[1, 3] = "Начало отпуска";
                sheet.Cells[1, 4] = "Окончание отпуска";
                sheet.Cells[1, 5] = "Продолжительность";
                sheet.Columns["A:E"].AutoFit();
                sheet.Cells.ColumnWidth = 30;
                sheet.Cells[1, 1].ColumnWidth = 10;
                int currentRow = 2;
                var otpuski = Connect.context.Otpuska
                          .Select(x => new
                          {
                              IdOtpuska = x.IdOtpuska,
                              SotrudnikId = x.IdSotr,
                              DataNachala = x.DataNachalaOtpuska,
                              DataOkonchaniya = x.DataOkonchaniyaOtpuska
                          })
                          .OrderBy(x => x.SotrudnikId);
                var result = otpuski.AsEnumerable().Select(x => new
                {
                    IdOtpuska = x.IdOtpuska,
                    FIO = $"{Connect.context.Sotrudniki.Find(x.SotrudnikId)?.Familia ?? ""} {Connect.context.Sotrudniki.Find(x.SotrudnikId)?.Imya ?? ""} ".Trim(),
                    DataNachala = x.DataNachala,
                    DataOkonchaniya = x.DataOkonchaniya
                });
                foreach (var item in result)
                {
                    TimeSpan duration = item.DataOkonchaniya.Value.Subtract(item.DataNachala.Value).Add(TimeSpan.FromDays(1));
                    sheet.Cells[currentRow, 1] = item.IdOtpuska;
                    sheet.Cells[currentRow, 2] = item.FIO;
                    sheet.Cells[currentRow, 3] = item.DataNachala.HasValue ? item.DataNachala.Value.ToShortDateString() : "";
                    sheet.Cells[currentRow, 4] = item.DataOkonchaniya.HasValue ? item.DataOkonchaniya.Value.ToShortDateString() : "";
                    sheet.Cells[currentRow, 5] = duration.Days;
                    currentRow++;
                }
                string filePath = @"C:\Users\Public\Downloads\Otchi_ob_otpuskah.xlsx";
                try
                {
                    sheet.SaveAs(filePath);
                    MessageBox.Show($"Файл успешно сохранён по адресу: {filePath}", "Успех");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка сохранения файла: {ex.Message}", "Ошибка");
                }
            } catch { }
        }

        private void BackBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }
    }
}
