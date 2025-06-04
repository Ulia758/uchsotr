
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
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
            try
            {
                Excel.Application app = new Excel.Application
                {
                    Visible = true,
                    SheetsInNewWorkbook = 1
                };
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                app.DisplayAlerts = false;
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
                sheet.Name = "Отчёт по сотрудникам";
                sheet.Cells[1, 1] = "Отчёт по сотрудникам";
                sheet.Range["A1:E1"].Merge();
                sheet.Cells[1, 1].Font.Bold = true;
                sheet.Cells[1, 1].Font.Size = 16;
                sheet.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sheet.Cells[2, 1] = $"Дата формирования: {DateTime.Now:dd.MM.yyyy}";
                sheet.Range["A2:E2"].Merge();
                sheet.Cells[2, 1].Font.Italic = true;
                sheet.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                sheet.Cells[4, 1] = "Фамилия";
                sheet.Cells[4, 2] = "Имя";
                sheet.Cells[4, 3] = "Отдел";
                sheet.Cells[4, 4] = "Должность";
                sheet.Cells.ColumnWidth = 30;
                sheet.Cells[1, 1].ColumnWidth = 10;
                sheet.Range["A4:E4"].Interior.Color = 0xD3D3D3; 
                sheet.Range["A4:E4"].Font.Bold = true;
                sheet.Range["A4:E4"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                var currentRow = 5;
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
                        var groupedEmployees = employeesInDepartment.GroupBy(emp => emp.Dolgnosti.Nazvanie).ToList();
                        foreach (var group in groupedEmployees)
                        {
                            sheet.Cells[currentRow, 5] = group.Key;
                            sheet.Cells[currentRow, 5].Font.Bold = true;
                            currentRow++;
                            foreach (var emp in group)
                            {
                                sheet.Cells[currentRow, 1] = emp.Familia;
                                sheet.Cells[currentRow, 2] = emp.Imya;
                                sheet.Cells[currentRow, 3] = department.Nazvanie;
                                sheet.Cells[currentRow, 4] = emp.Dolgnosti.Nazvanie;
                                currentRow++;
                            }
                            currentRow++;
                        }
                        currentRow++;
                    }
                }
                sheet.Range["A4:E" + (currentRow - 1)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Cells[currentRow, 1] = "Всего сотрудников:";
                sheet.Cells[currentRow, 2] = Connect.context.Sotrudniki.Count();
                sheet.Range["A" + currentRow + ":B" + currentRow].Font.Bold = true;
                sheet.Range["A" + currentRow + ":B" + currentRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                if (System.Windows.MessageBox.Show("Хотите сохранить отчёт?", "Сохранение отчёта", System.Windows.MessageBoxButton.YesNo) == System.Windows.MessageBoxResult.Yes)
                {
                    using (var saveDialog = new SaveFileDialog())
                    {
                        saveDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
                        saveDialog.Title = "Сохранить отчёт";
                        saveDialog.FileName = "Отчёт по сотрудникам.xlsx";

                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            workbook.SaveAs(saveDialog.FileName);
                            System.Windows.MessageBox.Show($"Файл успешно сохранён по адресу: {saveDialog.FileName}", "Успех");
                        }
                    }
                }
                if (System.Windows.MessageBox.Show("Хотите сохранить отчет?", "Сохранение отчета", System.Windows.MessageBoxButton.YesNo) == System.Windows.MessageBoxResult.Yes)
                {
                    using (var saveDialog = new System.Windows.Forms.SaveFileDialog())
                    {
                        saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                        saveDialog.Title = "Сохранить отчет";
                        saveDialog.FileName = "Отчет по сотрудникам.xlsx";

                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            workbook.SaveAs(saveDialog.FileName);
                            System.Windows.MessageBox.Show($"Файл успешно сохранён по адресу: {saveDialog.FileName}", "Успех");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка при формировании отчета: {ex.Message}", "Ошибка");
            }
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
            public double TotalHours { get; set; } 
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
                worksheet.Cells[1, 1] = "Отчёт по посещаемости сотрудников";
                worksheet.Range["A1:L1"].Merge();
                worksheet.Cells[1, 1].Font.Bold = true;
                worksheet.Cells[1, 1].Font.Size = 16;
                worksheet.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                worksheet.Cells[2, 1] = $"Дата формирования: {DateTime.Now:dd.MM.yyyy}";
                worksheet.Range["A2:L2"].Merge();
                worksheet.Cells[2, 1].Font.Italic = true;
                worksheet.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                worksheet.Cells[4, 1] = "ФИО сотрудника";
                worksheet.Cells[4, 2] = "Даты посещений";
                int columnIndex = 3;
                foreach (var date in GetLastTenWorkdays())
                {
                    worksheet.Cells[4, columnIndex++] = date.ToString("d MMMM yyyy");
                }
                worksheet.Cells[4, columnIndex++] = "Сумма часов за период";
                worksheet.Cells[4, columnIndex++] = "Присутствовал за период";
                worksheet.Range["A4:L4"].Interior.Color = 0xD3D3D3; 
                worksheet.Range["A4:L4"].Font.Bold = true;
                worksheet.Range["A4:L4"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                int rowIndex = 5;
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
                worksheet.Range["A4:L" + (rowIndex - 1)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                worksheet.Cells[rowIndex, 1] = "Всего сотрудников:";
                worksheet.Cells[rowIndex, 2] = data.Count;
                worksheet.Range["A" + rowIndex + ":B" + rowIndex].Font.Bold = true;
                worksheet.Range["A" + rowIndex + ":B" + rowIndex].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                worksheet.Columns.AutoFit();
                app.Visible = true;
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(app);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка при формировании отчета: {ex.Message}", "Ошибка");
            }
        }
        private void GenerateAndShowReport()
        {
            var data = PrepareVisitReport();
            ExportToExcel(data);
        }
        private void Otchet3_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Excel.Application app = new Excel.Application
                {
                    Visible = true,
                    SheetsInNewWorkbook = 1
                };
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                app.DisplayAlerts = false;
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.Item[1];
                sheet.Name = "Отчеты об отпусках";
                sheet.Cells[1, 1] = "Отчеты об отпусках сотрудников";
                sheet.Range["A1:E1"].Merge();
                sheet.Cells[1, 1].Font.Bold = true;
                sheet.Cells[1, 1].Font.Size = 16;
                sheet.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sheet.Cells[2, 1] = $"Дата формирования: {DateTime.Now:dd.MM.yyyy}";
                sheet.Range["A2:E2"].Merge();
                sheet.Cells[2, 1].Font.Italic = true;
                sheet.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                sheet.Cells[4, 1] = "Код отпуска";
                sheet.Cells[4, 2] = "ФИО сотрудника";
                sheet.Cells[4, 3] = "Начало отпуска";
                sheet.Cells[4, 4] = "Окончание отпуска";
                sheet.Cells[4, 5] = "Продолжительность";
                sheet.Range["A4:E4"].Interior.Color = 0xD3D3D3; 
                sheet.Range["A4:E4"].Font.Bold = true;
                sheet.Range["A4:E4"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Columns["A:E"].AutoFit();
                sheet.Cells.ColumnWidth = 30;
                sheet.Cells[1, 1].ColumnWidth = 10;
                int currentRow = 5;
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
                    FIO = $"{Connect.context.Sotrudniki.Find(x.SotrudnikId)?.Familia ?? ""} {Connect.context.Sotrudniki.Find(x.SotrudnikId)?.Imya ?? ""} {Connect.context.Sotrudniki.Find(x.SotrudnikId)?.Otchestvo ?? ""}".Trim(),
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
                sheet.Range["A4:E" + (currentRow - 1)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Cells[currentRow, 1] = "Всего записей:";
                sheet.Cells[currentRow, 2] = result.Count();
                sheet.Range["A" + currentRow + ":B" + currentRow].Font.Bold = true;
                sheet.Range["A" + currentRow + ":B" + currentRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                if (System.Windows.MessageBox.Show("Хотите сохранить отчет?", "Сохранение отчета", System.Windows.MessageBoxButton.YesNo) == System.Windows.MessageBoxResult.Yes)
                {
                    using (var saveDialog = new SaveFileDialog())
                    {
                        saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                        saveDialog.Title = "Сохранить отчет";
                        saveDialog.FileName = "Отчеты об отпусках.xlsx";

                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            workbook.SaveAs(saveDialog.FileName);
                            System.Windows.MessageBox.Show($"Файл успешно сохранён по адресу: {saveDialog.FileName}", "Успех");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка при формировании отчета: {ex.Message}", "Ошибка");
            }
        }
        private void BackBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }

        private void Otchet4_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Excel.Application app = new Excel.Application
                {
                    Visible = true,
                    SheetsInNewWorkbook = 1
                };
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                app.DisplayAlerts = false;
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.Item[1];
                sheet.Name = "Отчёт по сотрудникам";
                sheet.Cells[1, 1] = "Отчёт по сотрудникам";
                sheet.Range["A1:H1"].Merge();
                sheet.Cells[1, 1].Font.Bold = true;
                sheet.Cells[1, 1].Font.Size = 16;
                sheet.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sheet.Cells[2, 1] = $"Дата формирования: {DateTime.Now:dd.MM.yyyy}";
                sheet.Range["A2:H2"].Merge();
                sheet.Cells[2, 1].Font.Italic = true;
                sheet.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                sheet.Cells[4, 1] = "Название отдела";
                sheet.Cells[4, 2] = "ФИО сотрудника";
                sheet.Cells[4, 3] = "ID сотрудника";
                sheet.Cells[4, 4] = "Название должности";
                sheet.Cells[4, 5] = "Оклад";
                sheet.Cells[4, 6] = "Дата зачисления";
                sheet.Cells[4, 7] = "ID отделения";
                sheet.Range["A4:H4"].Interior.Color = 0xD3D3D3; 
                sheet.Range["A4:H4"].Font.Bold = true;
                sheet.Range["A4:H4"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Columns["A:H"].AutoFit(); 
                sheet.Cells.ColumnWidth = 30; 
                var departments = Connect.context.Otdeli.OrderBy(dept => dept.Nazvanie).ToList();
                int currentRow = 5;
                foreach (var department in departments)
                {
                    var employeesInDepartment = Connect.context.Sotrudniki
                        .Where(emp => emp.Dolgnosti.OtdelId == department.IdOtdel)
                        .ToList();
                    if (!employeesInDepartment.Any()) continue;
                    sheet.Cells[currentRow, 1] = department.Nazvanie;
                    sheet.Cells[currentRow, 1].Font.Bold = true;
                    sheet.Cells[currentRow, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    currentRow++;
                    foreach (var employee in employeesInDepartment)
                    {
                        var firstSchratnoeRaspisanie = employee.Dolgnosti.SchtatnoeRaspisanie.FirstOrDefault();
                        var oklad = firstSchratnoeRaspisanie?.Oklad ?? 0;
                        var firstZachislenie = employee.Zachislenie.FirstOrDefault();
                        var dateZach = firstZachislenie?.DateZach.ToShortDateString() ?? "-";
                        sheet.Cells[currentRow, 1] = department.Nazvanie;
                        sheet.Cells[currentRow, 2] = $"{employee.Familia} {employee.Imya} {employee.Otchestvo}";
                        sheet.Cells[currentRow, 3] = employee.IdSotr;
                        sheet.Cells[currentRow, 4] = employee.Dolgnosti.Nazvanie;
                        sheet.Cells[currentRow, 5] = oklad;
                        sheet.Cells[currentRow, 6] = dateZach;
                        sheet.Cells[currentRow, 7] = department.IdOtdel;
                        currentRow++;
                    }
                    currentRow += 2; 
                }
                sheet.Range["A4:H" + (currentRow - 1)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Cells[currentRow, 1] = "Всего сотрудников:";
                sheet.Cells[currentRow, 2] = Connect.context.Sotrudniki.Count();
                sheet.Range["A" + currentRow + ":B" + currentRow].Font.Bold = true;
                sheet.Range["A" + currentRow + ":B" + currentRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                if (System.Windows.MessageBox.Show("Хотите сохранить отчёт?", "Сохранение отчёта", System.Windows.MessageBoxButton.YesNo) == System.Windows.MessageBoxResult.Yes)
                {
                    using (var saveDialog = new System.Windows.Forms.SaveFileDialog())
                    {
                        saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                        saveDialog.Title = "Сохранить отчёт";
                        saveDialog.FileName = "Отчёт по сотрудникам.xlsx";

                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            workbook.SaveAs(saveDialog.FileName);
                            System.Windows.MessageBox.Show($"Файл успешно сохранён по адресу: {saveDialog.FileName}", "Успех");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка при формировании отчета: {ex.Message}", "Ошибка");
            }
        }
        private void Otchet5_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Excel.Application app = new Excel.Application
                {
                    Visible = true,
                    SheetsInNewWorkbook = 1
                };
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                app.DisplayAlerts = false;
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.Item[1];
                sheet.Name = "Отчёт по ставкам";
                sheet.Cells[1, 1] = "Отчёт по ставкам";
                sheet.Range["A1:H1"].Merge();
                sheet.Cells[1, 1].Font.Bold = true;
                sheet.Cells[1, 1].Font.Size = 16;
                sheet.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sheet.Cells[2, 1] = $"Дата формирования: {DateTime.Now:dd.MM.yyyy}";
                sheet.Range["A2:H2"].Merge();
                sheet.Cells[2, 1].Font.Italic = true;
                sheet.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                sheet.Cells[4, 1] = "Код должности";
                sheet.Cells[4, 2] = "Название отдела";
                sheet.Cells[4, 3] = "ID сотрудника";
                sheet.Cells[4, 4] = "ФИО сотрудника";
                sheet.Cells[4, 5] = "Оклад";
                sheet.Cells[4, 6] = "Кол-во ставок";
                sheet.Cells[4, 7] = "Занятые ставки"; 
                sheet.Cells[4, 8] = "Свободные ставки";
                sheet.Range["A4:H4"].Interior.Color = 0xD3D3D3; 
                sheet.Range["A4:H4"].Font.Bold = true;
                sheet.Range["A4:H4"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Columns["A:H"].AutoFit();
                sheet.Cells.ColumnWidth = 30; 
                var positions = Connect.context.Dolgnosti.Include(p => p.Otdeli).ToList();
                int currentRow = 5;
                foreach (var position in positions)
                {
                    var employeesInPosition = Connect.context.Sotrudniki
                        .Where(emp => emp.IdDolgnost == position.IdDolgnost)
                        .ToList();
                    if (!employeesInPosition.Any()) continue;
                    var mainSchedule = position.SchtatnoeRaspisanie.FirstOrDefault();
                    int totalStavki = 0;
                    int freeStavki = 0;
                    int oklad = 0;
                    if (mainSchedule != null)
                    {
                        totalStavki = mainSchedule.KolvoStavok;
                        freeStavki = totalStavki - employeesInPosition.Count;
                        oklad = mainSchedule.Oklad;
                    }
                    sheet.Cells[currentRow, 1] = position.IdDolgnost;
                    sheet.Cells[currentRow, 2] = position.Otdeli.Nazvanie;
                    sheet.Cells[currentRow, 3] = "-"; 
                    sheet.Cells[currentRow, 4] = "-"; 
                    sheet.Cells[currentRow, 5] = oklad;
                    sheet.Cells[currentRow, 6] = totalStavki;
                    sheet.Cells[currentRow, 7] = employeesInPosition.Count; 
                    sheet.Cells[currentRow, 8] = freeStavki;
                    currentRow++;
                    foreach (var employee in employeesInPosition)
                    {
                        sheet.Cells[currentRow, 3] = employee.IdSotr;
                        sheet.Cells[currentRow, 4] = $"{employee.Familia} {employee.Imya} {employee.Otchestvo}";
                        currentRow++;
                    }
                    currentRow += 2; 
                }
                sheet.Range["A4:H" + (currentRow - 1)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Cells[currentRow, 1] = "Итого позиций:";
                sheet.Cells[currentRow, 2] = positions.Count();
                sheet.Range["A" + currentRow + ":B" + currentRow].Font.Bold = true;
                sheet.Range["A" + currentRow + ":B" + currentRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                if (System.Windows.MessageBox.Show("Хотите сохранить отчёт?", "Сохранение отчёта", System.Windows.MessageBoxButton.YesNo) == System.Windows.MessageBoxResult.Yes)
                {
                    using (var saveDialog = new System.Windows.Forms.SaveFileDialog())
                    {
                        saveDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
                        saveDialog.Title = "Сохранить отчёт";
                        saveDialog.FileName = "Отчёт по должностям.xlsx";

                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            workbook.SaveAs(saveDialog.FileName);
                            System.Windows.MessageBox.Show($"Файл успешно сохранён по адресу: {saveDialog.FileName}", "Успех");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка при формировании отчёта: {ex.Message}", "Ошибка");
            }
        }
    }
}
