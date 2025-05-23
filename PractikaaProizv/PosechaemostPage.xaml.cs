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
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Excel = Microsoft.Office.Interop.Excel;
namespace PractikaaProizv
{
    /// <summary>
    /// Логика взаимодействия для PosechaemostPage.xaml
    /// </summary>
    public partial class PosechaemostPage : Page
    {
        public PosechaemostPage()
        {
            InitializeComponent();
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddPosechaemost(null));
        }
        private void EditBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddPosechaemost((sender as Button).DataContext as RavocheeVremya));
        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            var delClients = PosechaemDG.SelectedItems.Cast<RavocheeVremya>().ToList();
            if (MessageBox.Show($"Удалить{delClients.Count} записей", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                Connect.context.RavocheeVremya.RemoveRange(delClients);
            try
            {
                Connect.context.SaveChanges();
                PosechaemDG.ItemsSource = Connect.context.RavocheeVremya.ToList();
            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void Zapros_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new Zapros());
        }
        private void Nazad_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            PosechaemDG.ItemsSource = Connect.context.RavocheeVremya.ToList();
        }
        private void CreateReport_Click(object sender, RoutedEventArgs e)
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
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic excelApp = Activator.CreateInstance(excelType);
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
            }
            catch { }
        }
        private void GenerateAndShowReport()
        {
            var data = PrepareVisitReport();
            ExportToExcel(data);
        }
    }
}
