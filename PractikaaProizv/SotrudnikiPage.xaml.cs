using System;
using System.Collections.Generic;
using System.Linq;
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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
namespace PractikaaProizv
{
    /// <summary>
    /// Логика взаимодействия для SotrudnikiPage.xaml
    /// </summary>
    public partial class SotrudnikiPage : Page
    {
        private List<Sotrudniki> allEmployees;

        public SotrudnikiPage()
        {
            InitializeComponent();
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddSotrudniki(null));
        }
        private void EditBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddSotrudniki((sender as Button).DataContext as Sotrudniki));
        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            var delClients = SotrDG.SelectedItems.Cast<Sotrudniki>().ToList();
            if (MessageBox.Show($"Удалить {delClients.Count} записей", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                Connect.context.Sotrudniki.RemoveRange(delClients);
            try
            {
                Connect.context.SaveChanges();
                SotrDG.ItemsSource = Connect.context.Sotrudniki.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void Nazad_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            allEmployees = Connect.context.Sotrudniki.ToList();
            SotrDG.ItemsSource = allEmployees;

            var departments = Connect.context.Otdeli.ToList();
            DepartmentFilter.Items.Clear(); 
            DepartmentFilter.Items.Add(new ComboBoxItem { Content = "Все", Tag = "" }); 
            foreach (var department in departments)
            {
                DepartmentFilter.Items.Add(new ComboBoxItem { Content = department.Nazvanie, Tag = department.IdOtdel });
            }
        }
        private void DepartmentFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilter();
        }
        private void ApplyFilter()
        {
            if (SotrDG == null || allEmployees == null)
            {
                return;
            }
            string selectedDepartment = ((ComboBoxItem)DepartmentFilter.SelectedItem)?.Tag?.ToString();

            if (string.IsNullOrEmpty(selectedDepartment))
            {
                SotrDG.ItemsSource = allEmployees; 
            }
            else
            {
                var filteredList = allEmployees.Where(emp =>
                    emp.Dolgnosti != null &&
                    emp.Dolgnosti.Otdeli != null &&
                    emp.Dolgnosti.Otdeli.IdOtdel.ToString() == selectedDepartment).ToList();
                SotrDG.ItemsSource = filteredList;
            }
        }
        private void ReportBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic excelApp = Activator.CreateInstance(excelType);

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
                try
                {
                    sheet.SaveAs("C:\\Users\\Public\\Downloads\\otchet_po_sotrudnikam.xlsx");
                    MessageBox.Show("Отчет успешно сохранен в загрузки!", "Информация");

                }
                catch
                {
                    MessageBox.Show("Не удаётся получить путь к файлу", "Ошибка");
                }
            }
            catch { }
            
        }
    }
}
