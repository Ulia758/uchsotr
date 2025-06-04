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
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
namespace PractikaaProizv {
    /// <summary>
    /// Логика взаимодействия для OtpuskaPage.xaml
    /// </summary>
    public partial class OtpuskaPage : Page
    {
        public OtpuskaPage()
        {
            InitializeComponent();
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddOtpuska(null));
        }
        private void EditBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddOtpuska((sender as System.Windows.Controls.Button).DataContext as Otpuska));
        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            var delClients = OtpuskDG.SelectedItems.Cast<Otpuska>().ToList();
            if (System.Windows.MessageBox.Show($"Удалить{delClients.Count} записей", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                Connect.context.Otpuska.RemoveRange(delClients);
            try
            {
                Connect.context.SaveChanges();
                OtpuskDG.ItemsSource = Connect.context.Otpuska.ToList();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void Otchet_Click(object sender, RoutedEventArgs e)
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
        private void Nazad_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            OtpuskDG.ItemsSource = Connect.context.Otpuska.ToList();
        }
    }
}
