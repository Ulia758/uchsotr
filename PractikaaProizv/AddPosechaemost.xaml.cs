using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Validation;
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
namespace PractikaaProizv
{
    /// <summary>
    /// Логика взаимодействия для AddPosechaemost.xaml
    /// </summary>
    public partial class AddPosechaemost : Page
    {
        RavocheeVremya k;
        private Status tempSelectedStatus; // Временная переменная для выбранного статуса
        private bool isDirty;
        public AddPosechaemost(RavocheeVremya c)
        {
            InitializeComponent();
            StatusComboBox.ItemsSource = Connect.context.Status.ToList();

            // Передаем исходный объект и устанавливаем контекст
            if (c == null)
                c = new RavocheeVremya();
            DataContext = k = c;

            // Устанавливаем начальное значение статуса
            StatusComboBox.SelectedItem = k.Status ?? default(Status);

        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (!isDirty || tempSelectedStatus == null)
            {
                return;
            }

            k.Status = tempSelectedStatus;
            if (tempSelectedStatus.IdStatus != 1)
            {
                RemoveTimeFromDatabase();
            }

            if (k.IdRabocheeVremya == 0)
            {
                Connect.context.RavocheeVremya.Add(k);
            }

            try
            {
                Connect.context.SaveChanges();
            }
            catch (DbEntityValidationException ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            Nav.MainFrame.GoBack();
        }
    
        private void Back_Click(object sender, RoutedEventArgs e)
        {
            if (Connect.context.ChangeTracker.HasChanges())
            {
                Connect.context.ChangeTracker.Entries().ToList().ForEach(x => x.State = EntityState.Unchanged);
            }

            Nav.MainFrame.GoBack();
        }

        private void StatusComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var newStatus = StatusComboBox.SelectedItem as Status;
            if (newStatus != null)
            {
                tempSelectedStatus = newStatus;

                // Отмечаем, что данные были изменены
                isDirty = true;

                if (newStatus.IdStatus == 1)
                {
                    VremyaHachalaRabTextBox.IsEnabled = true;
                    VremyaOkonchaniyaRabTextBox.IsEnabled = true;
                }
                else
                {
                    VremyaHachalaRabTextBox.IsEnabled = false;
                    VremyaOkonchaniyaRabTextBox.IsEnabled = false;
                }
            }
        }
        private void RemoveTimeFromDatabase()
        {
            if (k != null)
            {
                k.VremyaHachalaRab = null;
                k.VremyaOkonchaniyaRab = null;
                try
                {
                    Connect.context.SaveChanges();
                }
                catch (DbEntityValidationException ex)
                {
                    MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        private void DateTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (DateTime.TryParse(((TextBox)sender).Text, out DateTime inputDate) && !string.IsNullOrEmpty(k.IdSotr.ToString()))
            {
                var isOnVacation = Connect.context.Otpuska.Any(o =>
                    o.IdSotr == k.IdSotr &&
                    inputDate >= o.DataNachalaOtpuska &&
                    inputDate <= o.DataOkonchaniyaOtpuska);

                if (isOnVacation)
                {
                    StatusComboBox.SelectedItem = Connect.context.Status.FirstOrDefault(s => s.IdStatus == 4);
                    StatusComboBox.IsEnabled = false;
                }
                else
                {
                    StatusComboBox.IsEnabled = true;
                }
            }
        }
    }
}
