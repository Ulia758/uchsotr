using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
    /// Логика взаимодействия для AddSotrudniki.xaml
    /// </summary>
    public partial class AddSotrudniki : Page
    {
        Sotrudniki k;
        public string Phone { get; set; } = "";
        public AddSotrudniki(Sotrudniki c)
        {
            InitializeComponent();
            if (c == null)
                c = new Sotrudniki();
            DataContext = k = c;
            DolgComboBox.ItemsSource = Connect.context.Dolgnosti.ToList();
        }
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (k.IdSotr == 0)
            {
                Connect.context.Sotrudniki.Add(k);
            }
            try
            {
                // Сохраняем изменения
                Connect.context.SaveChanges();

                // Создаем запись в таблице Zarchislenie
                var zachislenieRecord = new Zachislenie()
                {
                    IdSotr = k.IdSotr,
                    DateZach = DateTime.Now.Date // Используем сегодняшнюю дату
                };

                // Добавляем новую запись в контекст
                Connect.context.Zachislenie.Add(zachislenieRecord);

                // Сохраняем изменения снова
                Connect.context.SaveChanges();

                MessageBox.Show("Данные успешно сохранены.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (DbEntityValidationException ex)
            {
                foreach (var error in ex.EntityValidationErrors)
                {
                    foreach (var validationError in error.ValidationErrors)
                    {
                        MessageBox.Show($"Ошибка проверки сущности типа {error.Entry.Entity.GetType().Name}. Поле: {validationError.PropertyName}, Ошибка: {validationError.ErrorMessage}");
                    }
                }
                MessageBox.Show("Возникла ошибка при сохранении данных. Обратитесь к администратору.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Nav.MainFrame.GoBack(); // Возвращаемся назад
            }
        }
        private void Back_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }    
    }
}
