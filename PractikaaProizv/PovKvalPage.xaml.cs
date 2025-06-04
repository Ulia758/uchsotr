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

namespace PractikaaProizv
{
    /// <summary>
    /// Логика взаимодействия для PovKvalPage.xaml
    /// </summary>
    public partial class PovKvalPage : Page
    {
        public PovKvalPage()
        {
            InitializeComponent();
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddPovKval(null));
        }
        private void EditBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddPovKval((sender as Button).DataContext as PovishemieKvalific));
        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            var delClients = KvalifDG.SelectedItems.Cast<PovishemieKvalific>().ToList();
            if (MessageBox.Show($"Удалить {delClients.Count} записей", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                Connect.context.PovishemieKvalific.RemoveRange(delClients);
            }
            try
            {
                Connect.context.SaveChanges();
                KvalifDG.ItemsSource = Connect.context.PovishemieKvalific.ToList();
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
        private void Poisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            KvalifDG.ItemsSource = Connect.context.PovishemieKvalific.Where(x => x.Kurs.StartsWith(Poisk.Text)).ToList();
        }
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            KvalifDG.ItemsSource = Connect.context.PovishemieKvalific.ToList();
        }
    }
}
