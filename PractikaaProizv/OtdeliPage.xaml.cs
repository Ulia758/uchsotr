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
namespace PractikaaProizv {
    /// <summary>
    /// Логика взаимодействия для OtdeliPage.xaml
    /// </summary>
    public partial class OtdeliPage : Page
    {
        public OtdeliPage()
        {
            InitializeComponent();
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddOtdeli(null));
        }
        private void EditBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddOtdeli((sender as Button).DataContext as Otdeli));
        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            var delClients = OtdelDG.SelectedItems.Cast<Otdeli>().ToList();
            foreach (var delClient in delClients)
                if (Connect.context.Dolgnosti.Any(x => x.OtdelId == delClient.IdOtdel))
                {
                    MessageBox.Show("Данные используются в другой таблице", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            if (MessageBox.Show($"Удалить {delClients.Count} записей", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                Connect.context.Otdeli.RemoveRange(delClients);
            }
            try
            {
                Connect.context.SaveChanges();
                OtdelDG.ItemsSource = Connect.context.Otdeli.ToList();
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
            OtdelDG.ItemsSource = Connect.context.Otdeli.Where(x => x.Nazvanie.StartsWith(Poisk.Text)).ToList();
        }
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            OtdelDG.ItemsSource = Connect.context.Otdeli.ToList();
        }
    }
}
